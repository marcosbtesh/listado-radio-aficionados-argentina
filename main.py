import json
import csv
import re
import os
import socket
import ssl
import tempfile
import urllib.parse
import urllib.request

import certifi
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

URL = "https://hertz.enacom.gob.ar/se/portal/arg/publico/ListadoRadioaficionado.php"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}

TABLE_FORMAT = [
    "Radioaficionado",
    "Categoria",
    "Señal Distintiva",
    "Vigencia Certificado Antecedentes Penales",
    "Ciudad",
    "Provincia",
]

OUTPUT_DIR = "output"
PARSED_LIST_FILE = "parsed_list.txt"

# hertz.enacom.gob.ar serves only its leaf certificate and omits the Sectigo
# intermediate that signs it, so OpenSSL cannot build a path to a trusted root
# ("unable to get local issuer certificate"). Browsers hide this by downloading
# the missing issuer from the certificate's Authority Information Access (AIA)
# extension; the helpers below do the same and hand requests an extended CA
# bundle. Certificate verification stays fully enabled: a downloaded issuer is
# only useful if it is itself signed by a root already trusted by certifi.

# DER encoding of the id-ad-caIssuers OID (1.3.6.1.5.5.7.48.2) that prefixes the
# issuer URL inside the AIA extension.
CA_ISSUERS_OID_DER = b"\x06\x08\x2b\x06\x01\x05\x05\x07\x30\x02"
MAX_AIA_HOPS = 4


def _download_peer_certificate(host, port=443):
    """Return the DER leaf certificate the server presents, without validating it.

    Nothing is trusted based on this read; it is only used to discover the AIA
    URL. The chain is verified for real on the subsequent requests call.
    """
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE

    with socket.create_connection((host, port), timeout=30) as raw_socket:
        with context.wrap_socket(raw_socket, server_hostname=host) as tls_socket:
            return tls_socket.getpeercert(binary_form=True)


def _extract_ca_issuers_url(certificate_der):
    """Pull the caIssuers URL out of a DER certificate's AIA extension."""
    offset = certificate_der.find(CA_ISSUERS_OID_DER)
    if offset == -1:
        return None

    # The accessLocation follows the OID as a GeneralName tagged [6] (0x86),
    # holding the URL as an IA5String in definite short form.
    position = offset + len(CA_ISSUERS_OID_DER)
    if position + 2 > len(certificate_der) or certificate_der[position] != 0x86:
        return None

    length = certificate_der[position + 1]
    url = certificate_der[position + 2 : position + 2 + length]
    if len(url) != length:
        return None

    url = url.decode("ascii", "ignore")
    # AIA is fetched over plain HTTP by design (RFC 5280); reject anything else
    # so a malformed extension cannot point us at an unexpected scheme.
    if urllib.parse.urlparse(url).scheme not in ("http", "https"):
        return None

    return url


def _to_pem(certificate_bytes):
    """Normalise a downloaded certificate to PEM, accepting DER or PEM input.

    Returns None when the download is not a single X.509 certificate. AIA URLs
    routinely serve PKCS#7 (.p7c) bundles, and DER_cert_to_PEM_cert will happily
    base64 those into a CERTIFICATE block that OpenSSL then rejects, poisoning
    the whole CA bundle - so every candidate is parsed before it is accepted.
    """
    if certificate_bytes.lstrip().startswith(b"-----BEGIN"):
        pem = certificate_bytes.decode("ascii", "ignore")
    else:
        pem = ssl.DER_cert_to_PEM_cert(certificate_bytes)

    try:
        ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT).load_verify_locations(cadata=pem)
    except ssl.SSLError:
        return None

    return pem


def _ca_bundle_path(host):
    return os.path.join(tempfile.gettempdir(), f"{host}-ca-bundle.pem")


def build_ca_bundle(host):
    """Write certifi's roots plus the issuers the server omits to a bundle file."""
    certificate_der = _download_peer_certificate(host)
    issuer_pems = []
    seen_urls = set()

    for _ in range(MAX_AIA_HOPS):
        url = _extract_ca_issuers_url(certificate_der)
        if not url or url in seen_urls:
            break

        seen_urls.add(url)
        with urllib.request.urlopen(url, timeout=30) as response:
            certificate_bytes = response.read()

        issuer_pem = _to_pem(certificate_bytes)
        if issuer_pem is None:
            # Not a bare certificate (usually a PKCS#7 bundle); the issuers
            # collected so far are normally enough to reach a certifi root.
            break

        issuer_pems.append(issuer_pem)
        certificate_der = certificate_bytes

    if not issuer_pems:
        raise RuntimeError(
            f"{host} presented an incomplete certificate chain and published no "
            "caIssuers URL, so the missing issuer could not be recovered."
        )

    bundle_path = _ca_bundle_path(host)
    with open(bundle_path, "w", encoding="ascii") as bundle:
        bundle.write(certifi.contents())
        bundle.write("\n")
        bundle.write("\n".join(issuer_pems))

    return bundle_path


def repair_tls_trust(session, host):
    """Point the session at a CA bundle that can complete this host's chain.

    Reuses a previously built bundle when one is cached, and rebuilds from
    scratch if that cached copy has gone stale.
    """
    cached_bundle = _ca_bundle_path(host)
    if os.path.exists(cached_bundle):
        session.verify = cached_bundle
        return

    session.verify = build_ca_bundle(host)


def main():
    print("Starting...")

    generate_output_folder()
    parsed_list = []
    if os.path.exists(PARSED_LIST_FILE):
        parsed_list = load_txt_file_as_list()
    else:
        full_list = fetch_full_list()
        parsed_list = parse_html_list(full_list)

    # Generate Files

    # generateExcel(parse_html_list(fetch_full_list()))

    generateExcel(parsed_list)


def load_txt_file_as_list():
    parsed_list = []
    with open(PARSED_LIST_FILE, "r") as pl:
        list = pl.read()

        for ro in list.split("\n\n"):
            parsed_list.append(ro)

    return parsed_list


def fetch_full_list():
    with requests.Session() as s:
        s.headers.update(HEADERS)

        response = get_with_tls_repair(s, URL)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")
        token_input = soup.find("input", {"name": "csrf_token"})

        if not token_input or not token_input.get("value"):
            raise RuntimeError("csrf_token not found on website.")

        csrf = token_input["value"]

        data = {"valor": "", "csrf_token": csrf, "mostrarTodos": 1}

        response = s.post(URL, data, timeout=60)
        response.raise_for_status()

        return response.text


def get_with_tls_repair(session, url):
    """GET a URL, recovering once from a chain the server failed to send in full."""
    host = urllib.parse.urlparse(url).hostname

    try:
        return session.get(url, timeout=30)
    except requests.exceptions.SSLError:
        print(f"TLS verification failed for {host}: incomplete certificate chain.")
        print("Recovering the missing issuer certificate...")

    repair_tls_trust(session, host)

    try:
        return session.get(url, timeout=30)
    except requests.exceptions.SSLError:
        # The cached bundle was stale (rotated or expired issuer) - rebuild it.
        session.verify = build_ca_bundle(host)
        return session.get(url, timeout=30)


def parse_html_list(response):
    soup = BeautifulSoup(response, "html.parser")
    radio_operators_list = soup.find_all("table", {"class": "listado"})[0].find_all(
        "tr"
    )[1:]
    parsed_list = []

    print(radio_operators_list[0:2])
    for radio_operator in radio_operators_list:
        try:
            parsed_list.append(radio_operator.get_text())
            # with open("parsed_list.txt", "w+") as pl:
            # pl.write(radio_operator.get_text())
            # pl.write("\n")
            # print(f" {radio_operator.index()} - {radio_operator.get_text()}")
        except:
            continue

    print(parsed_list[0:2])

    with open(PARSED_LIST_FILE, "w+") as pl:
        pl.writelines(parsed_list)
    return parsed_list


def generateExcel(parsed_list):
    wb = Workbook()
    ws = wb.active

    ws.append(TABLE_FORMAT)

    # print(parsed_list[0:2])
    # for column in TABLE_FORMAT:
    #     print(f"Column in table format - {column}")

    for radio_afficionado in parsed_list:
        # print(f"Radio Afficionado - {radio_afficionado}")
        ws.append(radio_afficionado.split("\n"))
        # for radio_afficionado_data in radio_afficionado.split("\n"):
        #     print(f"Data - {radio_afficionado_data}")
        #     (ws.append(radio_afficionado_data))
    wb.save(f"{OUTPUT_DIR}/listado.xlsx")


# def generate_csv(parsed_list):


def generate_output_folder():
    if not os.path.exists(OUTPUT_DIR):
        os.mkdir(OUTPUT_DIR)


if __name__ == "__main__":
    main()
