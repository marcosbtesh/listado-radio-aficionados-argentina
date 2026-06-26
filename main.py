import csv
import re

import requests
from bs4 import BeautifulSoup

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


def main():
    print("Starting...")


def fetch_full_list():
    with requests.Session() as s:
        s.headers.update(HEADERS)

        response = s.get(URL, timeout=30)
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


def parse_html_list(response):
    soup = BeautifulSoup(response, "html.parser")
    radio_operators_list = soup.find_all("table", {"class": "listado"})[0].find_all("tr")[1:]
    parsed_list = []
    for radio_operator in radio_operators_list:
        try: 
            parsed_list.append(radio_operator.get_text())
        except:
            continue
        
        



def generate_csv(parsed_list):





if __name__ == "__main__":
    main()
