import json
import csv
import re
import os
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
