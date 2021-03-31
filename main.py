import click
import pandas as pd
import pymongo
from prettytable import PrettyTable
from colorama import Fore
import time

# No string due to we're using the default port and we're developing in local
client = pymongo.MongoClient()
db = client["cargo-expreso-control"]
guides_collection = db["guides"]


"""
Columns for <BusquedaGuias> spreadsheets:
    NumeroGuia
    Fecha
    Remitente
    Destinatario
    Referencia 1
    Referencia 2
    CCredito
    Estado
    Motivo
    Destino
    Recibido Por
    Fecha Recibido
    Recibido Hora
"""


def capitalize_each_word_in_string(string: "str", delimiter: "str" = " ") -> "str":
    words: "list[str]" = string.split(delimiter)

    capitalized_words: "list[str]" = [
        word.lower().capitalize() for word in words
    ]

    capitalized_string: "str" = " ".join(capitalized_words)
    return capitalized_string


def format_guide_data(guide, df):
    # Original data
    id = df.loc[guide, "NumeroGuia"]
    date = df.loc[guide, "Fecha"]
    sender = df.loc[guide, "Remitente"]
    addressee = df.loc[guide, "Destinatario"]
    reference_1 = df.loc[guide, "Referencia 1"]
    reference_2 = df.loc[guide, "Referencia 2"]
    credit_code = df.loc[guide, "CCredito"]
    status = df.loc[guide, "Estado"]
    reason = df.loc[guide, "Motivo"]
    destination = df.loc[guide, "Destino"]
    received_by = df.loc[guide, "Recibido Por"]
    received_date = df.loc[guide, "Fecha Recibido"]
    received_time = df.loc[guide, "Recibido Hora"]
    paid = False

    # Modifications
    id = df.loc[guide, "NumeroGuia"]
    date = date.strftime("%d/%m/%Y")
    sender = capitalize_each_word_in_string(sender, "-")
    addressee = capitalize_each_word_in_string(addressee)
    reference_1 = str(reference_1).lower() if type(
        reference_1) == str else ""
    reference_2 = str(reference_2).lower() if type(
        reference_2) == str else ""
    credit_code = str(int(credit_code))
    status = str(status).lower()
    reason = str(reason).lower() if type(
        reason) == str else ""
    destination = df.loc[guide, "Destino"]
    received_by = capitalize_each_word_in_string(received_by)
    received_date = received_date.strftime("%d/%m/%Y")
    received_time = df.loc[guide, "Recibido Hora"]
    paid = False

    formatted_guide = {
        "_id": id,
        "date": date,
        "sender": sender,
        "addressee": addressee,
        "reference 1": reference_1,
        "reference 2": reference_2,
        "credit code": credit_code,
        "status": status,
        "reason": reason,
        "destination": destination,
        "received by": received_by,
        "received date": received_date,
        "received time": received_time,
        "paid": paid
    }

    return formatted_guide


def save_guide_to_database(guide):
    try:
        guides_collection.insert_one(guide)
        return True
    except pymongo.errors.DuplicateKeyError:
        return False


@click.command()
@click.argument("filename", type=click.Path(exists=True))
def save_guides_to_database(filename):
    df = pd.read_excel(filename)
    df = df[0:len(df) - 1]  # To quit invalid field
    valid_df = df[
        (df["NumeroGuia"].str[1:3] != "DG")
        & (df["Motivo"] != "DEVOLUCION")
        & (df["Motivo"] != "FISCALIZACION")
    ]
    invalid_df = df[
        (df["NumeroGuia"].str[1:3] == "DG")
        | (df["Motivo"] == "DEVOLUCION")
        | (df["Motivo"] == "FISCALIZACION")
    ]

    saved_guides = 0
    guides_not_saved = 0
    print(f"{Fore.CYAN}Start saving guides:{Fore.RESET}")
    for guide in valid_df.index:
        formatted_guide = format_guide_data(guide, valid_df)
        is_guide_saved = save_guide_to_database(formatted_guide)
        print(f"Saving {formatted_guide['_id']}... ", end="")
        if is_guide_saved:
            saved_guides += 1
            print(f"{Fore.GREEN}saved{Fore.RESET}️")
        else:
            guides_not_saved += 1
            print(f"{Fore.RED}already saved{Fore.RESET}")
        time.sleep(.0025)

    print(f"{saved_guides} guides were saved and {guides_not_saved} are already saved from this document\n")

    if len(invalid_df) > 0:
        print(f"{Fore.YELLOW}These guides are invalid:")
        table = PrettyTable()
        table.field_names = ["Guide Number", "Addressee", "Reason", "Date"]
        for guide in invalid_df.index:
            formatted_guide = format_guide_data(guide, invalid_df)
            id = formatted_guide["_id"]
            addressee = formatted_guide["addressee"]
            reason = formatted_guide["reason"]
            date = formatted_guide["date"]
            table.add_row([id, addressee, reason, date])
        print(f"{Fore.YELLOW}{table}{Fore.RESET}")

    print()
    print(f"{Fore.LIGHTBLACK_EX}Analyzed document: {filename}{Fore.RESET}")


@click.command()
@click.argument("filename", type=click.Path(exists=True))
def check_guides_paid(filename):
    df = pd.read_excel(filename)
    print(df)
