import click
import pandas as pd
import pymongo

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


def capitalize_each_word_in_string(string: "str", delimiter: "str" = " ", is_upper_case: "bool" = False) -> "str":
    words: "list[str]" = string.split(delimiter)
    if is_upper_case:
        capitalized_words: "list[str]" = [
            word.lower().capitalize() for word in words
        ]
    else:
        capitalized_words: "list[str]" = [
            word.capitalize() for word in words
        ]

    capitalized_string: "str" = " ".join(capitalized_words)
    return capitalized_string


@click.command()
@click.argument("filename", type=click.Path(exists=True))
def save_guides_into_database(filename):
    df = pd.read_excel(filename)
    # df = df[0:len(df) - 1]
    df = df[df["NumeroGuia"].str[1:3] == "DG"]
    print(df)
    print(df.index)

    for guide in range(len(df)):
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
        sender = capitalize_each_word_in_string(sender, "-", True)
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

        guide_to_save = {
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
    try:
        guide_id = guides_collection.insert_one(guide_to_save).inserted_id
    except pymongo.errors.DuplicateKeyError:
        print("The guide already exists in the database.")
