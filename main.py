from pandas import DataFrame, ExcelWriter
from os import listdir, mkdir
from os.path import isdir
from re import findall
from xml.etree import ElementTree
from PyPDF2 import PdfReader

# Default header texts for generated sheets
df_description = "Omschrijving"
df_unit = "Eenheid"
df_name = "Op naam"
df_amount = "Aantal"
df_price = "Prijs (€)"

# Default folder names to read invoices from / write sheets to
folder_invoices = "Invoices"
folder_sheets = "Sheets"


def prepare_invoice_folders(folders: list) -> None:
    for folder_dir in folders:
        if not isdir(folder_dir):
            mkdir(folder_dir)


def get_int_input(min_value: int, max_value: int, message: str="> ") -> int:
    while True:
        n = input(message)
        if n.isnumeric():
            n = int(n)
            if min_value <= n <= max_value:
                return n
            print(f"Please enter a number between {min_value} and {max_value}")
        else:
            print("Please enter a number.")


def get_invoice_filename(message: str) -> str:
    files = list(file for file in listdir(folder_invoices) if file.lower().endswith(".xml"))

    print(message)
    for i, choice in enumerate(files):
        print(f"[{i}] - {choice}")
    index = get_int_input(0, max(len(files) - 1, 0))
    return files[index][:-4]


def read_invoice_xml(path: str) -> DataFrame:
    invoices_root = ElementTree.parse(path).getroot()
    invoices = list(node for node in invoices_root if node.tag.endswith("InvoiceLine"))
    invoices_df = DataFrame(index=range(len(invoices)), columns=[df_description, df_unit, df_name, df_amount, df_price])

    for invoice, i in zip(invoices, range(len(invoices))):
        quantity = float(invoice[2].text)
        total_excl_tax = float(invoice[3].text)
        tax = float(invoice[6][0].text)
        price_incl_tax = round((total_excl_tax + tax) / quantity, 2)

        invoices_df.loc[i, df_description] = invoice[7][1].text
        invoices_df.loc[i, df_name] = invoice[1].text
        invoices_df.loc[i, df_amount] = quantity
        invoices_df.loc[i, df_price] = price_incl_tax
    return invoices_df


def read_invoice_pdf(path: str) -> list:
    start_sentence = "BTW bedrag Prijs\n% "
    start_len = len(start_sentence)
    end_sentence_1 = "Totaal exclusief BTW"
    end_sentence_2 = "Delftsche Studenten Bond"

    reader = PdfReader(path)
    pdf_size = len(reader.pages)
    orders = list()

    for i in range(pdf_size):
        text = reader.pages[i].extract_text()
        start = text.find(start_sentence) + start_len
        if i == pdf_size - 1:
            end = text.find(end_sentence_1)
        else:
            end = text.find(end_sentence_2)
        orders.extend(order.replace("\n", "") for order in text[start:end].split("\n% "))
    return orders


def edit_units_row(invoices: DataFrame, orders: list) -> None:
    for order, description, names, i in zip(orders, invoices[df_description], invoices[df_name], range(len(orders))):
        if type(names) is str:
            match = names
        else:
            match = findall("-?[0-9]+ € ", order)[0]
            # match = findall("-?[0-9]+,[0-9][0-9] € ", order)[0]
        start = order.find(description) + len(description)
        end = order.find(match)
        invoices.loc[i, df_unit] = order[start:end].strip()


def main():
    prepare_invoice_folders([folder_sheets, folder_invoices])

    while True:
        filename = get_invoice_filename("Choose one of the xml files by its index:")
        invoices = read_invoice_xml("{}\\{}.XML".format(folder_invoices, filename))
        orders = read_invoice_pdf("{}\\{}.PDF".format(folder_invoices, filename))
        edit_units_row(invoices, orders)
        writer = ExcelWriter("{}\\{}.xlsx".format(folder_sheets, filename))
        invoices.to_excel(writer)
        sheet = writer.sheets["Sheet1"]
        for column in ["B", "C", "D"]:
            sheet.column_dimensions[column].width = 20
        writer.close()

        if input("Enter 'exit' to stop. Enter anything else to continue.\n> ") == "exit":
            break


if __name__ == '__main__':
    main()
