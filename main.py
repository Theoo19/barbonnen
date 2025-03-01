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

# Sentences used to read the PDF
pdf_start = "BTW bedrag Prijs\n% "
pdf_end_1 = "Delftsche Studenten Bond"
pdf_end_2 = "Totaal exclusief BTW"

# Regex patterns used to read the PDF (change this in edit_units_row() function depending on PDF formatting)
pdf_pattern_1 = "-?[0-9]+ € "
pdf_pattern_2 = "-?[0-9]+,[0-9][0-9] € "


def prepare_invoice_folders(folders: list) -> None:
    # Create folders if they do not yet exist
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


def get_invoice_filenames() -> list:
    # Create list of xml invoice files
    return list(file for file in listdir(folder_invoices) if file.lower().endswith(".xml"))


def get_invoice_choice(files) -> str:
    # Make the user choose one of the possible xml files to process
    print("Choose one of the xml files by its index:")
    for i, choice in enumerate(files):
        print(f"[{i}] - {choice}")
    index = get_int_input(0, max(len(files) - 1, 0))
    return files[index][:-4]


def read_invoice_xml(path: str) -> DataFrame:
    # Read the XML invoice and extract the invoice entries into a DataFrame

    # Read the XML invoice
    invoices_root = ElementTree.parse(path).getroot()
    invoices = list(node for node in invoices_root if node.tag.endswith("InvoiceLine"))

    # Create a DataFrame with a row for each InvoiceLine
    invoices_df = DataFrame(index=range(len(invoices)), columns=[df_description, df_unit, df_name, df_amount, df_price])

    # Loop through each invoice, read the information and store it in a DataFrame row
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
    # Read the PDF invoice and extract the invoice entries into a list

    # Read the PDF invoice
    reader = PdfReader(path)
    pdf_size = len(reader.pages)
    orders = list()

    # Loop through each page and filter out the invoice entries using start and end patterns
    pdf_end = pdf_end_1
    for i in range(pdf_size):
        if i == pdf_size - 1:
            pdf_end = pdf_end_2
        text = reader.pages[i].extract_text()
        start = text.find(pdf_start) + len(pdf_start)
        end = text.find(pdf_end)
        orders.extend(order.replace("\n", "") for order in text[start:end].split("\n% "))
    return orders


def edit_units_row(invoices: DataFrame, orders: list) -> None:
    # Use the data from the XML invoice to filter the "Eenheid" part of the invoice entries of the PDF invoice
    for order, description, names, i in zip(orders, invoices[df_description], invoices[df_name], range(len(orders))):
        if type(names) is str:
            match = names
        else:
            match = findall(pdf_pattern_1, order)[0] # Change to pattern_1 or pattern_2 depending on PDF formatting
        start = order.find(description) + len(description)
        end = order.find(match)
        invoices.loc[i, df_unit] = order[start:end].strip()


def main():
    prepare_invoice_folders([folder_sheets, folder_invoices])

    while True:
        filenames = get_invoice_filenames()
        if len(filenames) == 0:
            input(f"Error: folder '{folder_invoices}' is empty. Press enter to exit.\n> ")
            break

        filename = get_invoice_choice(filenames)
        invoices = read_invoice_xml(f"{folder_invoices}\\{filename}.XML")
        orders = read_invoice_pdf(f"{folder_invoices}\\{filename}.PDF")
        edit_units_row(invoices, orders)
        writer = ExcelWriter(f"{folder_sheets}\\{filename}.xlsx")
        invoices.to_excel(writer)
        sheet = writer.sheets["Sheet1"]
        for column in ["B", "C", "D"]:
            sheet.column_dimensions[column].width = 20
        writer.close()

        if input("Enter 'exit' to stop. Enter anything else to continue.\n> ") == "exit":
            break


if __name__ == '__main__':
    main()
