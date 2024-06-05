"""
Program operation:
Reading data from Excel: The program first reads data from the Excel file, checking whether the file exists and whether a sheet with the given name is available.
Data Processing: Iterates through rows 6 through 16, reading the appropriate values ​​from the columns and creating a list of data.
Emailing: For each tuple in the list, checks whether the value in column J is negative. If so, it sends an email to the appropriate email address.
Displaying information: The program displays diagnostic information such as names, deposit balances, and email success or error messages.
"""


import openpyxl  # Importuje moduł openpyxl, który umożliwia odczyt i zapis plików Excel
import os  # Importuje moduł os, który umożliwia interakcję z systemem operacyjnym, np. sprawdzanie, czy plik istnieje
import smtplib   # Importuje moduł smtplib, który umożliwia wysyłanie e-maili za pomocą protokołu SMTP
from datetime import datetime
from email.mime.text import MIMEText  # Importuje klasę MIMEText do tworzenia wiadomości e-mail w formacie MIME
from utw_12_add import FROM_PASSWORD
from typing import List, Tuple, Optional

EMAIL_BROADCASTING = "wieslaw.ziewiecki@gmail.com"
SUBJECT = "PUTW"
BODY = "( Niniejsza wiadomość została wygenerowana automatycznie - proszę na nią nie odpowiadać).\nZanotowaliśmy zaległości w płatnościach do UTW w kwocie: "
# BODY = "(Płatności - O.K. Pozdrawiam - Skarbnik "
EXCEL_FILE_PATH = r'C:\Users\Komputer\OneDrive\Pulpit\UTW_2.xlsx'
SHEET_NAME = 'BILANS NALEŻNOŚCI'
ROW_FROM = 7
ROW_TO = 19


def date_() -> str:
    now = datetime.now()
    current_date = now.strftime("%Y-%m-%d")
    return current_date


def read_excel(file_path: str, sheet_name: str) -> Optional[List[Tuple[Optional[str], Optional[str], Optional[float], Optional[str]]]]:
    """
    Function to read values from an Excel file
    """
    if not os.path.isfile(file_path):
        print(f"Plik nie istnieje: {file_path}")
        return None

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)  # data_only=True zwraca wartości zamiast formuł
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            data = []
            for row in range(ROW_FROM, ROW_TO): # Iteracja przez wiersze od 6 do 16, odczytuje wartości z kolumn C, D, J i G.
                surname_c = sheet[f'C{row}'].value
                name_d = sheet[f'D{row}'].value
                value_j = sheet[f'J{row}'].value
                email_g = sheet[f'G{row}'].value
                data.append((surname_c, name_d, value_j, email_g))  # lista krotek
            return data 
        else:
            print(f"Arkusz o nazwie {sheet_name} nie istnieje.")
            return None
    except PermissionError as e:
        print(f"Permission error: {e}")
        return None
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


def send_email(to_email: str, subject: str, body: str, from_email: str, from_password: str) -> None:
    """ 
    Function to send an e-mail
    """
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = to_email

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(from_email, from_password)
            server.sendmail(from_email, to_email, msg.as_string())
        print(f" -> E-mail wysłany do {to_email}")
    except smtplib.SMTPException as e:
        print(f"SMTP error occurred: {e}")
    except Exception as e:
        print(f"Error sending email: {e}")


def main():
    data = read_excel(EXCEL_FILE_PATH, SHEET_NAME)
    
    if data:
        for index, (surname_c, name_d, value_j, email_g) in enumerate(data, start=6):
            if value_j is not None:
                try:
                    value_j = str(value_j).replace(",", "").strip()  # Czyszczenie wartości
                    value_j = float(value_j)  # Konwersja na liczbę zmiennoprzecinkową
                    print(f"{surname_c} {name_d} - bilans wpłat: {value_j}", end=' ')
                    if value_j < 0:
                        send_email(email_g, SUBJECT, BODY + f"{value_j} zł" + " " + f"{"  wg stanu na dzień  "}" + " " + date_(), EMAIL_BROADCASTING, FROM_PASSWORD)
                    else:
                        print(f" -> Brak zaległości lub wartość w komórce J{index} jest nieprawidłowa.")
                except ValueError:
                    print(f"Wartość w komórce J{index} nie jest liczbą.")
            else:
                print(f"Nie udało się odczytać wartości z komórki J{index}.")
    else:
        print("Nie udało się odczytać danych z pliku Excel.")


if __name__ == "__main__":
    main()