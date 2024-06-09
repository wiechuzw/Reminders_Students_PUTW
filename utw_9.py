import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk  # Pillow for handling images
import openpyxl
import os
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from utw_12_add import FROM_PASSWORD
from typing import List, Tuple, Optional

EMAIL_BROADCASTING = "wieslaw.ziewiecki@gmail.com"
SUBJECT = "PUTW"
BODY = "(Niniejsza wiadomość została wygenerowana automatycznie - proszę na nią nie odpowiadać).\nZanotowaliśmy zaległości w płatnościach do UTW w kwocie: "
EXCEL_FILE_PATH = r'C:\Users\Komputer\OneDrive\Pulpit\UTW_2.xlsx'
SHEET_NAME = 'BILANS NALEŻNOŚCI'
ROW_FROM = 7
ROW_TO = 19

def date_() -> str:
    now = datetime.now()
    current_date = now.strftime("%Y-%m-%d")
    return current_date

def read_excel(file_path: str, sheet_name: str) -> Optional[List[Tuple[Optional[str], Optional[str], Optional[float], Optional[str]]]]:
    if not os.path.isfile(file_path):
        print(f"Plik nie istnieje: {file_path}")
        return None

    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            data = []
            for row in range(ROW_FROM, ROW_TO):
                surname_c = sheet[f'C{row}'].value
                name_d = sheet[f'D{row}'].value
                value_j = sheet[f'J{row}'].value
                email_g = sheet[f'G{row}'].value
                data.append((surname_c, name_d, value_j, email_g))
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
    print("Rozpoczęcie funkcji main()")
    data = read_excel(EXCEL_FILE_PATH, SHEET_NAME)
    
    if data:
        for index, (surname_c, name_d, value_j, email_g) in enumerate(data, start=6):
            if value_j is not None:
                try:
                    value_j = str(value_j).replace(",", "").strip()
                    value_j = float(value_j)
                    print(f"{surname_c} {name_d} - bilans wpłat: {value_j}", end=' ')
                    if value_j < 0:
                        send_email(email_g, SUBJECT, BODY + f"{value_j} zł" + " " + f"wg stanu na dzień {date_()}", EMAIL_BROADCASTING, FROM_PASSWORD)
                    else:
                        print(f" -> Brak zaległości lub wartość w komórce J{index} jest nieprawidłowa.")
                except ValueError:
                    print(f"Wartość w komórce J{index} nie jest liczbą.")
            else:
                print(f"Nie udało się odczytać wartości z komórki J{index}.")
    else:
        print("Nie udało się odczytać danych z pliku Excel.")

    print("Zakończenie funkcji main()")

def on_button_click():
    try:
        print("Przycisk został kliknięty")
        main()
        messagebox.showinfo("Informacja", "Program zakończył działanie pomyślnie.")
    except Exception as e:
        messagebox.showerror("Błąd", f"Wystąpił błąd: {e}")

def create_gui():
    root = tk.Tk()
    root.title("Reminders for Students PUTW - Created by Wiesław Ziewiecki")
    root.geometry("593x711")  # Adjust the window size (width increased by 15%, height decreased by 5%)

    # Display an image
    image = Image.open(r"F:\PYTHON\UTW_3\Reminders_Students_PUTW\motyl_.png")  # Replace with your image file
    photo = ImageTk.PhotoImage(image)
    image_label = tk.Label(root, image=photo)
    image_label.image = photo  # Keep a reference to avoid garbage collection
    image_label.pack(pady=10)

    # Instruction label
    instruction_label = tk.Label(root, text="Naciśnij przycisk poniżej, aby uruchomić program.", font=("Arial", 14))
    instruction_label.pack(pady=10)

    # Title label with larger font
    # title_label = tk.Label(root, text="Program Wspomagający Pracę Księgowej", font=("Arial", 24, "bold"))
    # title_label.pack(pady=20)
    # title_label.place(relx=0.5, rely=0.1, anchor='center')  # Center the title

    # Round button
    button_canvas = tk.Canvas(root, width=100, height=100, bg="white", highlightthickness=0)  # Reduced size to 50%
    button_canvas.pack(pady=20)
    button_canvas.create_oval(10, 10, 90, 90, fill="green")  # Adjusted coordinates for smaller size

    button_text = button_canvas.create_text(50, 50, text="START", fill="white", font=("Arial", 12, "bold"))

    def on_canvas_click(event):
        on_button_click()

    button_canvas.bind("<Button-1>", on_canvas_click)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
