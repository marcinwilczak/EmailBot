import imaplib
import email
from email.header import decode_header
from dotenv import load_dotenv
import os
import openai
import csv
import json
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import messagebox
import logging

logging.basicConfig(filename='email_bot.log', level=logging.ERROR,
                    format='%(asctime)s:%(levelname)s:%(message)s')

load_dotenv()

IMAP_HOST = os.getenv('IMAP_HOST')
IMAP_USER = os.getenv('IMAP_USER')
IMAP_PASSWORD = os.getenv('IMAP_PASSWORD')
OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
OPENAI_MODEL = os.getenv('OPENAI_MODEL')
openai.api_key = OPENAI_API_KEY

def connect_to_mailbox(imap_host, imap_user, imap_password):
    try:
        mail = imaplib.IMAP4_SSL(imap_host)
        mail.login(imap_user, imap_password)
        return mail
    except Exception as e:
        logging.error(f"Nie udało się połączyć z serwerem IMAP: {e}")
        #print(f"Nie udało się połączyć z serwerem IMAP: {e}")
        return None
def fetch_emails(mail, folder="inbox", unread_only=False):
    try:
        mail.select(folder)
        status, messages = mail.search(None, 'UNSEEN' if unread_only else 'ALL')
        email_ids = messages[0].split()
        emails = []
        for email_id in email_ids:
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    email_info = parse_email(msg)
                    if email_info:
                        emails.append(email_info)
        #print(f"Pobrano {len(emails)} maili.")
        return emails
    except Exception as e:
        logging.error(f"Nie udało się pobrać maili: {e}")
        #print(f"Nie udało się pobrać maili: {e}")
        return []
def parse_email(msg):
    try:
        subject, encoding = decode_header(msg["Subject"])[0]

        if isinstance(subject, bytes):
            subject = subject.decode(encoding if encoding else 'utf-8')

        from_ = msg.get("From")
        date = msg.get("Date")
        body, has_non_text_content, has_attachment = "", False, False

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if content_type == "text/plain" and "attachment" not in content_disposition:
                    body += part.get_payload(decode=True).decode(errors='replace')
                elif "attachment" in content_disposition:
                    has_attachment = True
                elif content_type.startswith("image/"):
                    has_non_text_content = True
        else:
            body += msg.get_payload(decode=True).decode(errors='replace')

        info = ""

        if has_attachment:
            info += "Mail zawiera załączniki."

        if has_non_text_content:
            info += "Mail zawiera obrazy lub inne niesparsowane elementy."

        email_data = {"subject": subject, "from": from_, "date": date, "body": body.strip(), "info": info.strip()}
        #print(f"Odczytano mail: {email_data}")

        return email_data

    except Exception as e:
        logging.error(f"Nie udało się sparsować wiadomości: {e}")
        #print(f"Nie udało się sparsować wiadomości: {e}")

        return None

def analyze_email_body(body):
    try:
        response = openai.ChatCompletion.create(
            model=OPENAI_MODEL,
            messages=[
                {"role": "system",
                 "content": "You are a specialized assistant that accurately extracts order information, including product names and quantities, from email bodies. You should always respond with a well-structured JSON object."},
                {"role": "user", "content": (
                    "Please analyze the following email body to extract order information."
                    "Always give your answers in Polish."
                    "The response should be in JSON format with the structure: "
                    "{\"orders\": [{\"product\": \"product_name\", \"quantity\": \"quantity\"}]}. "
                    "If no order information is found, return an empty JSON object. Here is the email body:"
                    f"\n\n{body}"
                )}
            ]
        )

        result = response['choices'][0]['message']['content']
        order_data = json.loads(result)

        if isinstance(order_data, dict) and "orders" in order_data:
            #print(f"OpenAI Response: {order_data}")
            return order_data["orders"]
        else:
            #print("Brak informacji o zamówieniach.")
            return []

    except Exception as e:
        logging.error(f"Nie udało się zdekodować odpowiedzi JSON od OpenAI: {e}")
        #print(f"Nie udało się zdekodować odpowiedzi JSON od OpenAI: {e}")

        return []


def save_to_csv(email_info, order_data, filename):
    try:
        with open(filename, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)

            for item in order_data:
                product = item.get('product')
                quantity = item.get('quantity')
                writer.writerow([email_info['date'], email_info['from'], email_info['subject'], product, quantity])

        #print(f"Dane zapisane do pliku {filename}")

    except Exception as e:
        logging.error(f"Nie udało się zapisać do pliku CSV: {e}")
        #print(f"Nie udało się zapisać do pliku CSV: {e}")


def process_email(email_info, filename):
    if email_info['body']:
        order_data = analyze_email_body(email_info['body'])

        if order_data:
            save_to_csv(email_info, order_data, filename)


def process_emails_in_background(root, unread_only_var, status_label):
    try:
        status_label.config(text="Łączenie z serwerem IMAP...")
        root.update()
        mail = connect_to_mailbox(IMAP_HOST, IMAP_USER, IMAP_PASSWORD)

        if mail:
            status_label.config(text="Połączono z serwerem IMAP. Pobieranie maili...")
            root.update()
            emails = fetch_emails(mail, unread_only=unread_only_var.get())

            if not emails:
                raise Exception("Brak maili do przetworzenia.")

            if not os.path.exists('CSV_files'):
                os.makedirs('CSV_files')

            filename = f"CSV_files/orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

            with open(filename, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(["Date", "From", "Subject", "Product", "Quantity"])

            status_label.config(text="Przetwarzanie maili...")
            root.update()

            with ThreadPoolExecutor() as executor:
                futures = [executor.submit(process_email, email_info, filename) for email_info in emails]

                for future in futures:
                    future.result()

            mail.logout()

            with open(filename, 'r', encoding='utf-8') as file:
                reader = csv.reader(file)
                rows = list(reader)

                if len(rows) <= 1:
                    file.close()
                    os.remove(filename)

                    raise Exception("Brak danych zamówień do zapisania.")

            status_label.config(text="Wszystko poszło w porządku. Plik CSV został wygenerowany.")
            root.update()
            messagebox.showinfo("Sukces",
                                f"Wszystko poszło w porządku. Plik CSV został wygenerowany pod nazwą: {filename}")
        else:
            raise Exception("Nie udało się połączyć z serwerem IMAP.")

    except Exception as e:
        logging.error(f"Error in background processing: {e}")
        #print(f"Error in background processing: {e}")

        status_label.config(text=f"Coś poszło nie tak: {e}")
        root.update()
        messagebox.showerror("Błąd", f"Coś poszło nie tak: {e}")

def start_gui():
    root = tk.Tk()
    root.title("MY FIRST BOOOT!!!")

    def on_start():
        if not OPENAI_API_KEY or len(OPENAI_API_KEY) == 0:
            messagebox.showerror("Błąd", "OPENAI_API_KEY nie jest określony.")
            return

        status_label.config(text="Łączenie z serwerem IMAP...")
        root.update()
        process_emails_in_background(root, var_unread_only, status_label)

    tk.Label(root, text="Wybierz opcję przetwarzania wiadomości:").pack(pady=10)

    var_unread_only = tk.BooleanVar(value=False)
    tk.Radiobutton(root, text="Wszystkie maile", variable=var_unread_only, value=False).pack()
    tk.Radiobutton(root, text="Tylko nieodczytane maile", variable=var_unread_only, value=True).pack()

    tk.Button(root, text="Start", command=on_start).pack(pady=20)

    status_label = tk.Label(root, text="")
    status_label.pack(pady=10)

    root.geometry("400x250")
    root.mainloop()

if __name__ == "__main__":
    start_gui()