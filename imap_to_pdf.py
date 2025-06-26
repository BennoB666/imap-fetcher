
import imaplib
import email
import os
import time
import json
from fpdf import FPDF
from datetime import datetime

def load_config():
    with open("config.json", "r") as f:
        return json.load(f)

def save_email_as_pdf(subject, sender, body, save_folder):
    try:
        print("PDF wird erstellt...")
        pdf = FPDF()
        pdf.add_page()
        pdf.add_font("DejaVu", "", "DejaVuSans.ttf", uni=True)
        pdf.set_font("DejaVu", size=12)
        pdf.multi_cell(0, 10, f"From: {sender}")
        pdf.multi_cell(0, 10, f"Subject: {subject}")
        pdf.multi_cell(0, 10, "Body:\n" + body)

        safe_subject = "".join(c if c.isalnum() else "_" for c in subject)[:50]
        filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe_subject}.pdf"
        filepath = os.path.join(save_folder, filename)
        pdf.output(filepath)
        print(f"Gespeichert als: {filepath}")
    except Exception as e:
        print("Fehler beim Erstellen der PDF:", e)


def process_mailbox(mail, save_folder):
    status, messages = mail.search(None, "UNSEEN")
    print("Suchstatus:", status)
    print("UNSEEN-Nachrichten:", messages)

    if status != "OK":
        print("Keine neuen Nachrichten.")
        return

    for num in messages[0].split():
        print(f"Verarbeite Nachricht Nr. {num.decode()}")
        status, data = mail.fetch(num, "(RFC822)")
        if status != "OK":
            print("Fehler beim Abrufen der Nachricht.")
            continue

        msg = email.message_from_bytes(data[0][1])
        subject = msg.get("Subject", "Kein Betreff")
        sender = msg.get("From", "Unbekannt")

        print("Absender:", sender)
        print("Betreff:", subject)

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        print("Body-Auszug:", body[:200])  # nur ein Ausschnitt
        save_email_as_pdf(subject, sender, body, save_folder)

def main():
    config = load_config()
    while True:
        try:
            mail = imaplib.IMAP4_SSL(config["imap_server"])
            mail.login(config["email_account"], config["email_password"])
            mail.select(config["mailbox"])
            if not os.path.exists(config["save_folder"]):
                os.makedirs(config["save_folder"])
            process_mailbox(mail, config["save_folder"])
            mail.logout()
        except Exception as e:
            print(f"Fehler: {e}")
        time.sleep(config["check_interval"])

if __name__ == "__main__":
    main()
