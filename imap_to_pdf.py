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
    pdf = FPDF()
    pdf.add_page()
    # Unicode-Schriftart (DejaVuSans.ttf muss im Script-Ordner liegen)
    pdf.add_font("DejaVu", "", "DejaVuSans.ttf", uni=True)
    pdf.set_font("DejaVu", size=12)
    pdf.multi_cell(0, 10, f"From: {sender}")
    pdf.multi_cell(0, 10, f"Subject: {subject}")
    pdf.multi_cell(0, 10, "Body:\n" + body)
    safe = "".join(c if c.isalnum() else "_" for c in subject)[:50]
    fn = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe}.pdf"
    path = os.path.join(save_folder, fn)
    pdf.output(path)
    print(f"PDF gespeichert: {path}")

def save_attachment(part, save_folder):
    filename = part.get_filename()
    if not filename:
        return
    # sauberes Dateinamen‑Sanitizing
    safe = "".join(c if c.isalnum() or c in "._-" else "_" for c in filename)
    path = os.path.join(save_folder, safe)
    with open(path, "wb") as fp:
        fp.write(part.get_payload(decode=True))
    print(f"Anhang gespeichert: {path}")

def process_mailbox(mail, save_folder):
    status, msgs = mail.search(None, "UNSEEN")
    if status != "OK" or not msgs[0]:
        print("Keine neuen Nachrichten.")
        return
    for num in msgs[0].split():
        status, data = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(data[0][1])
        subject = msg.get("Subject", "(kein Betreff)")
        sender = msg.get("From", "(unbekannt)")
        body = ""
        # Durch alle Parts iterieren
        for part in msg.walk():
            cdisp = part.get_content_disposition()
            ctype = part.get_content_type()
            if ctype == "text/plain" and cdisp is None:
                body = part.get_payload(decode=True).decode(errors="ignore")
            elif cdisp == "attachment":
                save_attachment(part, save_folder)
        save_email_as_pdf(subject, sender, body, save_folder)

def main():
    cfg = load_config()
    while True:
        try:
            m = imaplib.IMAP4_SSL(cfg["imap_server"])
            m.login(cfg["email_account"], cfg["email_password"])
            m.select(cfg["mailbox"])
            os.makedirs(cfg["save_folder"], exist_ok=True)
            process_mailbox(m, cfg["save_folder"])
            m.logout()
        except Exception as e:
            print("Fehler:", e)
        time.sleep(cfg["check_interval"])

if __name__=="__main__":
    main()
