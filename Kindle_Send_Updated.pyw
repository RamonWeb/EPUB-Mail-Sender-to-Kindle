import os
import smtplib
import json
from tkinter import *
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

CONFIG_FILE = "config.json"


def load_config():
    """Lädt die Konfigurationsdatei."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as file:
            return json.load(file)
    return {
        "source_folder": "",
        "sent_folder": "",
        "smtp_server": "",
        "smtp_port": 587,
        "username": "",
        "password": "",
        "kindle1": "",
        "kindle2": "",
        "active_kindle": "kindle1"
    }


def save_config(config_data):
    """Speichert die Konfigurationsdatei."""
    with open(CONFIG_FILE, "w") as file:
        json.dump(config_data, file, indent=4)


config = load_config()


def select_folder(entry):
    """Öffnet einen Dialog zur Auswahl eines Ordners."""
    folder = filedialog.askdirectory()
    if folder:
        entry.delete(0, END)
        entry.insert(0, folder)


def send_email(epub_file, recipient_email):
    """Sendet eine E-Mail mit der EPUB-Datei."""
    try:
        msg = MIMEMultipart()
        msg["From"] = config["username"]
        msg["To"] = recipient_email
        msg["Subject"] = "Your Kindle Document"

        with open(epub_file, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(epub_file)}")
        msg.attach(part)

        with smtplib.SMTP(config["smtp_server"], config["smtp_port"]) as server:
            server.starttls()
            server.login(config["username"], config["password"])
            server.send_message(msg)
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Senden: {e}")


def process_files():
    """Verarbeitet die Dateien im Quellordner und sendet sie per E-Mail."""
    source_folder = config["source_folder"]
    sent_folder = config["sent_folder"]
    recipient_email = config[config["active_kindle"]]

    if not os.path.exists(source_folder):
        messagebox.showerror("Fehler", "Quellordner existiert nicht!")
        return

    files = [f for f in os.listdir(source_folder) if f.endswith(".epub")]
    if not files:
        messagebox.showinfo("Info", "Keine EPUB-Dateien gefunden.")
        return

    if not recipient_email:
        messagebox.showerror("Fehler", "Keine Kindle-E-Mail-Adresse konfiguriert!")
        return

    for file in files:
        full_path = os.path.join(source_folder, file)
        try:
            # Datei senden
            send_email(full_path, recipient_email)
            messagebox.showinfo("Erfolg", f"Datei gesendet: {file}")

            # Backup-Ordner erstellen, falls er nicht existiert
            if not os.path.exists(sent_folder):
                os.makedirs(sent_folder)

            # Zielpfad überprüfen und bei Konflikt einen neuen Namen vergeben
            new_path = os.path.join(sent_folder, file)
            if os.path.exists(new_path):
                base, ext = os.path.splitext(file)
                counter = 1
                while os.path.exists(new_path):
                    new_path = os.path.join(sent_folder, f"{base}_{counter}{ext}")
                    counter += 1

            # Datei verschieben
            os.rename(full_path, new_path)
            print(f"Datei verschoben: {full_path} -> {new_path}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Problem mit Datei {file}: {e}")
            print(f"Fehler beim Verarbeiten von {file}: {e}")


def preview_files():
    """Zeigt eine Vorschau der Dateien im Quellordner in einem separaten Fenster."""
    source_folder = config["source_folder"]

    if not os.path.exists(source_folder):
        messagebox.showerror("Fehler", "Quellordner existiert nicht!")
        return

    files = [f for f in os.listdir(source_folder) if f.endswith(".epub")]
    if not files:
        messagebox.showinfo("Info", "Keine EPUB-Dateien gefunden.")
    else:
        preview_window = Toplevel()
        preview_window.title("Vorschau der Dateien")
        #preview_window.geometry("400x300")

        Label(preview_window, text="Folgende Dateien sind bereit:", font=("Arial", 12, "bold")).pack(pady=10)

        listbox = Listbox(preview_window, width=50, height=15)
        listbox.pack(pady=10, padx=10)

        for file in files:
            listbox.insert(END, file)

        Button(preview_window, text="Schließen", command=preview_window.destroy).pack(pady=10)



def open_config_window():
    """Öffnet das Konfigurationsfenster."""
    config_window = Toplevel()
    config_window.title("Konfiguration")

    Label(config_window, text="Quellordner:").grid(row=0, column=0, padx=5, pady=5)
    source_folder_entry = Entry(config_window, width=40)
    source_folder_entry.grid(row=0, column=1, padx=5, pady=5)
    source_folder_entry.insert(0, config.get("source_folder", ""))
    Button(config_window, text="Durchsuchen", command=lambda: select_folder(source_folder_entry)).grid(row=0, column=2, padx=5, pady=5)

    Label(config_window, text="Gesendet-Ordner:").grid(row=1, column=0, padx=5, pady=5)
    sent_folder_entry = Entry(config_window, width=40)
    sent_folder_entry.grid(row=1, column=1, padx=5, pady=5)
    sent_folder_entry.insert(0, config.get("sent_folder", ""))
    Button(config_window, text="Durchsuchen", command=lambda: select_folder(sent_folder_entry)).grid(row=1, column=2, padx=5, pady=5)

    Label(config_window, text="SMTP-Server:").grid(row=2, column=0, padx=5, pady=5)
    smtp_entry = Entry(config_window, width=40)
    smtp_entry.grid(row=2, column=1, padx=5, pady=5)
    smtp_entry.insert(0, config.get("smtp_server", ""))

    Label(config_window, text="SMTP-Port:").grid(row=3, column=0, padx=5, pady=5)
    port_entry = Entry(config_window, width=40)
    port_entry.grid(row=3, column=1, padx=5, pady=5)
    port_entry.insert(0, config.get("smtp_port", ""))

    Label(config_window, text="Absender E-Mail:").grid(row=4, column=0, padx=5, pady=5)
    email_entry = Entry(config_window, width=40)
    email_entry.grid(row=4, column=1, padx=5, pady=5)
    email_entry.insert(0, config.get("username", ""))

    Label(config_window, text="Passwort:").grid(row=5, column=0, padx=5, pady=5)
    password_entry = Entry(config_window, show="*", width=40)
    password_entry.grid(row=5, column=1, padx=5, pady=5)
    password_entry.insert(0, config.get("password", ""))

    Label(config_window, text="Kindle1 E-Mail:").grid(row=6, column=0, padx=5, pady=5)
    kindle1_entry = Entry(config_window, width=40)
    kindle1_entry.grid(row=6, column=1, padx=5, pady=5)
    kindle1_entry.insert(0, config.get("kindle1", ""))

    Label(config_window, text="Kindle2 E-Mail:").grid(row=7, column=0, padx=5, pady=5)
    kindle2_entry = Entry(config_window, width=40)
    kindle2_entry.grid(row=7, column=1, padx=5, pady=5)
    kindle2_entry.insert(0, config.get("kindle2", ""))

    Label(config_window, text="Aktive Kindle-Adresse:").grid(row=8, column=0, padx=5, pady=5)
    active_kindle_var = StringVar(value=config.get("active_kindle", "kindle1"))
    Radiobutton(config_window, text="Kindle1", variable=active_kindle_var, value="kindle1").grid(row=8, column=1, sticky="w")
    Radiobutton(config_window, text="Kindle2", variable=active_kindle_var, value="kindle2").grid(row=9, column=1, sticky="w")

    def save_and_close():
        """Speichert die Konfiguration und schließt das Fenster."""
        new_config = {
            "source_folder": source_folder_entry.get(),
            "sent_folder": sent_folder_entry.get(),
            "smtp_server": smtp_entry.get(),
            "smtp_port": int(port_entry.get()),
            "username": email_entry.get(),
            "password": password_entry.get(),
            "kindle1": kindle1_entry.get(),
            "kindle2": kindle2_entry.get(),
            "active_kindle": active_kindle_var.get(),
        }
        save_config(new_config)
        global config
        config = new_config
        config_window.destroy()

    Button(config_window, text="Speichern", command=save_and_close).grid(row=10, column=0, columnspan=3, pady=10)


# GUI erstellen
root = Tk()
root.title("EPUB-Mail-Sender")
root.geometry("800x600")

# Hintergrundbild setzen
bg_image = Image.open("background.jpg")
bg_photo = ImageTk.PhotoImage(bg_image.resize((800, 600)))
background_label = Label(root, image=bg_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

# Icons laden
config_icon = PhotoImage(file="configure.png")
preview_icon = PhotoImage(file="preview.png")
start_icon = PhotoImage(file="start.png")

# Buttons nebeneinander am unteren Rand positionieren
button_frame = Frame(root, bg="white")
button_frame.pack(side="bottom", pady=20)

Button(button_frame, text="Konfiguration", image=config_icon, compound=TOP, command=open_config_window).pack(side=LEFT, padx=20)
Button(button_frame, text="Vorschau", image=preview_icon, compound=TOP, command=preview_files).pack(side=LEFT, padx=20)
Button(button_frame, text="Senden", image=start_icon, compound=TOP, command=process_files).pack(side=LEFT, padx=20)

root.mainloop()
