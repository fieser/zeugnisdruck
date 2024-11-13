import os
import shutil
import json
import time
import tkinter as tk
from tkinter import messagebox, ttk
import win32print
import win32security
import subprocess
from threading import Thread
from datetime import datetime, timedelta
import csv


# Laden der Konfiguration
def load_config():
    with open('config.json', 'r') as f:
        return json.load(f)


config = load_config()
printer_path = config["printer_path"]
sumatra_path = os.path.abspath(config["sumatra_path"])
status_file_path = os.path.abspath(config["status_file_path"])
folder_path = os.path.abspath(config["folder_path"])
temp_folder = os.path.abspath(config["temp_folder"])
blacklist_file = os.path.abspath(config["blacklist_file"])
timeout_minutes = config.get("timeout_minutes", 10)


# Funktion zum Einlesen der Blacklist
def load_blacklist():
    blacklist = set()
    if os.path.exists(blacklist_file):
        with open(blacklist_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=';')
            next(reader)  # Überspringe die Kopfzeile
            for row in reader:
                if len(row) == 2:
                    blacklist.add((row[0].strip(), row[1].strip()))  # (Nachname, Vorname)
    return blacklist


blacklist = load_blacklist()


# Funktion zum Überprüfen, ob ein Schüler auf der Blacklist steht
def is_blacklisted(nachname, vorname):
    return (nachname, vorname) in blacklist


# Funktion zum Überprüfen, ob der Drucker erreichbar ist
def check_printer_availability():
    try:
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_CONNECTIONS | win32print.PRINTER_ENUM_LOCAL)
        for printer in printers:
            if printer_path in printer[2]:
                return True
        raise Exception(f"Drucker '{printer_path}' nicht in der Liste verfügbarer Drucker gefunden.")
    except Exception as e:
        print(f"Fehler beim Erreichen des Druckers: {e}")
        return False


# Funktion zum Ermitteln des Dateierstellers
def get_file_owner(file_path):
    try:
        sd = win32security.GetFileSecurity(file_path, win32security.OWNER_SECURITY_INFORMATION)
        owner_sid = sd.GetSecurityDescriptorOwner()
        owner, domain, _ = win32security.LookupAccountSid(None, owner_sid)
        return owner  # Nur den Benutzernamen zurückgeben, ohne Domain
    except Exception as e:
        print(f"Fehler beim Ermitteln des Dateibesitzers: {e}")
        return None


# Funktion zum Auswerten der PDF-Dateien im angegebenen Ordner
def analyse_pdfs(folder_path, temp_folder):
    owner_data = {}
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            file_owner = get_file_owner(file_path)

            parts = filename.split("_")
            if len(parts) >= 3:
                nachname = parts[1].strip()
                vorname = parts[2].split(".")[0].strip()
                owner_folder = os.path.join(temp_folder, file_owner)

                # Prüfen, ob die Datei bereits im Archiv vorhanden ist
                is_correction = os.path.exists(os.path.join(owner_folder, filename))

                if not file_owner:
                    continue  # Überspringe, wenn kein Besitzer gefunden wird

                # Füge den Besitzer und dessen Dateien zur Liste hinzu
                owner_data[file_owner] = owner_data.get(file_owner, {'count': 0, 'files': []})
                owner_data[file_owner]['count'] += 1
                owner_data[file_owner]['files'].append((filename, nachname, vorname, is_correction))

    return owner_data


# Funktion zum Laden des Status aus der JSON-Datei
def load_status():
    if os.path.exists(status_file_path):
        try:
            with open(status_file_path, 'r') as f:
                content = f.read().strip()
                if content:
                    return json.loads(content)
                else:
                    return {}
        except json.JSONDecodeError:
            print(f"Fehler beim Laden der Statusdatei. Die Datei {status_file_path} ist möglicherweise beschädigt.")
            return {}
    return {}


# Funktion zum Speichern des Status in die JSON-Datei
def save_status(status):
    with open(status_file_path, 'w') as f:
        json.dump(status, f, indent=4)


# Funktion zur Prüfung, ob ein Besitzer aktuell blockiert ist
def is_owner_blocked(owner, status):
    current_time = datetime.now()
    owner_info = status.get(owner, {})
    last_print_time = owner_info.get("last_print_time")
    if last_print_time:
        last_print_time = datetime.fromisoformat(last_print_time)
        if current_time - last_print_time < timedelta(minutes=timeout_minutes):
            return True
    return False


# Funktion zum Freigeben eines Besitzers nach dem Timeout
def release_blocked_owner(owner, status):
    current_time = datetime.now()
    owner_info = status.get(owner, {})
    last_print_time = owner_info.get("last_print_time")
    if last_print_time:
        last_print_time = datetime.fromisoformat(last_print_time)
        if current_time - last_print_time >= timedelta(minutes=timeout_minutes):
            del status[owner]
            save_status(status)


# Funktion zum Drucken der Datei mit SumatraPDF
def print_with_sumatra(file_path):
    try:
        subprocess.run([sumatra_path, "-print-to", printer_path, file_path], check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Fehler beim Drucken der Datei {file_path} mit SumatraPDF: {e}")
        return False


# Funktion zum Drucken der ausgewählten Dateien im Archiv
def print_selected_files(selected_files, temp_folder, owner):
    failed_prints = []
    for file, nachname, vorname in selected_files:
        file_path = os.path.join(temp_folder, owner, file)
        if is_blacklisted(nachname, vorname):
            messagebox.showinfo("Druck gesperrt",
                                f"Der Druck für {nachname}, {vorname} ist gesperrt. Datei wurde nicht gedruckt.")
        elif not print_with_sumatra(file_path):
            failed_prints.append(file)

    if failed_prints:
        messagebox.showwarning("Teilweise Fehler",
                               f"Die folgenden Druckaufträge konnten nicht abgeschlossen werden: {', '.join(failed_prints)}")
    else:
        messagebox.showinfo("Drucken abgeschlossen", "Die ausgewählten Dateien wurden gedruckt.")


# Funktion zum Verschieben und Drucken der Dateien in Besitzer-Unterordner
def move_and_print_files(owner, folder_path, temp_folder, gui_refresh):
    status = load_status()
    current_time = datetime.now()

    # Prüfen, ob der Besitzer blockiert ist
    if is_owner_blocked(owner, status):
        messagebox.showerror("Druckauftrag blockiert",
                             f"Der Druckauftrag für {owner} ist noch blockiert. Bitte warten Sie den Timeout ab.")
        return

    # Aktualisiere den Status und blockiere den Besitzer
    status[owner] = {"last_print_time": current_time.isoformat()}
    save_status(status)

    if not check_printer_availability():
        messagebox.showerror("Drucker nicht erreichbar",
                             f"Der Drucker {printer_path} ist momentan nicht erreichbar. Die Dateien wurden nicht verschoben.")
        del status[owner]
        save_status(status)
        return

    owner_folder = os.path.join(temp_folder, owner)
    if not os.path.exists(owner_folder):
        os.makedirs(owner_folder)

    moved_files = []
    failed_prints = []
    blocked_files = []

    for filename, nachname, vorname, is_correction in analyse_pdfs(folder_path, temp_folder)[owner]["files"]:
        file_path = os.path.join(folder_path, filename)
        dst = os.path.join(owner_folder, filename)

        try:
            shutil.move(file_path, dst)
            moved_files.append(filename)
            if is_blacklisted(nachname, vorname):
                blocked_files.append(f"{nachname}, {vorname}")
            elif not print_with_sumatra(dst):
                failed_prints.append(filename)
        except Exception as e:
            print(f"Fehler beim Verschieben oder Drucken der Datei {filename}: {e}")
            failed_prints.append(filename)

    if blocked_files:
        messagebox.showinfo("Druck gesperrt",
                            f"Der Druck für die folgenden Schüler ist gesperrt: {', '.join(blocked_files)}. Dateien wurden nicht gedruckt.")

    if moved_files and not blocked_files:
        if failed_prints:
            messagebox.showwarning(
                "Teilweise Fehler",
                f"Die Dateien wurden verschoben, aber die folgenden Druckaufträge konnten nicht abgeschlossen werden: {', '.join(failed_prints)}"
            )
        else:
            messagebox.showinfo("Erfolg", f"Alle Dateien von {owner} wurden erfolgreich verschoben und gedruckt.")

    # Prüfen, ob der Besitzer nach dem Timeout freigegeben werden kann
    release_blocked_owner(owner, status)
    gui_refresh()


# Funktion zur automatischen Aktualisierung des GUI-Status
def update_gui_status(buttons, root):
    while True:
        time.sleep(5)
        status = load_status()
        root.after(0, lambda: apply_status_to_buttons(status, buttons))


# Funktion zum Anwenden des Status auf die Schaltflächen im Hauptthread
def apply_status_to_buttons(status, buttons):
    current_time = datetime.now()
    for owner, info in status.items():
        last_print_time = info.get("last_print_time")
        if last_print_time:
            last_print_time = datetime.fromisoformat(last_print_time)
            if current_time - last_print_time >= timedelta(minutes=timeout_minutes):
                del status[owner]
                save_status(status)


# Funktion zum Erstellen der GUI mit automatischer Aktualisierung im Akkordeon-Stil
def create_gui(folder_path, temp_folder):
    root = tk.Tk()
    root.title("Aufträge Zeugnisdruck")

    tk.Label(root, text="Aufträge Zeugnisdruck", font=("Arial", 14)).pack(pady=10)

    canvas = tk.Canvas(root)
    scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    archive_button = tk.Button(root, text="Zum Tagesarchiv", command=lambda: open_archive(temp_folder))
    archive_button.pack(pady=10)

    buttons = {}

    def refresh_gui():
        for widget in scrollable_frame.winfo_children():
            widget.destroy()
        owner_data = analyse_pdfs(folder_path, temp_folder)
        status = load_status()

        for owner, data in owner_data.items():
            owner_frame = tk.Frame(scrollable_frame)
            owner_frame.pack(fill="x", pady=5)

            # Fettgedruckter Besitzername
            owner_label = tk.Label(owner_frame, text=f"{owner}: {data['count']} Dateien", font=("Arial", 10, "bold"))
            owner_label.pack(side="left", padx=10)

            # Erstelle ein Frame für die Dateiliste und verstecke es initial
            files_frame = tk.Frame(owner_frame)
            files_frame.pack(fill="x", padx=20, pady=5)
            files_frame.pack_forget()

            # Button, um die Dateiliste ein- oder auszublenden (Akkordeon-Effekt)
            toggle_button = tk.Button(owner_frame, text="Dateien anzeigen")
            toggle_button.pack(side="right")

            # Setze den Befehl für den Button, nachdem er definiert wurde
            toggle_button.config(
                command=lambda f=files_frame, b=toggle_button: toggle_files_view(f, b)
            )

            for filename, nachname, vorname, is_correction in data['files']:
                file_color = "blue" if is_correction else "black"
                file_label = tk.Label(files_frame, text=filename, fg=file_color)
                file_label.pack(anchor="w", padx=10)

            print_button = tk.Button(owner_frame, text="Verschieben und Drucken")
            print_button.pack(side="right")
            print_button.config(command=lambda o=owner: move_and_print_files(o, folder_path, temp_folder, refresh_gui))

            buttons[owner] = print_button

    # Funktion zum Ein- oder Ausblenden der Dateiliste
    def toggle_files_view(frame, button):
        if frame.winfo_ismapped():
            frame.pack_forget()
            button.config(text="Dateien anzeigen")
        else:
            frame.pack(fill="x", padx=20, pady=5)
            button.config(text="Dateien verbergen")

    refresh_gui()

    # Funktion zur automatischen Aktualisierung des Hauptfensters
    def auto_refresh():
        refresh_gui()
        root.after(5000, auto_refresh)

    auto_refresh()
    Thread(target=update_gui_status, args=(buttons, root), daemon=True).start()
    root.mainloop()


# Funktion zum Öffnen des Tagesarchivs mit Scroll-Funktion
def open_archive(temp_folder):
    if hasattr(open_archive, 'archive_window') and open_archive.archive_window.winfo_exists():
        open_archive.archive_window.lift()
        return

    archive_window = tk.Toplevel()
    open_archive.archive_window = archive_window
    archive_window.title("Tagesarchiv")
    archive_window.geometry("600x500")

    # Hinzufügen eines Scrollable Canvas
    archive_canvas = tk.Canvas(archive_window)
    archive_scrollbar = tk.Scrollbar(archive_window, orient="vertical", command=archive_canvas.yview)
    archive_frame = tk.Frame(archive_canvas)

    archive_frame.bind("<Configure>", lambda e: archive_canvas.configure(scrollregion=archive_canvas.bbox("all")))
    archive_canvas.create_window((0, 0), window=archive_frame, anchor="nw")
    archive_canvas.configure(yscrollcommand=archive_scrollbar.set)

    archive_canvas.pack(side="left", fill="both", expand=True)
    archive_scrollbar.pack(side="right", fill="y")

    tab_control = ttk.Notebook(archive_frame)
    tab_control.pack(expand=1, fill="both")

    owner_data = {}
    check_vars = {}

    for owner in os.listdir(temp_folder):
        owner_folder = os.path.join(temp_folder, owner)
        if os.path.isdir(owner_folder):
            files = os.listdir(owner_folder)
            owner_data[owner] = files

    for owner, files in owner_data.items():
        owner_frame = tk.Frame(tab_control)
        tab_control.add(owner_frame, text=owner)

        owner_label = tk.Label(owner_frame, text=f"{owner}: {len(files)} Dateien")
        owner_label.pack(anchor="w", pady=5)

        owner_var = tk.IntVar()
        select_all_cb = tk.Checkbutton(owner_frame, text="Alle auswählen", variable=owner_var)
        select_all_cb.pack(anchor="w", padx=10)

        check_vars[owner] = []
        for file in files:
            var = tk.IntVar()
            nachname, vorname = extract_name_from_filename(file)
            cb = tk.Checkbutton(owner_frame, text=file, variable=var,
                                state=tk.DISABLED if is_blacklisted(nachname, vorname) else tk.NORMAL)
            cb.pack(anchor="w", padx=10)
            check_vars[owner].append((file, var))

        print_button = tk.Button(owner_frame, text="Ausgewählte Dateien drucken",
                                 command=lambda o=owner: print_selected_files(
                                     [(f, extract_name_from_filename(f)[0], extract_name_from_filename(f)[1]) for f, v
                                      in check_vars[o] if v.get() == 1], temp_folder, o))
        print_button.pack(pady=10)
        select_all_cb.config(command=lambda o=owner, ov=owner_var: select_all_files(o, ov.get(), check_vars))

    tab_control.pack(expand=1, fill="both")


# Hilfsfunktion zum Extrahieren von Nachname und Vorname aus dem Dateinamen
def extract_name_from_filename(filename):
    parts = filename.split("_")
    if len(parts) >= 3:
        nachname = parts[1].strip()
        vorname = parts[2].split(".")[0].strip()
        return nachname, vorname
    return "", ""


# Funktion zum Auswählen aller Dateien
def select_all_files(owner, select_all, check_vars):
    for file, var in check_vars[owner]:
        var.set(1 if select_all else 0)


if __name__ == "__main__":
    create_gui(folder_path, temp_folder)
