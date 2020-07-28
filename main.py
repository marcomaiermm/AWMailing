import smtplib
import datetime
import sys
import os
import fnmatch
import sqlite3
import configparser
import time
import shutil
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from shutil import copy2
from pathlib import Path

__author__ = "Marco Maier"
__copyright__ = "Copyright 2020, Allweier Präzisionsteile GmbH"
__credits__ = ["Marco Maier"]
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "Marco Maier"
__email__ = "mm.maiermarco@gmail.com"
__status__ = "Development"


def new_folder():
    today = datetime.datetime.today().date().strftime('%d.%m.%Y')
    arch_date = input(
        "Datum für den Namen des Archivordners (z.B.: '"+today+"') Default bei keiner Eingabe ist "+today+" ... ") or today
    if not arch_date:
        print("that was nothing!")

    try:
        date_name = datetime.datetime.strptime(
            arch_date, '%d.%m.%Y').strftime('%Y-%m-%d')
    except ValueError:
        raise ValueError("Bitte richtiges Datumsformat wählen: DD.MM.YYYY")
        # print("Bitte richtiges Datumsformat wählen: DD.MM.YYYY")

    path = os.path.join(TARGET_AV_FOLDER, date_name)
    try:
        os.mkdir(path)
    except OSError:
        print("Erstellung des Ordners %s war nicht erfolgreich. Bereits vorhanden?." % path)
        pass
    else:
        print("Ordner %s erfolgreich erstellt." % path)

    return path


def date_check(path, dest):
    today = datetime.datetime.today().date()
    new_files = {
        "pure": [],
        "head": [],
        "tail": [],
        "path": [],
        "old_path": []
    }
    old_files = []
    for file in os.listdir(path):
        if os.path.isfile(os.path.join(path, file)) and not file.endswith('.link'):
            file_path = os.path.join(path, file)
            file_age = datetime.datetime.fromtimestamp(
                os.stat(file_path).st_mtime).date()
            if file_age < today:
                old_file = str(str(file_path)+" alter: "+str(file_age))
                old_files.append(old_file)
                cont = input("\n"+"Datei ist älter als heute: " +
                             old_file + "\n" + "\nfortfahren? ... (y/n)") or 'y'
                if cont.capitalize() != 'Y':
                    break

            else:
                # "Archiv\"
                file_name_raw_head, file_name_raw_tail = os.path.split(file)
                file_name = os.path.splitext(file)[0]
                # copy2 kopiert auch Metadaten und Schreibrechte der Dateien
                new_dest = copy2(file_path, dest)
                new_files['pure'].append(file_name)
                new_files['head'].append(file_name_raw_head)
                new_files['tail'].append(file_name_raw_tail)
                new_files['path'].append(new_dest)
                new_files['old_path'].append(file_path)

    print("Dateien erfolgreich kopiert: ", new_files)
    if old_files:
        print("Folgende Dateien sind älter als heute: ", old_files)
    return new_files


def get_ticket_no(path):
    ticket_no = ""
    for file in os.listdir(path):
        if os.path.isfile(os.path.join(path, file)):
            if fnmatch.fnmatch("_Ticket_", "*"):
                # substring ohne extension
                file_name = os.path.splitext(file)[0]
                ticket_no = file_name.partition("_Ticket_")[2]
    return ticket_no


def cleanse_tickets(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Konnte %s nicht löschen. Grund: %s' % (file_path, e))


def remove_files(paths):
    for path in paths:
        os.remove(path)


def execute_read_query(connection, query):
    cursor = connection.cursor()
    result = None
    try:
        cursor.execute(query)
        result = cursor.fetchall()
        return result
    except Exception as e:
        print(f"Fehler '{e}' trat auf")


def execute_write_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        connection.commit()
        print("Eintrag erfolgreich eingespeichert.")
    except Exception as e:
        print(f"Fehler '{e}' trat auf")


class User(object):
    def __init__(self):
        super(User, self).__init__()

        self.name = ""
        self.vorname = ""
        self.role_1 = ""
        self.role_2 = ""
        self.phone = ""
        self.email = ""

    def get_user(self, name):
        connection = sqlite3.connect(USER_DATABASE)
        select = open(os.path.join(THIS_FOLDER, "sql/user.sql"),
                      encoding='utf8').read().format(name=name)
        user = execute_read_query(connection, select)
        if not user:
            option = input("\nKeine Einträge gefunden. Erstellen? (y/n) ... ")
            if option == 'y':
                name = input("Nachname: ... ")
                vorname = input("Vorname: ... ")
                role_1 = input("Rolle 1: ... ")
                role_2 = input("Rolle 2: ... ")
                phone = input("Telefonnummer: ... ")
                email = input("Email-Adresse: ... ")

                query = open(os.path.join(THIS_FOLDER, "sql/input_user.sql"),
                             encoding='utf8').read().format(name=name, vorname=vorname, role_1=role_1, role_2=role_2, phone=phone, email=email)
                execute_write_query(connection, query)
                user = execute_read_query(connection, select)
            else:
                return
        self.name = user[0][1]
        self.vorname = user[0][2]
        self.role_1 = user[0][3]
        self.role_2 = user[0][4]
        self.phone = user[0][5]
        self.email = user[0][6]


class Mail(User):
    def __init__(self):
        super(Mail, self).__init__()
        #dt = datetime.datetime.now()
        self.host = ("localhost",)
        self.ticket_no = ticket_no
        self.to = config["mail"]["To"]
        self.cc = config["mail"]["Cc"]
        self.subj = "Aktuelle Auswertungen - Ticket#" + self.ticket_no
        self.html = ""

    def create_message(self):
        return HTML_EMAIL.format(
            auswertungen=html_files, name=self.name+" "+self.vorname, role_1=self.role_1, role_2=self.role_2, phone=self.phone, mail=self.email)

    def send_mail(self):
        message = MIMEMultipart()
        message["From"] = self.email
        message["To"] = self.to
        message["Subject"] = self.subj
        message["Cc"] = self.cc

        part1 = MIMEText(self.html, "html")
        message.attach(part1)

        # alle dateien anhängen
        for i, f in enumerate(files['path']):
            try:
                with open(f, 'rb') as fp:
                    part2 = MIMEBase('application', "octet-stream")
                    part2.set_payload(fp.read())
                    encoders.encode_base64(part2)
                    part2.add_header('Content-Disposition',
                                     'attachment', filename=files["tail"][i])
                    message.attach(part2)
            except Exception as e:
                print("Anhängen der Daten hat nicht funktioniert..Fehler:", e)
                pass

        if self.host[1] != '':
            s = smtplib.SMTP(self.host[0], self.host[1])
        else:
            s = smtplib.SMTP(self.host[0])
        s.sendmail(self.email, self.to, message.as_string())
        s.quit()


THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))

# Auslesen der Config Datei für Mailing Host
config = configparser.ConfigParser()
config.read(os.path.join(THIS_FOLDER, "DATA/config.ini"))

USER_DATABASE = os.path.join(THIS_FOLDER, "sql/user.db")
AV_FOLDER = Path(config["path"]["Auswertungen"])
TARGET_AV_FOLDER = Path(config["path"]["Auswertungen"]+"/Archiv/")
AV_TEMP_FOLDER = Path(str(AV_FOLDER) + r"/Auswertung_Temp/")
HTML_EMAIL = open(os.path.join(THIS_FOLDER, "DATA/email.html"),
                  encoding='utf8').read()


path_archiv_folder = new_folder()

rp_folder_str = str(AV_FOLDER)
rp = input("\n"+"Ist " + rp_folder_str +
           " der richtige Dateipfad? (y/n) ... ") or 'y'
if rp.capitalize() != 'Y':
    AV_FOLDER = Path(input("Bitte neuen Pfad eingeben ... " + "\n"))
    config["path"]["Auswertungen"] = AV_FOLDER

ticket_no = get_ticket_no(AV_TEMP_FOLDER)


# Dateinamen aus Liste generieren und als string definieren
files = date_check(AV_FOLDER, path_archiv_folder)
# Dateinamen als HTML Text generieren
html_files = "".join(["<p>"+e+"</p>" for e in files['pure']])

print("\nAngaben für die E-mail Signatur...")
user_name = input("\nNachname: ... ")

# Benutzer wird von Mail Klasse vererbt
# Mail wird erstellt und gesendet

new_email = Mail()
new_email.host = (config['mailserver']['IP'], config['mailserver']['Port'])
new_email.get_user(user_name)
print("Mail wird generiert mit der Signatur von \n",
      new_email.vorname, " ", new_email.name)
new_email.html = new_email.create_message()
new_email.send_mail()

delete_tickets = input(
    "Email erfolgreich gesendet.. \nTickets löschen? (y/n)") or 'y'
if delete_tickets.capitalize() == 'Y':
    cleanse_tickets(AV_TEMP_FOLDER)
    print("Tickets wurden gelöscht.")
else:
    print("Tickets wurden nicht gelöscht")


# Löschen der Gesendeten Dateien aus dem alten Pfad
print("Kopierte Original-Dateien werden gelöscht ... ")
try:
    remove_files(files["old_path"])
except Exception as e:
    print("Alte Dateien konnten wegen ", e, " nicht gelöscht werden.")

print("\nProgramm schließt sich in 10 sekunden automatisch...")

# Löschen der verbleibenden Daten im Temp Ordner

# time.sleep(10)
