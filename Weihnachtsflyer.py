# Input: Reporter-Abfrage "Gesamtliste Weihnachtsflyer"
#        in Zwischenablage, dann dieses Skript starten

import csv
import io
import win32clipboard
import datetime
import re

# Reguläre Ausdrücke, um verschiedene Schreibweisen von Adressen vereinheitlichen zu können
re_str = re.compile(r"[ -]?str(?:\.|a[ßs]+e\b)", re.IGNORECASE)
re_weg = re.compile(r"[ -]?weg\b", re.IGNORECASE)

heute = datetime.datetime.strftime(datetime.datetime.today(), "%Y-%m-%d")
methusalem = "01.01.1900"  # Dummy-Geburtsdatum für Einträge, die keines haben

print("Bitte Reporter-Abfrage 'Gesamtliste Weihnachtsflyer'")
print("durchführen und Daten in Zwischenablage ablegen.")
input("Bitte ENTER drücken, wenn dies geschehen ist!")

win32clipboard.OpenClipboard()
data = win32clipboard.GetClipboardData()
win32clipboard.CloseClipboard()

if not data.startswith("lfd. Nr.\t"):
    print("Fehler: Unerwarteter Inhalt der Zwischenablage!")
    exit()

nummern = set()
daten = []
with io.StringIO(data) as infile:
    for i, eintrag in enumerate(csv.DictReader(infile, delimiter="\t")):
        if not (nr:=eintrag["Nummer"]) in nummern:
            daten.append(eintrag)
            nummern.add(nr)
    fieldnames = list(eintrag.keys())

print(f"{i} Datensätze in der Abfrage.")
print(f"{i-len(nummern)} Dubletten entfernt, bleiben noch {len(nummern)}.")

# Ähnliche Schreibweisen bei Adressen normalisieren
# und Geburtsdatum in Zeitstempel umwandeln (für evtl. Sortierung/Vergleiche)
for eintrag in daten:
    temp = re_str.sub("straße", eintrag["Straße/Postfach"])
    temp = re_weg.sub("straße", temp)
    eintrag["str_norm"] = re.sub(r"\s+","",temp)
    if geb := eintrag["Geburtsdatum"]:
        eintrag["Geburtsdatum"] = datetime.datetime.strptime(geb, "%d.%m.%Y")
    else:
        eintrag["Geburtsdatum"] = datetime.datetime.strptime(methusalem, "%d.%m.%Y")

daten.sort(key = lambda l: l["Geburtsdatum"], reverse=True)

familien = {}       # Sammlung aller identischen Adressen (key = Adresse, value = Nummer der 1. Person)
personen = {}       # Sammlung aller Personen, deren Adresse zum Einsatz kommen wird
adressen = {}       # Sammlung aller Adressen, an denen Personen mit unterschiedlichen Namen wohnen
                    
for eintrag in daten:
    name = eintrag["Name"]
    straße = eintrag["str_norm"]
    plz = eintrag["PLZ"]
    # Gibt es schon einen Eintrag mit diesen Daten?
    if (adr := (name, straße, plz)) in familien:
        eintrag["Familie"] = familien[adr]
        personen[familien[adr]]["Familie"].append(f'{eintrag["Vorname"]} {eintrag["Name"]} ({eintrag["Nummer"]})')
    else:
        familien[adr] = eintrag["Nummer"]
        personen[eintrag["Nummer"]] = {"Daten": eintrag, "Familie": []}

# Nur Einträge mit mehr als einer Person behalten
familien = {k: v for k, v in familien.items() if personen[v]["Familie"]}

print(f'{len(familien)} gleiche Nachnamen/Anschrift-Kombinationen mit zusammen {len(familien) + sum(len(v["Familie"]) for v in personen.values())} Personen.')
print(f"Es bleiben somit noch {len(personen)} eindeutige Anschriften übrig.")

for eintrag in daten:
    name = eintrag["Name"]
    straße = eintrag["str_norm"]
    plz = eintrag["PLZ"]
    adressen.setdefault((straße, plz), {}).setdefault(eintrag["Name"], []).append(eintrag["Nummer"])

# Nur Adressen mit mehr als einem Nachnamen behalten

adressen = {k: v for k, v in adressen.items() if len(v) > 1}

print(f"Darüber hinaus gibt es {len(adressen)} Adressen, an denen Personen mit unterschiedlichen Nachnamen wohnen.")

fieldnames.extend(["Familie", "gleiche Adresse", "str_norm", "Adressat"])

output = []

for eintrag in personen.values():
    daten = eintrag["Daten"]
    daten["Familie"] = ", ".join(eintrag["Familie"])
    daten["gleiche Adresse"] = ", ".join(adressen.get((daten["str_norm"], daten["PLZ"]), []))
    alter = datetime.datetime.today() - daten["Geburtsdatum"]
    # Wohnt an dieser Adresse mehr als ein Mitglied/VIP... mit dem gleichen Nachnamen?
    # Oder ist der Adressat unter 18? Dann "Familie X" als Anschrift
    if daten["Familie"] or alter.days < 18*365.25:
        daten["Adressat"] = f'\nFamilie {daten["Name"]}'
    # Sonst die normale Anrede (Herrn/Frau Titel Vorname Nachname)
    else:
        if (anr:=daten["Adressanrede"]) == "Familie":
            daten["Adressat"] = f'\nFamilie '
        else:
            daten["Adressat"] = f'{anr}\n'
        if titel:=daten["Titel"]: 
            daten["Adressat"] += f'{titel} '
        if vorname:=daten["Vorname"]: 
            daten["Adressat"] += f'{vorname} '
        daten["Adressat"] += f'{daten["Name"]}'
    output.append(daten)

output.sort(key=lambda l:(l["Land"], l["PLZ"], l["Straße/Postfach"], l["Name"]))
            
with open(f"Weihnachtsflyer {heute}.csv", "w", encoding="cp1252", newline="") as outfile:
    writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter=";")
    writer.writeheader()
    writer.writerows(output)