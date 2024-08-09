import json

json_bestand = r"klaslijsten\klaslijsten zesde jaar.json"
# Laad de JSON gegevens
with open(json_bestand, "r", encoding="utf-8") as file:
    data = json.load(file)


# Functie om de naam om te draaien
def draai_naam_om(naam):
    delen = naam.split()
    voornaam = delen[-1]
    achternaam = " ".join(delen[:-1])
    return f"{voornaam} {achternaam}"


# Pas de namen aan
for klasgroep in data["klasgroepen"]:
    for leerling in klasgroep["leerlingen"]:
        leerling["Naam"] = draai_naam_om(leerling["Naam"])

# Schrijf de gewijzigde gegevens terug naar het JSON-bestand
with open(json_bestand, "w", encoding="utf-8") as file:
    json.dump(data, file, indent=2, ensure_ascii=False)

print("De namen zijn succesvol omgedraaid en opgeslagen in het bestand.")
