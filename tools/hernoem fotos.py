import os
import json

# Pad naar de map met foto's en het JSON-bestand
fotos_map = r"pasfotos\1e jaar 18-19"  # Vul dit aan met het juiste pad
json_bestand = r"klaslijsten eerste jaar.json"  # Vul dit aan met het juiste pad

# Laad de JSON gegevens
with open(json_bestand, "r", encoding="utf-8") as file:
    data = json.load(file)

# Itereer door elke klasgroep
for klasgroep in data["klasgroepen"]:
    klasgroep_naam = klasgroep["klasgroep"]
    studenten = klasgroep["leerlingen"]

    # Pad naar de huidige klasgroep map
    klas_map = os.path.join(fotos_map, klasgroep_naam)

    if not os.path.exists(klas_map):
        print(
            f"De map voor klasgroep {klasgroep_naam} bestaat niet. Controleer het pad."
        )
        continue

    # Laad bestandsnamen in
    bestandsnamen = os.listdir(klas_map)

    # Hernoem de bestanden in de map
    for student, bestandsnaam in zip(studenten, bestandsnamen):
        student_nummer = student["Nr"]
        student_naam = student["Naam"]

        # Vorm nieuwe bestandsnaam
        nieuwe_bestandsnaam = f"{student_naam}.jpg"

        # Volledig pad naar de oude en nieuwe bestanden
        oud_pad = os.path.join(klas_map, bestandsnaam)
        nieuw_pad = os.path.join(klas_map, nieuwe_bestandsnaam)

        if os.path.exists(oud_pad):
            os.rename(oud_pad, nieuw_pad)
            print(f"Hernoemd: {bestandsnaam} -> {nieuwe_bestandsnaam}")
        else:
            print(f"Bestand niet gevonden: {bestandsnaam} in map {klas_map}")

print("Hernoem proces voltooid.")
