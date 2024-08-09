import os


def laagste_bestandsnaam(mapnaam):
    bestandsnamen = os.listdir(mapnaam)
    laagste_naam = None
    laagste_nummer = float("inf")  # Begin met oneindig groot nummer

    for bestandsnaam in bestandsnamen:
        if bestandsnaam.endswith(".jpg"):
            nummer = int(
                bestandsnaam[:-4]
            )  # Verwijder de extensie en converteer naar een integer
            if nummer < laagste_nummer:
                laagste_nummer = nummer
                laagste_naam = bestandsnaam

    return laagste_naam


# Voorbeeldgebruik:
mapnaam = r"pasfotos eerste jaar\1e jaar\1L1"
laagste = laagste_bestandsnaam(mapnaam)
print("Bestand met laagste nummer:", laagste)
