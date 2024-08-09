# Proclamatie PowerPoint Generator

Dit Python script genereert automatisch een PowerPoint-presentatie voor de proclamatie, inclusief foto's van leerlingen en informatie over hun klassen. De code is getest met Python 3.12 en pptx 1.0.2.

## Installatie

1. **Benodigdheden:**
   - Python
   - PowerPoint
   - python-pptx:

     Voer het volgende command in om   PPTX te installeren:
     ```bash
     pip install python-pptx
     ```

2. **Klaslijsten en Foto's:**
   - Zorg ervoor dat je de klaslijsten hebt in het volgende JSON-formaat:
     ```json
     {
        "titularissen": [
            {
                "naam": "[voornaam achternaam]",
                "klassen": [
                    {
                        "klas": "[klas]",
                        "leerlingen": [
                            {
                                "Nr": "[nummer]",
                                "Naam": "[voornaam achternaam]"
                            }
                        ]
                    }
                ]
            }
        ]
     }
     ```
   - Je kunt ChatGPT gebruiken om PDF-bestanden met klaslijsten om te zetten naar dit formaat. **Zorg ervoor dat je dit doet in een tijdelijke sessie om privacy te waarborgen!!**
   - Alle foto's moeten worden opgeslagen met de naam van de persoon zoals in het JSON-bestand. Bijvoorbeeld: `voornaam achternaam.jpg` of `voornaam.achternaam.jpg`. Zorg er voor dat het python script het correcte path gebruikt om de foto's te zoeken.

## Gebruik

1. **Paden Configureren:**
   - Pas de paden in de code aan naar de locaties van je JSON-bestand en fotomappen.
   - pas de titel slide aan naar het juiste jaartal
   - zorg dat de placeholders voor de fotos in het powerpoint sjabloon de correcte dimenties hebben, zie opmerkingen.

2. **Voer de Code Uit:**
   - Na het instellen van de paden, kun je de code uitvoeren om de PowerPoint-presentatie te genereren:
     ```bash
     python 'Powerpoint Maker v3.py'
     ```

3. **Opmerkingen bij de Code:**
   - Aanpassingen aan de lay-out moeten gebeuren in het PowerPoint-sjabloon door het diamodel aan te passen. Dit doe je door de presentatie te openen in PowerPoint, naar de tab 'Beeld' te gaan en daar op 'Diamodel' te klikken. Controleer vervolgens of de ID-nummers in de code nog steeds overeenkomen met de placeholders in het sjabloon.
   - Controleer zeker de leerlingen met accenten of meerdere namen, en maak waar nodig aanpassingen.

### Voorbeeld
Kijk in de map 'example' om een voorbeeld van de benodigdheden en het resultaat van dit script te zien.

## Uitbreidingsideeën

- Quotes toevoegen bij elke leerling.
- Een foto van de klastitularis toevoegen op de klas slide.
- Een groepsfoto van de klas toevoegen op de klas slide.

## Meer Informatie

Voor meer informatie over het gebruik van `python-pptx` om PowerPoint-presentaties te bewerken, bezoek de officiële documentatie: [python-pptx documentatie](https://python-pptx.readthedocs.io/en/latest/index.html).

## Bijdragen

Contributies aan deze code zijn zeker welkom! Als je vragen hebt of tegen problemen aanloopt, open dan gerust een issue in de repository.

