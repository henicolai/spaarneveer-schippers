# Schippersrooster Spaarneveer

Online rooster voor de vrijwillige schippers van het Spaarneveer in Haarlem.

- **Pagina:** https://henicolai.github.io/spaarneveer-schippers/ (ook ingebed op spaarneveer.nl)
- **Beheer:** Henk Nicolai – hnicolai@gmail.com

## Hoe het in elkaar zit

| Onderdeel | Waar | Wat |
|---|---|---|
| Pagina | `index.html` in deze repository (GitHub Pages) | Alles wat schippers en beheerders zien. Eén bestand, gewone HTML/JavaScript. |
| Opslag en e-mail | Google Apps Script "Spaarneveer schippersrooster" in de Drive van de beheerder | Web-app die de pagina aanroept. Reservekopie: `beheer/Code.gs`. |
| Gegevens | Google Sheet "Spaarneveer wijzigingen" | Tabbladen: Wijzigingen, Mededelingen, Rooster, Schippers (namen + e-mail), Passagiers. |

De schipperscode en de beheercode staan **niet** in deze (openbare) repository.
Ze staan alleen in het Apps Script en worden persoonlijk doorgegeven.

## Handleidingen

- `beheer/Handleiding_schippers.txt` – voor alle schippers
- `beheer/Handleiding_beheerders.txt` – voor beheerders (zonder codes)

## Overnemen door een nieuwe beheerder

1. Laat de huidige eigenaar de Sheet en het Apps Script met je delen (bewerkrechten),
   en je toevoegen als collaborator op deze repository.
2. Wil je alles in je eigen Google-account: maak een kopie van de Sheet, maak een nieuw
   Apps Script-project, plak `beheer/Code.gs` erin en vul in:
   - `SCHIPPERSCODE` en `BEHEERCODE`
   - het ID van je eigen Sheet bij `openById(...)`
   - `BEHEERDER_EMAIL` (afzender van de mails)
3. Voer in de editor eenmalig `testMail` uit om toestemming voor e-mail te geven.
4. Implementeren > Nieuwe implementatie > Web-app, uitvoeren als "Mij",
   toegang "Iedereen". Kopieer de web-app-URL.
5. Zet die URL in `index.html` bij `API_URL` en commit. Klaar.

Na elke wijziging in het script: Implementeren > Implementaties beheren > potlood >
Versie: Nieuwe versie > Implementeren (de URL blijft dan hetzelfde).

## Limieten

Een gewoon Gmail-account mag via Apps Script ongeveer 100 ontvangers per dag mailen.
Is dat op, dan slaat het script de mail over; het rooster blijft gewoon werken.
