# Handleiding voor de leraar — Les 6: Zoeken als een onderzoeker

**Vak:** Toegepaste Informatica
**Doelgroep:** klas 3MWWE — 2de graad Maatschappij- en welzijnswetenschappen, doorstroomfinaliteit
**Lesduur:** 1 × 50 minuten: **10 minuten instructie + 40 minuten keuzewerktijd**
**Context:** het Onderzoekslabo Jongeren & Welzijn (fictief), begeleider Ines Coppens: een eerste literatuurverkenning voor een onderzoek naar schermgebruik en slaap
**Waar:** lokaal 18 (Windows 11)
**Lesdag:** donderdag 26 november 2026, 8ste lesuur (14.35–15.25) — de laatste les vóór de examens (vanaf 2 december) en Kerstmis
**Kernleerplandoel:** `BV2_13.04` (doelgericht informatie zoeken; subdoel `BV2_13.04.01`: een zoekstrategie bepalen) · ondersteunend `BV2_02.05` · contextueel `BV2_06.19` (alleen de zoekoperatoren)
**Opdracht in Classroom:** *Les 6 — Zoeken als een onderzoeker*, deadline vrijdag 27 november 2026, 20.00 uur

<!-- Beginvragen (stap 1 van CONTEXT.md), beantwoord door Jonas op 25-11-2026 (telkens het voorstel gekozen):
     klas 3MWWE · les 6 van het lessenplan MWW (Zoeken als een onderzoeker), de laatste les vóór de examens (vanaf 2-12)
     en Kerstmis · do 26-11-2026, 8ste uur (14.35–15.25) · vorige les: les 5 (Netiquette, phishing en de groepschat,
     19-11) is gegeven → lesstart = ophalen uit les 5 · één kern (regels van 19-11-2026): een zoekstrategie in Google
     voor een onderzoeksvraag van het labo (trefwoorden, OR, - om uit te sluiten, "aanhalingstekens", één filter) en
     vergelijken wat elke zoekopdracht oplevert; product: een kort zoekverslag. Waarheidstabellen, Google Scholar en de
     bibliotheekcatalogus vallen weg · lokaal 18, Windows 11 · standaard: één werkdocument (Google-document, kopie per
     leerling), geen zip · 20 punten, deadline vr 27-11-2026 20.00 · context Onderzoekslabo Jongeren & Welzijn. -->

## 1. Inhoud van het pakket

```text
W12 - Les 06 - MWW - Zoeken als een onderzoeker/
├── index.html              de lespagina: startpagina in beelden · route · drie stappen · Inleveren · checklist (7 taken)
├── presentatie.html        6 dia's voor het bord; dia 1 is de vaste startdia, dia 6 heeft kloktijden uit de planning
├── css/style.css           opmaak van de lespagina (onderaan: de tekeningen van Google van deze les)
├── css/slides.css          opmaak van de dia's
├── js/script.js, js/slides.js   interactie (zelftest + codewoord TREFWOORD in stap 4, kloktijden)
├── assets/                 logo's Onderzoekslabo en GO! Dalton Gent, lettertypes (OFL), mascotte, video/inleveren.mp4
├── werkdocument/
│   ├── Zoekverslag-onderzoekslabo.docx   het werkdocument dat de leerling INLEVERT (een kopie per leerling)
│   ├── maak_werkdocumenten.py            genereert het docx (python-docx)
│   └── huisstijl.py, go-dalton-gent.png  de huisstijl van het werkdocument
├── lesvoorbereiding.md     de lesvoorbereiding (28 punten + bronnenonderzoek)
├── dalton-lesfiche.html    de Dalton-lesfiche (instructie 10 + KWT 40): openen, Kopieer de fiche, plakken
├── lesdoelen.json          de doelen voor je jaaroverzicht
├── classroom.json          de opdracht voor de Classroom-koppeling
└── README.md               deze handleiding
```

Geen zip, geen gedeeld document en geen losse bestanden: alles wat de leerling maakt, typt hij **in** het werkdocument.

### 1b. Hoe de lespagina werkt

Volgens de regels van 19-11-2026: één kern (een zoekstrategie in Google), één product (het zoekverslag), drie delen, hoogstens één nieuwe
handeling (het filter onder **Tools**), de uitleg in het werkdocument, zes dia's.

- **Startpagina**: *Wat heb je nodig?* in drie beelden, *Dit maak je vandaag* (het zoekverslag als skelet, geen antwoorden), *Na deze
  les kan je* en *Vast?*.
- **Stap 2, *Toon stapjes*** (bijgestuurd op 26-11-2026, Jonas): *Zoek drie keer: A, B en C* heeft een stapje per zoekopdracht met een
  tekening van het zoekvak en wat elk stuk betekent; bij B twee keer *Niet* (de zin met OR erbij, alleen `gsm OR smartphone`).
  *Tel na elke zoekopdracht de resultaten* toont de tabel van deel 2 ingevuld als voorbeeld.
- **Vier stappen**: 1 *Van vraag naar trefwoorden* (trefwoorden en synoniemen kiezen, typen in het lichtblauwe vak), 2 *Drie keer
  zoeken* (Google openen in een nieuw tabblad, A, B en C zoeken, de gewone resultaten tellen), 3 *Filter en zoekverslag* (**Tools** ›
  **Elke periode** › **Afgelopen jaar**, een webadres kopiëren uit de adresbalk), 4 *Inleveren*. Sinds 26-11-2026 (Jonas) **begint**
  elke stap met *Maak deel n in je werkdocument.* en daaronder *Gebruik de informatie uit het blauwe kader hieronder*; de handelingen
  erna zeggen wat de leerling in dat deel doet. Vette regels van hoogstens 8 woorden. Alle beelden zijn tekeningen in HTML en CSS (het
  zoekvak van Google, de tabs met **Tools**, het menu, de resultaten met wat meetelt en wat niet, de adresbalk).
- **De uitleg staat op de pagina**, in een blauw kader onderaan stap 1, 2 en 3 (*Van vraag naar trefwoorden*, *Zo typ je het in
  Google*, *Een filter en een bruikbaar resultaat*), niet meer in het werkdocument (Jonas, 26-11-2026: "Ik zou ook wel graag de
  theoretische informatie toch terug in de lespagina willen verwerken"). Geen theoriekaart.
- **Stap 4 (*Inleveren*)**: *Wat moet je doen?* is vergrendeld achter drie meerkeuzevragen (OR, het minteken, het AI-overzicht niet
  tellen). Zijn ze alle drie juist, dan toont de groene balk het codewoord **TREFWOORD** (dat de leerling bovenaan deel 4 typt).
- **Checklist**: 7 taken. **Geen AI-hulp** (`data-ai-hulp="uit"`).

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages

Repository [`jonasdaltongent/Zoeken-Labo`](https://github.com/jonasdaltongent/Zoeken-Labo), Pages op branch `main`, map `/ (root)`:
- **Lespagina:** <https://jonasdaltongent.github.io/Zoeken-Labo/>
- **Presentatie:** <https://jonasdaltongent.github.io/Zoeken-Labo/presentatie.html>

### Stap 2 — De opdracht in Classroom, met de koppeling

De opdracht staat als **concept** klaar in `3MWWE - Informatica` (id `25614259410`) en is sinds 26-11-2026 **ingepland voor donderdag 26 november 2026 om 14.05 uur** (een half uur vóór het 8ste lesuur van 14.35 uur). Classroom publiceert dan zelf.

| | Opdracht: **Les 6 — Zoeken als een onderzoeker** |
|---|---|
| **Onderwerp** | Module 2 — Informatievaardigheden en mediawijsheid (nieuw in `3MWWE - Informatica`) |
| **Bijlage 1** | de link naar de lespagina (`Lespagina · Les 6 · Zoeken als een onderzoeker · Onderzoekslabo`) |
| **Bijlage 2** | `Zoekverslag-onderzoekslabo` (Google-document) — **Een kopie maken voor elke leerling** |
| **Punten** | 20 |
| **Deadline** | vrijdag 27 november 2026, 20.00 uur |

De instructietekst bevat geen e-mailadres (er wordt niets gedeeld).

```bash
python3 "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/_tools/zet_opdracht_klaar.py" "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W12 - Les 06 - MWW - Zoeken als een onderzoeker" --proef
```

### Stap 3 — Met de hand, als de koppeling niet werkt

Upload `werkdocument/Zoekverslag-onderzoekslabo.docx` in Drive zelf (**Nieuw** › **Bestand uploaden**; met de instelling *Uploads converteren*
wordt het een Google-document). Maak in Classroom een opdracht met **Bijvoegen** › **Link** (de lespagina) en **Bijvoegen** › **Drive**
(het werkdocument, **Een kopie maken voor elke leerling**), 20 punten, deadline vr 27-11 om 20.00 uur.

### Afvinklijst vóór de les

- [ ] De lespagina en presentatie staan online (HTTP 200).
- [ ] De opdracht staat als **concept** klaar in `3MWWE - Informatica` en is ingepland voor do 26-11 om 14.05 uur.
- [ ] Probeer zoekopdracht A (de hele vraag) en B (`slaap jongeren gsm OR smartphone`) één keer uit op de klascomputer: de demo moet werken.
- [ ] Het wachtwoord van de klascomputers op het bord (staat nooit op een dia).
- [ ] Open `presentatie.html` op het bord op **dia 1 (*Zo start je*)** terwijl de leerlingen binnenkomen.
- [ ] Kopieer de Dalton-lesfiche uit `dalton-lesfiche.html` naar je planner.

### 2b. Nagelezen klikpaden

Nagekeken op **google.com** (nl-BE, `hl=nl&gl=be`) in de ingebouwde browser, **26-11-2026** (venster van 1024 px breed, niet aangemeld bij
Google). Waar jouw scherm anders zegt, geldt jouw scherm.

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| De tabs onder het zoekvak | **AI-modus** · **Alle** · **Afbeeldingen** · **Video's** · **Korte video's** · **Web** · **Meer** · **Tools** | gezien op het scherm, 26-11-2026 |
| Een filter kiezen | onder het zoekvak **Tools** › het menu toont **Elke taal**, **Elke periode**, **Alle resultaten**, **Geavanceerd zoeken** | gezien op het scherm, 26-11-2026; de help ([Zoeken 142143](https://support.google.com/websearch/answer/142143?hl=nl&co=GENIE.Platform%3DDesktop)) schrijft *Tools* en de datumtool *Publicatiedatum* |
| De datumtool | heet op het scherm **Elke periode**; opties **Afgelopen uur**, **Afgelopen 24 uur**, **Afgelopen week**, **Afgelopen maand**, **Afgelopen jaar**, **Aangepaste periode…**. Na een keuze heet de knop bv. **Afgelopen jaar** (het adres krijgt `tbs=qdr:y`) | gezien op het scherm, 26-11-2026 |
| Filter weer weg | in het menu van **Tools** staat dan rood **Wissen** | gezien op het scherm, 26-11-2026; de help schrijft ook **Wissen** |
| Het AI-overzicht | boven de resultaten, kop **AI-overzicht** (met **Meer tonen**); bij een zoekopdracht met een periodefilter was het weg | gezien op het scherm, 26-11-2026; niet bij elke zoekopdracht |
| Advertenties | bij *smartphone kopen*: een kop **Gesponsorde resultaten** boven productkaartjes. Een gewone tekstadvertentie met het woord **Gesponsord** zag ik niet bij de zoekopdrachten van de les; de lespagina en het werkdocument schrijven *advertenties* en *Gesponsord* | gezien op het scherm, 26-11-2026 (de tekstadvertentie **niet** gezien: nakijken in lokaal 18) |
| Het webadres van een resultaat | na een klik op het resultaat staat het adres in de adresbalk van Chrome; kopiëren met **Ctrl + C** | algemene kennis (Windows), niet in de help |
| Inleveren | **Inleveren** › venster **Je werk inleveren?** › **Inleveren** | `_afspraken/klikpaden.md` (demoleerling3, 15-10-2026) |
| Operatoren | spatie · `OR` (hoofdletters) · `-` vlak voor het woord · `"…"` | [Zoeken 2466433](https://support.google.com/websearch/answer/2466433?hl=nl) en *Geavanceerd zoeken*, gelezen op 26-11-2026 (`_vakinhoud/informatievaardigheden.md`) |

## 3. Het verloop van de les

### Instructie en demo (10 minuten, dia 1–6)

- **Dia 1 (0–1'):** *Zo start je.* Wacht tot de meesten de lespagina open hebben.
- **Dia 2 (1–3'):** *Weet je het nog?* Vingervraag over les 5 (phishing) → **1**. Brug: de groepschat ging over het onderzoek naar slaap.
- **Dia 3 (3–4'):** het bericht van Ines met de onderzoeksvraag en het zoekverslag als skelet.
- **Dia 4 (4–9'):** demo in Google: de vraag als zin, tel de eerste 5; vul samen de eerste rij *gsm · smartphone* in (de andere twee trefwoorden kiezen ze zelf); typ
  `slaap jongeren gsm OR smartphone`, tel opnieuw; toon `-app`. **Niet** voorgedaan: aanhalingstekens en het filter.
- **Dia 5 (9–10'):** *Controle · iedereen antwoordt.* `slaap gsm -app` → **2**.
- **Dia 6 (10'):** *Zo werk je verder.* Stap 1 (Google openen in een nieuw tabblad, 3') doen ze nu samen: duim omhoog.

### Keuzewerktijd (40 minuten)

- **Eerste rondgang (stap 1 en 2):** trefwoorden en geen zin? `OR` in hoofdletters, het minteken vlak tegen het woord? Het AI-overzicht en
  de advertenties niet meegeteld?
- **Tweede rondgang (stap 3):** een webadres uit de adresbalk (niet van Google zelf), en een resultaat dat over slaap én schermen gaat?
- **Laatste minuut:** de exitvraag in deel 4, en zeg mondeling dat wie niet klaar is, toch inlevert. Vergeet niet: de volgende les is pas op 7 januari.

## 4. Verbetersleutel

Per criterium van § 22 van de lesvoorbereiding (samen 20):

- **Codewoord** bovenaan deel 4: **TREFWOORD** (ontbreekt of fout = −2; een tikfout zoals *TREFWORD* telt als juist).
- **Deel 1, tabel (2 punten):** twee trefwoorden uit de vraag naast *gsm* (0,5 per trefwoord), elk met een synoniem (0,5 per
  synoniem), elk vak één woord. Verwacht: *slapen/slaap* → *nachtrust, slaaptekort, slaapkwaliteit, sleep*; *jongeren* → *tieners,
  adolescenten, pubers, teenagers*; ook *'s avonds* als trefwoord. Niet: een klein woord, *slechter* of *zitten* als trefwoord, *gsm*
  nog eens, een zin, hetzelfde woord als synoniem, een woord met een andere betekenis (*bed*, *kinderen*, *dromen*). Sinds
  26-11-2026 zonder kolom *Idee* (Jonas: het idee was al het trefwoord).
- **Deel 1, vraag 2 (1 punt):** met twee woorden vind je ook pagina's die het andere woord gebruiken. Niet: *omdat het hetzelfde is*.
- **Deel 2, de aantallen (1 punt):** bij A, B en C een getal van 0 tot 5 (één ontbreekt 0,5; twee of meer 0). De getallen zelf tellen niet.
- **Deel 2, zoekopdracht B (2 punten):** alleen trefwoorden, geen zin (1); `OR` in hoofdletters tussen twee synoniemen (1).
- **Deel 2, zoekopdracht C (3 punten):** een woordgroep van minstens twee woorden tussen aanhalingstekens (1); een minteken vlak voor een
  woord, zonder spatie (1); de trefwoorden van B staan er nog in (1).
- **Deel 2, vraag 2 (2 punten):** welk soort resultaat van B weg moest (1), en of het bij C weg was (1).
- **Deel 3, vraag 1 en 2 (2 punten):** welk filter en welke periode (1); zoekopdracht D gekopieerd, B of C (1).
- **Deel 3, vraag 3 (4 punten):** per resultaat titel (0,5), volledig webadres met `https://` naar een gewone pagina (1), over slaap én schermen (0,5).
- **Deel 3, vraag 4 (2 punten):** waarom D beter werkt dan A (1), met iets wat de leerling in de resultaten zag (1).
- **Deel 4, exitvraag (1 punt):** een zoekopdracht met trefwoorden en `OR` in hoofdletters, bv. `stress examens jongeren OR tieners`.

Wat niet telt: de aantallen zelf, welke synoniemen precies, welke woordgroep en welk uitgesloten woord, welke periode, spelling en hoofdletters buiten `OR`.
Aftrek volgens `_afspraken/evaluatie-en-privacy.md` (niets ingeleverd = 0, te laat = −5, codewoord leeg of fout = −2).

### Essentiële fouten — geef hier altijd feedback op

- In B of C nog een **hele zin** als zoekopdracht.
- `or` in kleine letters, `NOT` in plaats van het minteken, of een **spatie na het minteken** (`- app`).
- Een resultaat uit het **AI-overzicht** of een **advertentie** in deel 3, of een webadres van Google zelf.

## 5. Schermafbeeldingen (optioneel)

Alle beelden op de lespagina en de dia's zijn tekeningen in HTML en CSS (het zoekvak, de tabs met **Tools**, het menu van **Tools**, de
lijst met periodes, de resultaten met wat meetelt, de tabbladen van Chrome, de opdracht in Classroom). Echte schermafbeeldingen van
Google wisselen per zoekopdracht en per dag; ze zijn niet nodig. `assets/screenshots/LEESMIJ.md` houdt dat bij.

## 6. Het materiaal opnieuw maken

```bash
cd "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W12 - Les 06 - MWW - Zoeken als een onderzoeker/werkdocument" && python3 maak_werkdocumenten.py
```

Het script maakt ook de **verbetersleutel**: hetzelfde werkdocument met elk lichtblauw vak ingevuld (in het groen, met
wat niet telt in het grijs), in `werkdocument/sleutel/` (niet op GitHub). Zet ze daarna privé in je Drive (Mijn Drive ›
Verbetersleutels) en op de overzichtspagina (knop **Sleutel**):

```bash
cd "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica" && python3 _tools/zet_sleutel_klaar.py "2026-2027/W12 - Les 06 - MWW - Zoeken als een onderzoeker" && python3 _tools/maak_overzicht.py
```

## 7. Leerplandoelen in je jaaroverzicht

`lesdoelen.json` (klasgroep `3 MWW`): `BV2_13.04` (kern) · `BV2_02.05` (ondersteunend) · `BV2_06.19` (contextueel). De pre-push hook
werkt `Leerplandoelen 2026-2027.xlsx` bij.

## 8. Wat nog moet blijken in de klas

- **De knop met de datum:** heet ze in lokaal 18 ook **Elke periode** en staat **Afgelopen jaar** in de lijst? Google wijzigt zijn menu's;
  de lespagina tekent ze zoals gezien op 26-11-2026.
- **Het label van advertenties:** de lespagina en het werkdocument zeggen *advertenties (Gesponsord)*; bij de zoekopdrachten van de les zag
  ik geen tekstadvertentie. Zien de leerlingen het woord **Gesponsord**, of iets anders?
- **Het AI-overzicht:** het stond boven de resultaten van B en verdween met het periodefilter. Tellen de leerlingen het toch mee, en
  krijgen ze het ook op de schoolaccounts te zien?
- **De Engelse zoekopdracht** (Extra) en `OR` in kleine letters: niet nagekeken in Google (`_vakinhoud/informatievaardigheden.md`, *Nog na te kijken* 3).
- **Tijd:** past deel 2 (drie keer zoeken en tellen) in 14 minuten, en deel 3 in 10? Of is het zoeken van twee goede resultaten te veel?
- **Hetzelfde voorbeeld:** B en C staan als voorbeeld in de demo en bij *Toon stapjes*; kiezen de leerlingen daardoor allemaal dezelfde woordgroep?
- **Het webadres kopiëren:** lukt het uit de adresbalk (Ctrl + C), en plakken ze het adres van het resultaat en niet van Google?
- **De verbetertool:** kan ze de tabelvakken van deel 1 en 2 en de resultaten van deel 3 lezen uit het omgezette Google-document?
- **Niet goedgekeurd:** `_vakinhoud/informatievaardigheden.md` is een concept van 26-11-2026 (Jonas keurt het één keer goed).
- **De zelftest met het codewoord TREFWOORD en het filmpje Inleveren:** werkt het slot in lokaal 18 vlot en leveren alle leerlingen in?
