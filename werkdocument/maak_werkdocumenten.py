#!/usr/bin/env python3
"""
maak_werkdocumenten.py
Genereert het werkdocument voor les 6 "Zoeken als een onderzoeker" (Onderzoekslabo Jongeren & Welzijn, klas 3MWWE):

  Zoekverslag-onderzoekslabo.docx   het werkdocument dat de leerling INLEVERT (wordt in Classroom een Google-document,
                                    een kopie voor elke leerling)

Gebruik:  python3 maak_werkdocumenten.py     (vereist python-docx en Pillow; huisstijl.py en go-dalton-gent.png
          staan ernaast)

Geen zip en geen losse bestanden: alles wat de leerling maakt, typt hij IN dit werkdocument.

LET OP bij aanpassen:
 - Alle namen zijn verzonnen (het labo en Ines Coppens zijn fictief).
 - Het voorbeeld (rij gsm · smartphone en zoekopdracht A) staat al ingevuld, in het grijs. Laat het ingevuld.
 - Deel 1 heeft geen kolom *Idee* meer (Jonas, 26-11-2026: het idee was al het trefwoord): de leerling kiest zelf
   twee trefwoorden uit de vraag, naast gsm, en zoekt bij elk een synoniem.
 - Elk antwoord staat in een eigen vak, en elk vak hoort bij één criterium van lesvoorbereiding § 22
   (_afspraken/verbeteren.md, *Een les die je kan verbeteren*). Wijzig je de vakken, wijzig dan § 14 en § 22.
 - De punten in deel 4 (2 · 1 · 1 · 2 · 3 · 2 · 2 · 4 · 2 · 1 = 20) komen overeen met lesvoorbereiding § 22.
 - Het codewoord (TREFWOORD) staat bovenaan deel 4. De lespagina toont het pas als de zelftest juist is.
"""
import os

import huisstijl as hs

HERE = os.path.dirname(os.path.abspath(__file__))
UIT = os.path.join(HERE, "Zoekverslag-onderzoekslabo.docx")
B = hs.BREEDTE

VRAAG = "Slapen jongeren slechter als ze 's avonds op hun gsm zitten?"
ZOEKOPDRACHT_A = VRAAG

# De criteria van lesvoorbereiding § 22 (samen 20): 2 · 1 · 1 · 2 · 3 · 2 · 2 · 4 · 2 · 1
CRITERIA = [
    ("Deel 1: twee trefwoorden uit de vraag (naast gsm), elk met een synoniem (één woord per vak).", 2),
    ("Deel 1: vraag 2 zegt dat je met twee woorden ook pagina's vindt die het andere woord gebruiken.", 1),
    ("Deel 2: bij A, B en C staat een getal van 0 tot 5.", 1),
    ("Deel 2: zoekopdracht B bestaat uit trefwoorden (geen zin), met OR in hoofdletters tussen twee synoniemen.", 2),
    ("Deel 2: zoekopdracht C heeft een woordgroep tussen aanhalingstekens, een minteken vlak voor een woord, en de trefwoorden van B.", 3),
    ("Deel 2: vraag 2 zegt welk soort resultaat weg moest, en of het bij C weg was.", 2),
    ("Deel 3: vraag 1 en 2: het filter en de periode, en zoekopdracht D (B of C, met het filter).", 2),
    ("Deel 3: vraag 3: twee resultaten, elk met titel en een volledig webadres, over slaap en schermen.", 4),
    ("Deel 3: vraag 4 zegt waarom D beter werkt dan A, met iets wat je in de resultaten zag.", 2),
    ("Deel 4: de exitvraag is een zoekopdracht met trefwoorden en OR in hoofdletters.", 1),
]


def werkdocument():
    """In de huisstijl van 19-11-2026 (huisstijl.py): per deel de opdracht en de vragen met cijfers. De uitleg
    (de vroegere infokaders) staat sinds 26-11-2026 op de lespagina, in het blauwe kader van de stap (Jonas: "Ik zou ook
    wel graag de theoretische informatie toch terug in de lespagina willen verwerken, in plaats van in het
    werkdocument"). Waar de leerling klikt (Tools, het filter, de adresbalk), staat er ook."""
    doc = hs.nieuw_document()
    doc.core_properties.title = "Zoekverslag onderzoekslabo - werkdocument"
    doc.core_properties.author = "Toegepaste Informatica"
    hs.kopblok(doc, "ZOEKVERSLAG ONDERZOEKSLABO", punten=20)

    hs.kader(doc, "Een bericht van Ines Coppens, Onderzoekslabo Jongeren & Welzijn", [
        ("Opdracht:", "Ons labo start een onderzoek naar schermgebruik en slaap. Zoek wat er al over geschreven is, "
                      "en maak een kort zoekverslag."),
        ("Onderzoeksvraag:", VRAAG),
    ])

    # ---- Deel 1: van vraag naar trefwoorden
    hs.deel(doc, 1, "Van vraag naar trefwoorden")
    hs.opdracht(doc, "De vraag van Ines is te lang voor Google. Kies er de trefwoorden uit en zoek bij elk een synoniem. "
                     "In deel 2 typ je die woorden in Google. De eerste rij is een voorbeeld: je mag het overnemen.")
    hs.vraag(doc, "1.  Kies nog twee trefwoorden uit de vraag van Ines. Schrijf bij elk een synoniem. Eén woord per vak.")
    hs.tabel(doc, ["Trefwoord uit de vraag", "Synoniem"], [
        ("gsm", "smartphone"),
        ("", ""),
        ("", ""),
    ], [B / 2, B / 2], invul=(0, 1), voorbeeld_rij=1)
    hs.vraag(doc, "2.  Waarom zoek je met gsm én smartphone? Eén zin, kort mag.")
    hs.antwoordvak(doc, 1.6)

    # ---- Deel 2: drie keer zoeken
    hs.deel(doc, 2, "Drie keer zoeken", nieuwe_pagina=True)   # elk deel op een eigen blad (Jonas, 26-11-2026)
    p = hs.opdracht(doc, "Zoek drie keer in Google.")
    # Elke zin op een eigen regel (Jonas, in het Google-document, 26-11-2026).
    for zin in ("A is de vraag van Ines.", "Bij B typ je geen zin meer: alleen je drie trefwoorden uit deel 1, met OR.",
                "Bij C bouw je verder op B.",
                "Tel telkens hoeveel van de eerste 5 resultaten gaan over slaap én schermen én jongeren."):
        p.add_run().add_break()
        p.add_run(zin).bold = True
    hs.vraag(doc, "1.  Vul de tabel in: wat typte je, en hoeveel van de eerste 5 resultaten gaan over de vraag (0 tot 5)?")
    t = hs.tabel(doc, ["Zoekopdracht", "Wat typte je?", "0–5"], [
        ("A  De vraag van Ines, als zin", ZOEKOPDRACHT_A, ""),
        ("B  Geen zin: je drie trefwoorden uit deel 1, met OR tussen gsm en smartphone", "", ""),
        ("C  Zoekopdracht B + een woordgroep tussen aanhalingstekens en een woord met een minteken", "", ""),
    ], [4.6, B - 4.6 - 1.8, 1.8], invul=(1, 2), voorbeeld_rij=1, rijhoogte=1.5)
    # Bij C staan de woordgroep en het woord erbij, in het vet (Jonas, in het Google-document, 26-11-2026).
    p = t.rows[3].cells[0].paragraphs[0]
    for label, woord in (("Woordgroep: ", "Blauw licht"), ("Woord dat niet in de zoekresultaten mag zitten: ", "kopen")):
        p.add_run().add_break()
        p.add_run(label)
        p.add_run(woord).bold = True
    hs.vraag(doc, "2.  Welk soort resultaat van B wilde je kwijt met je minteken? Was het bij C weg? Twee korte zinnen.")
    hs.antwoordvak(doc, 2.0)

    # ---- Deel 3: filter en zoekverslag
    hs.deel(doc, 3, "Filter en zoekverslag", nieuwe_pagina=True)
    hs.opdracht(doc, "Voeg een filter toe aan je beste zoekopdracht (A,B of C) en kies twee bruikbare "
                     "resultaten.")   # (A,B of C): Jonas, in het Google-document, 26-11-2026
    hs.vraag(doc, "1.  Welk filter koos je, en welke periode? Kort mag.")
    hs.tabel(doc, ["Welk filter?", "Welke periode?"], [("", "")], [B / 2, B / 2], invul=(0, 1), rijhoogte=1.1)
    hs.vraag(doc, "2.  Zoekopdracht D: je beste zoekopdracht (B of C), met het filter. Kopieer wat je typte.")
    hs.antwoordvak(doc, 1.4)
    hs.vraag(doc, "3.  Kies twee resultaten die bruikbaar zijn voor het labo. Schrijf de titel en het webadres.")
    hs.tabel(doc, ["", "Titel", "Webadres"], [("Resultaat 1", "", ""), ("Resultaat 2", "", "")],
             [2.6, 6.6, B - 9.2], invul=(1, 2), rijhoogte=1.5)
    hs.vraag(doc, "4.  Waarom vond je met D meer bruikbare resultaten dan met A? Gebruik iets wat je zag. Twee zinnen.")
    hs.antwoordvak(doc, 2.4)

    # ---- Deel 4: klaar
    hs.deel(doc, 4, "Klaar?", nieuwe_pagina=True)
    hs.opdracht(doc, "Typ het codewoord. Kruis aan wat klopt. Beantwoord de exitvraag.")
    hs.vraag(doc, "Het codewoord van de zelftest (stap 4 op de lespagina):")
    hs.antwoordvak(doc, 0.9)
    hs.controletabel(doc, CRITERIA)
    hs.vraag(doc, "Exitvraag: een vriend zoekt informatie over stress bij jongeren tijdens de examens. "
                  "Schrijf één zoekopdracht met OR. Kort mag.")
    hs.antwoordvak(doc, 1.4)

    # ---- Extra
    hs.deel(doc, None, "Extra — niet verplicht", nieuwe_pagina=True)
    hs.opdracht(doc, "Zoek je beste zoekopdracht in het Engels. Tel opnieuw de eerste 5 resultaten. "
                     "Je mag dit voorbeeld overnemen: sleep teenagers phone OR smartphone")
    hs.vraag(doc, "1.  Wat typte je?")
    hs.antwoordvak(doc, 1.2)
    hs.vraag(doc, "2.  Hoeveel van de eerste 5 resultaten gaan over de vraag (0 tot 5)? Zijn het er meer of minder dan bij D?")
    hs.antwoordvak(doc, 1.6)

    doc.save(UIT)


# ---- De verbetersleutel: hetzelfde werkdocument, elk lichtblauw vak ingevuld met het verwachte antwoord
# (Jonas, 26-11-2026: "de verbetersleutel (het ingevulde document) ook linken op de lessenpagina").
# Uit lesvoorbereiding § 14 en § 22 (en README § 4). Het bestand komt in werkdocument/sleutel/ (niet op GitHub:
# .gitignore) en gaat met _tools/zet_sleutel_klaar.py als privé Google-document naar Jonas' Drive.
SLEUTEL = os.path.join(HERE, "sleutel", "Zoekverslag-onderzoekslabo - verbetersleutel.docx")
TELT_NIET = "  Telt niet: "
ANTWOORDEN = [
    # Deel 1, vraag 1 (2 punten: 0,5 per vak)
    "slapen (of: slaap; ook goed: 's avonds)",
    "nachtrust (ook: slaaptekort, slaapkwaliteit, sleep)",
    "jongeren",
    "tieners (ook: adolescenten, pubers, teenagers)",
    # Deel 1, vraag 2 (1 punt)
    "Zo vind je ook pagina's die het andere woord gebruiken (smartphone in plaats van gsm)."
    + TELT_NIET + "'omdat het hetzelfde is' zonder te zeggen wat je dan meer vindt.",
    # Deel 2, vraag 1: A (aantal), B, aantal, C, aantal (aantallen 1 punt samen; B 2; C 3)
    "bv. 1",
    "slapen jongeren gsm OR smartphone   (alleen trefwoorden: 1 · OR in hoofdletters tussen twee synoniemen: 1)",
    "bv. 3",
    'slapen jongeren gsm OR smartphone "blauw licht" -kopen   (woordgroep tussen aanhalingstekens: 1 · minteken vlak '
    "voor het woord, zonder spatie: 1 · de trefwoorden van B: 1)",
    "bv. 4",
    # Deel 2, vraag 2 (2 punten)
    "Ik wilde de webshops kwijt die gsm's verkopen (1). Bij C waren ze weg (1)."
    + TELT_NIET + "een woord uitsluiten zonder te zeggen welk resultaat van B weg moest.",
    # Deel 3, vraag 1 (1 punt)
    "Elke periode (de datum)",
    "Afgelopen jaar (elke periode is goed)",
    # Deel 3, vraag 2 (1 punt)
    'slapen jongeren gsm OR smartphone "blauw licht" -kopen   (B of C; het filter zie je niet in wat je typt)',
    # Deel 3, vraag 3 (4 punten: per resultaat titel 0,5, volledig webadres 1, over slaap én schermen 0,5)
    "de titel van een pagina over slaap en schermen (bv. van een krant, een gezondheidssite of een universiteit)",
    "https://… (het volledige adres uit de adresbalk; niet google.com/search?…, niet het AI-overzicht, geen advertentie)",
    "een tweede titel over slaap en schermen",
    "https://… (zoals hierboven)",
    # Deel 3, vraag 4 (2 punten)
    "A zocht op alle woorden van de zin, ook als, ze en hun, en gaf vooral forums en vragen van mensen (1). D zoekt op "
    "trefwoorden met een synoniem, sluit webshops uit en toont alleen pagina's van het afgelopen jaar: bij D gingen 4 "
    "van de 5 over de vraag, bij A maar 1 (1: iets wat de leerling zelf zag)."
    + TELT_NIET + "'omdat het beter is', of een algemene zin over Google zonder iets uit de eigen resultaten.",
    # Deel 4: codewoord (−2 als het leeg of fout is; een tikfout telt als juist)
    "TREFWOORD",
    # Deel 4: controletabel, kolom Klopt?
    *["x"] * 10,
    # Deel 4: exitvraag (1 punt)
    "stress examens jongeren OR tieners   (ook: stress OR spanning examens jongeren)"
    + TELT_NIET + "een zin; or in kleine letters.",
    # Extra (telt niet mee)
    "sleep teenagers phone OR smartphone",
    "een getal van 0 tot 5, en meer of minder dan bij D",
]


def sleutel():
    """Vult elk lichtblauw vak (HEX_VAK) van het werkdocument, in volgorde, met het verwachte antwoord."""
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import RGBColor
    doc = Document(UIT)
    vakken = [c for c in doc.element.body.iter(qn("w:tc"))
              if any(shd.get(qn("w:fill"), "").upper() == hs.HEX_VAK for shd in c.iter(qn("w:shd")))]
    if len(vakken) != len(ANTWOORDEN):
        raise SystemExit(f"{len(vakken)} vakken, maar {len(ANTWOORDEN)} antwoorden: pas ANTWOORDEN aan.")
    from docx.table import _Cell
    groen = RGBColor(0x1B, 0x6E, 0x2F)
    for tc, antwoord in zip(vakken, ANTWOORDEN):
        cel = _Cell(tc, None)
        p = cel.paragraphs[0]
        hoofd, _, niet = antwoord.partition(TELT_NIET)
        r = p.add_run(hoofd)
        hs.stijl(r, grootte=10, vet=True, kleur=groen)
        if niet:
            r = p.add_run("  Telt niet: " + niet)
            hs.stijl(r, grootte=9, cursief=True, kleur=hs.GRIJS)
    # bovenaan, vóór alles: een duidelijke titel
    eerste = doc.paragraphs[0]
    titel = eerste.insert_paragraph_before()
    r = titel.add_run("VERBETERSLEUTEL — alleen voor de leraar, niet delen met leerlingen")
    hs.stijl(r, grootte=12, vet=True, kleur=RGBColor(0xB0, 0x1E, 0x1E))
    doc.core_properties.title = "Zoekverslag onderzoekslabo - verbetersleutel"
    os.makedirs(os.path.dirname(SLEUTEL), exist_ok=True)
    doc.save(SLEUTEL)


if __name__ == "__main__":
    werkdocument()
    print("Gemaakt: Zoekverslag-onderzoekslabo.docx")
    sleutel()
    print("Gemaakt: sleutel/Zoekverslag-onderzoekslabo - verbetersleutel.docx (niet op GitHub)")
