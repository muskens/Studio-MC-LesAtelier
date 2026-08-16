"""
Vercel serverless functie.

Neemt geplakte tekst aan, laat Claude er een slide-structuur (JSON) van maken
en bouwt daarmee een .pptx in de KW1C-huisstijl.

Vereist de omgevingsvariabele ANTHROPIC_API_KEY in Vercel.
"""

import json
import os
import re
import tempfile
from http.server import BaseHTTPRequestHandler

from anthropic import Anthropic
from api.build_kw1c_pptx import build


MODEL = "claude-haiku-4-5-20251001"

TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__),
    "kw1c_template.pptx",
)


SYSTEM_PROMPT = """Je zet losse tekst om naar een gestructureerde slide-opzet voor een
PowerPoint in de KW1C-huisstijl. Je krijgt brontekst van een docent.

Geef UITSLUITEND geldige JSON terug.
Geen uitleg, geen markdown en geen ``` tekens.

Structuur:
{"slides": [ ... ]}

Elke slide is een van deze types:

- {"type": "titelslide", "titel": "...", "ondertitel": "..."}
- {"type": "inhoudsopgave", "titel": "...", "items": ["...", "..."]}
- {"type": "sectie", "titel": "...", "ondertitel": "..."}
- {"type": "content", "variant": "blauw", "titel": "...", "bullets": ["...", "..."]}
- {"type": "content", "variant": "rood", "titel": "...", "bullets": ["...", "..."]}
- {"type": "quote", "titel": "...", "quote": "...", "bron": "..."}
- {"type": "tweekolommen", "titel": "...", "links": ["..."], "rechts": ["..."]}
- {"type": "afsluiting", "titel": "...", "tekst": "..."}

Regels:

- Begin altijd met een titelslide.
- Eindig altijd met een afsluiting-slide.
- Houd titels kort en krachtig.
- Gebruik maximaal 5 bullets per contentslide.
- Houd bullets bij voorkeur korter dan 8 woorden.
- Splits veel inhoud over meerdere slides.
- Gebruik sectie-slides voor grotere onderwerpen.
- Gebruik rood spaarzaam, maximaal ongeveer 1 op de 3 contentslides.
- Gebruik quote alleen voor een korte krachtige uitspraak.
- Gebruik tweekolommen alleen voor een echte vergelijking.
- Gebruik nooit het em-dash teken.
- Behoud de taal van de brontekst.
- Verzin geen feiten.
- Baseer je uitsluitend op de aangeleverde tekst.
"""


def _strip_to_json(text):
    """
    Haal een JSON-object uit de modelrespons,
    ook als er per ongeluk tekst omheen staat.
    """
    text = text.strip()

    text = re.sub(r"^```(?:json)?", "", text, flags=re.IGNORECASE).strip()
    text = re.sub(r"```$", "", text).strip()

    start = text.find("{")
    end = text.rfind("}")

    if start == -1 or end == -1:
        raise ValueError("Claude gaf geen geldig JSON-object terug.")

    text = text[start:end + 1]

    return json.loads(text)


def _slugify(naam, fallback="presentatie"):
    naam = (naam or "").strip().lower()
    naam = re.sub(r"[^a-z0-9]+", "-", naam).strip("-")
    return naam or fallback


def maak_pptx(
    tekst,
    onderwerp="",
    inhoudsopgave=True,
    contact="",
):
    client = Anthropic()

    titel_instructie = (
        onderwerp
        if onderwerp
        else "leid zelf een passende titel af uit de tekst"
    )

    instructie = (
        f"Onderwerp/titel van de presentatie: {titel_instructie}.\n"
        "Bepaal zelf een passend aantal slides op basis van de hoeveelheid "
        "en structuur van de brontekst. Maak niet onnodig veel slides en "
        "maak geen overvolle slides.\n"
    )

    if inhoudsopgave:
        instructie += (
            "Voeg na de titelslide een inhoudsopgave-slide toe.\n"
        )
    else:
        instructie += (
            "Voeg GEEN inhoudsopgave-slide toe.\n"
        )

    if contact:
        instructie += (
            "Zet op de afsluiting-slide deze contactgegevens als tekst: "
            f"{contact}.\n"
        )

    instructie += (
        '\nBrontekst:\n"""\n'
        + tekst.strip()
        + '\n"""'
    )

    response = client.messages.create(
        model=MODEL,
        max_tokens=4000,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": instructie,
            }
        ],
    )

    raw = "".join(
        block.text
        for block in response.content
        if getattr(block, "type", None) == "text"
    )

    content = _strip_to_json(raw)

    slides = content.get("slides", [])

    if not slides:
        raise ValueError("Geen slides ontvangen van Claude.")

    if not inhoudsopgave:
        slides = [
            slide
            for slide in slides
            if slide.get("type") != "inhoudsopgave"
        ]

    for slide in slides:
        if slide.get("type") == "titelslide":
            slide["contact"] = contact or ""
            break

    content["slides"] = slides

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(
            f"PowerPoint-template niet gevonden: {TEMPLATE_PATH}"
        )

    with tempfile.TemporaryDirectory() as tmp:
        content_path = os.path.join(tmp, "content.json")
        output_path = os.path.join(tmp, "out.pptx")

        with open(
            content_path,
            "w",
            encoding="utf-8",
        ) as file:
            json.dump(
                content,
                file,
                ensure_ascii=False,
                indent=2,
            )

        build(
            content_path,
            output_path,
            TEMPLATE_PATH,
        )

        if not os.path.exists(output_path):
            raise FileNotFoundError(
                "De PowerPoint is niet aangemaakt door build_kw1c_pptx.py."
            )

        with open(output_path, "rb") as file:
            data = file.read()

    return data, len(slides)


class handler(BaseHTTPRequestHandler):

    def _cors(self):
        self.send_header(
            "Access-Control-Allow-Origin",
            "*",
        )
        self.send_header(
            "Access-Control-Allow-Methods",
            "GET, POST, OPTIONS",
        )
        self.send_header(
            "Access-Control-Allow-Headers",
            "Content-Type",
        )

    def do_GET(self):
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/json; charset=utf-8",
        )
        self._cors()
        self.end_headers()

        response = {
            "status": "ok",
            "message": "Studio MC LesAtelier API werkt.",
        }

        self.wfile.write(
            json.dumps(
                response,
                ensure_ascii=False,
            ).encode("utf-8")
        )

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_POST(self):
        try:
            length = int(
                self.headers.get("Content-Length", 0)
            )

            body = (
                self.rfile.read(length)
                if length
                else b"{}"
            )

            payload = json.loads(
                body.decode("utf-8")
            )

            tekst = (
                payload.get("tekst") or ""
            ).strip()

            if not tekst:
                raise ValueError(
                    "Plak eerst tekst om een presentatie van te maken."
                )

            onderwerp = (
                payload.get("onderwerp") or ""
            ).strip()

            contact = (
                payload.get("contact") or ""
            ).strip()

            inhoudsopgave = True

            data, aantal_slides = maak_pptx(
                tekst=tekst,
                onderwerp=onderwerp,
                inhoudsopgave=inhoudsopgave,
                contact=contact,
            )

            bestandsnaam = (
                _slugify(onderwerp)
                + ".pptx"
            )

            self.send_response(200)

            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument."
                "presentationml.presentation",
            )

            self.send_header(
                "Content-Disposition",
                f'attachment; filename="{bestandsnaam}"',
            )

            self.send_header(
                "X-Slide-Count",
                str(aantal_slides),
            )

            self._cors()
            self.end_headers()

            self.wfile.write(data)

        except Exception as error:
            self.send_response(500)

            self.send_header(
                "Content-Type",
                "application/json; charset=utf-8",
            )

            self._cors()
            self.end_headers()

            fout = {
                "error": str(error),
                "type": type(error).__name__,
            }

            self.wfile.write(
                json.dumps(
                    fout,
                    ensure_ascii=False,
                ).encode("utf-8")
            )
