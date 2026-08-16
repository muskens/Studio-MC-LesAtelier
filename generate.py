"""
Vercel serverless functie.

Neemt geplakte tekst aan, laat Claude er een slide-structuur (JSON) van maken
in het schema van de kw1c-pptx skill, en bouwt daarmee een .pptx in de
KW1C-huisstijl met jouw bestaande buildscript en template.

Vereist de omgevingsvariabele ANTHROPIC_API_KEY (instellen in Vercel).
"""

import json
import os
import re
import tempfile
from http.server import BaseHTTPRequestHandler

from anthropic import Anthropic
from build_kw1c_pptx import build

# Goedkoop en snel. Zet op "claude-sonnet-4-6" als je rijkere decks wilt.
MODEL = "claude-haiku-4-5-20251001"

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "kw1c_template.pptx")

SYSTEM_PROMPT = """Je zet losse tekst om naar een gestructureerde slide-opzet voor een \
PowerPoint in de KW1C-huisstijl. Je krijgt brontekst van een docent.

Geef UITSLUITEND geldige JSON terug. Geen uitleg, geen markdown, geen ``` tekens.

Structuur: {"slides": [ ... ]}. Elke slide is een van deze types:

- {"type": "titelslide", "titel": "...", "ondertitel": "..."}  (altijd de eerste slide)
- {"type": "inhoudsopgave", "titel": "...", "items": ["...", "..."]}
- {"type": "sectie", "titel": "...", "ondertitel": "..."}  (leidt een nieuw onderwerp in)
- {"type": "content", "variant": "blauw", "titel": "...", "bullets": ["...", "..."]}
- {"type": "content", "variant": "rood", "titel": "...", "bullets": ["...", "..."]}
- {"type": "quote", "titel": "...", "quote": "...", "bron": "..."}
- {"type": "tweekolommen", "titel": "...", "links": ["..."], "rechts": ["..."]}
- {"type": "afsluiting", "titel": "...", "tekst": "..."}  (altijd de laatste slide)

Regels:
- Begin met een titelslide, eindig met een afsluiting-slide.
- Titels kort en krachtig.
- Bullets bondig (richtlijn: max 8 woorden), max 5 bullets per contentslide.
  Splits lange tekst over meerdere contentslides in plaats van volle slides.
- Gebruik sectie-slides om grotere onderwerpen in te leiden.
- Gebruik variant "rood" spaarzaam (hooguit 1 op de 3 contentslides), voor nadruk,
  risico's of contrast. De rest is "blauw".
- Quote-slide alleen bij een korte, krachtige uitspraak (max 2 zinnen).
- Tweekolommen alleen bij een echte vergelijking of voor/na.
- Gebruik NOOIT het em-dash teken. Gebruik een komma, een punt of een nieuwe bullet.
- Behoud de taal van de brontekst (meestal Nederlands).
- Verzin geen feiten. Baseer je uitsluitend op de aangeleverde tekst."""


def _strip_to_json(text):
    """Haal een JSON-object uit de modelrespons, ook als er per ongeluk tekst omheen staat."""
    text = text.strip()
    text = re.sub(r"^```(?:json)?", "", text).strip()
    text = re.sub(r"```$", "", text).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1:
        text = text[start:end + 1]
    return json.loads(text)


def _slugify(naam, fallback="presentatie"):
    naam = (naam or "").strip().lower()
    naam = re.sub(r"[^a-z0-9]+", "-", naam).strip("-")
    return naam or fallback


def maak_pptx(tekst, onderwerp="", inhoudsopgave=False, contact=""):
    client = Anthropic()  # leest ANTHROPIC_API_KEY uit de omgeving

    instructie = (
        f"Onderwerp/titel van de presentatie: {onderwerp or 'leid zelf een passende titel af uit de tekst'}.\n"
        "Bepaal zelf een passend aantal slides op basis van de hoeveelheid en structuur "
        "van de brontekst. Maak niet onnodig veel slides en geen overvolle slides.\n"
        f"{'Voeg na de titelslide een inhoudsopgave-slide toe.' if inhoudsopgave else 'Voeg GEEN inhoudsopgave-slide toe.'}\n"
    )
    if contact:
        instructie += f"Zet op de afsluiting-slide deze contactgegevens als tekst: {contact}.\n"

    instructie += "\nBrontekst:\n\"\"\"\n" + tekst.strip() + "\n\"\"\""

    resp = client.messages.create(
        model=MODEL,
        max_tokens=4000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": instructie}],
    )
    raw = "".join(block.text for block in resp.content if block.type == "text")
    content = _strip_to_json(raw)

    slides = content.get("slides", [])
    if not slides:
        raise ValueError("Geen slides ontvangen van het model.")

    # Toggle inhoudsopgave hard afdwingen
    if not inhoudsopgave:
        slides = [s for s in slides if s.get("type") != "inhoudsopgave"]

    # Contact op de titelslide injecteren (lege string = geen contactbalk),
    # zodat het deck niet automatisch Marja's gegevens toont.
    for s in slides:
        if s.get("type") == "titelslide":
            s["contact"] = contact or ""
            break

    content["slides"] = slides

    with tempfile.TemporaryDirectory() as tmp:
        content_path = os.path.join(tmp, "content.json")
        output_path = os.path.join(tmp, "out.pptx")
        with open(content_path, "w", encoding="utf-8") as f:
            json.dump(content, f, ensure_ascii=False)
        build(content_path, output_path, TEMPLATE_PATH)
        with open(output_path, "rb") as f:
            data = f.read()

    return data, len(slides)


class handler(BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length) if length else b"{}"
            payload = json.loads(body.decode("utf-8"))

            tekst = (payload.get("tekst") or "").strip()
            if not tekst:
                raise ValueError("Plak eerst wat tekst om een presentatie van te maken.")

            onderwerp = payload.get("onderwerp", "")
            inhoudsopgave = True  # altijd een inhoudsopgave-slide
            contact = payload.get("contact", "")

            data, n = maak_pptx(tekst, onderwerp, inhoudsopgave, contact)

            bestandsnaam = _slugify(onderwerp) + ".pptx"
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            self.send_header(
                "Content-Disposition", f'attachment; filename="{bestandsnaam}"'
            )
            self.send_header("X-Slide-Count", str(n))
            self._cors()
            self.end_headers()
            self.wfile.write(data)

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self._cors()
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
