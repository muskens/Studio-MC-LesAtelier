"""
Vercel serverless functie: bouwt een PowerPoint uit de workshoptekst, waarbij
ELKE '## Dia N' kop precies één slide wordt, in dezelfde volgorde. Eromheen komen
een titelslide, een inhoudsopgave en een afsluiting.

De korte bullets per dia worden door een goedkoop model (Haiku) uit de dia-tekst
gehaald. Lukt dat niet, dan valt de functie terug op een simpele tekstsplitsing.
Als er geen '## Dia' koppen in de tekst staan, valt de functie terug op de vrije
indeling van /api/generate.

Vereist ANTHROPIC_API_KEY in de omgeving.
"""

import json
import os
import re
import tempfile
from http.server import BaseHTTPRequestHandler

from anthropic import Anthropic
from build_kw1c_pptx import build

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "kw1c_template.pptx")
BULLET_MODEL = "claude-haiku-4-5-20251001"

HEADING_RE = re.compile(r"(?m)^#{1,6}\s")
DIA_RE = re.compile(r"(?mi)^##\s+Dia\s*\d+\s*[:.\-]?\s*(.*)$")


def parse_dia_blocks(md):
    """Geef een lijst (titel, tekst) terug, één per '## Dia N' kop."""
    heads = [m.start() for m in HEADING_RE.finditer(md)]
    blocks = []
    for i, m in enumerate(DIA_RE.finditer(md)):
        titel = (m.group(1) or "").strip() or f"Dia {i + 1}"
        start = m.end()
        latere_heads = [h for h in heads if h > m.start()]
        end = min(latere_heads) if latere_heads else len(md)
        tekst = md[start:end].strip()
        blocks.append((titel, tekst))
    return blocks


def _fallback_bullets(tekst):
    """Simpele terugval: eerste paar zinnen als korte bullets."""
    tekst = re.sub(r"(?i)\b(doel|kerninhoud|visualisatie|vragen|koppeling|tijd)\s*:\s*", "", tekst)
    stukken = re.split(r"(?<=[.!?])\s+", tekst.replace("\n", " "))
    bullets = []
    for s in stukken:
        s = s.strip(" -*").strip()
        if len(s) < 3:
            continue
        woorden = s.split()
        if len(woorden) > 12:
            s = " ".join(woorden[:12])
        bullets.append(s)
        if len(bullets) >= 4:
            break
    return bullets or ["Zie het uitgewerkte document"]


def bullets_per_dia(blocks):
    """Vraag Haiku om korte bullets per dia. Bij twijfel: terugval per dia."""
    try:
        client = Anthropic()
        payload = [{"titel": t, "beschrijving": b} for t, b in blocks]
        system = (
            "Je krijgt dia-beschrijvingen uit een workshopontwerp. Geef per dia 2 tot 5 "
            "korte bullets (elk maximaal ongeveer 8 woorden) met alleen de kerninhoud die "
            "op de dia hoort. Gebruik nooit het em-dash teken. Antwoord met UITSLUITEND "
            "geldige JSON: een lijst van lijsten met strings, in dezelfde volgorde als de "
            "invoer. Geen uitleg, geen markdown."
        )
        resp = client.messages.create(
            model=BULLET_MODEL,
            max_tokens=2000,
            system=system,
            messages=[{"role": "user", "content": json.dumps(payload, ensure_ascii=False)}],
        )
        raw = "".join(b.text for b in resp.content if getattr(b, "type", None) == "text").strip()
        raw = re.sub(r"^```(?:json)?|```$", "", raw).strip()
        data = json.loads(raw)
        if isinstance(data, list) and len(data) == len(blocks):
            schoon = []
            for item in data:
                if isinstance(item, list) and item:
                    schoon.append([str(x).strip() for x in item][:5])
                else:
                    schoon.append(None)
            return [
                s if s else _fallback_bullets(blocks[i][1])
                for i, s in enumerate(schoon)
            ]
    except Exception:
        pass
    return [_fallback_bullets(b) for _, b in blocks]


def build_from_dias(markdown, onderwerp="", contact=""):
    blocks = parse_dia_blocks(markdown)
    if not blocks:
        return None  # geen dia-koppen; laat de aanroeper terugvallen

    bullets = bullets_per_dia(blocks)
    dia_titels = [t for t, _ in blocks]

    slides = [
        {"type": "titelslide", "titel": onderwerp or "Workshop",
         "ondertitel": "Workshop", "contact": contact or ""},
        {"type": "inhoudsopgave", "titel": "Programma", "items": dia_titels},
    ]
    for (titel, _), bl in zip(blocks, bullets):
        slides.append({"type": "content", "variant": "blauw", "titel": titel, "bullets": bl})
    slides.append({"type": "afsluiting", "titel": "Aan de slag", "tekst": contact or ""})

    content = {"slides": slides}
    with tempfile.TemporaryDirectory() as tmp:
        cpath = os.path.join(tmp, "content.json")
        opath = os.path.join(tmp, "out.pptx")
        with open(cpath, "w", encoding="utf-8") as f:
            json.dump(content, f, ensure_ascii=False)
        build(cpath, opath, TEMPLATE_PATH)
        with open(opath, "rb") as f:
            return f.read(), len(slides)


def _slug(naam, fallback="workshop"):
    naam = (naam or "").strip().lower()
    naam = re.sub(r"[^a-z0-9]+", "-", naam).strip("-")
    return naam or fallback


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

            markdown = (payload.get("markdown") or "").strip()
            if not markdown:
                raise ValueError("Er is geen workshoptekst om een presentatie van te maken.")
            onderwerp = payload.get("onderwerp", "")
            contact = payload.get("contact", "")

            result = build_from_dias(markdown, onderwerp, contact)
            if result is None:
                # Geen dia-koppen gevonden: vrije indeling via generate.py
                from generate import maak_pptx
                data, n = maak_pptx(markdown, onderwerp, inhoudsopgave=True, contact=contact)
            else:
                data, n = result

            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{_slug(onderwerp)}.pptx"')
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
