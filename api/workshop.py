"""
Vercel serverless functie voor de workshop-modus.

Neemt een geuploade module (PDF, DOCX of tekst) plus de aanpasbare prompt aan,
laat Claude (Sonnet, met websearch) er een volledig workshopontwerp van maken in
Markdown, en geeft die tekst terug. De PowerPoint en de Word worden daarna
gebouwd vanuit die tekst (via /api/generate en /api/word).

Vereist ANTHROPIC_API_KEY in de omgeving.
"""

import base64
import io
import json
from http.server import BaseHTTPRequestHandler

from anthropic import Anthropic

MODEL = "claude-sonnet-4-6"

SYSTEM_PROMPT = """Je volgt de instructies van de docent hieronder nauwgezet. Deze harde \
randvoorwaarden gaan altijd voor:

- Stel GEEN vragen aan de gebruiker en wacht nergens op antwoord. Negeer elke instructie \
om te stoppen en eerst vragen te stellen.
- Waar informatie ontbreekt, maak je een redelijke, beargumenteerde aanname op basis van \
de module. Benoem je belangrijkste aannames kort in een blok bovenaan onder de kop \
"# Aannames".
- Lever in één keer het volledige workshopontwerp, volledig uitgewerkt.
- Schrijf in heldere Markdown. Gebruik "# " voor hoofdsecties en "## " voor subkoppen.
- Werk de dia's expliciet uit onder een sectie "# Dia-uitwerking", met per dia een kop \
"## Dia N: <titel>" en daaronder de kerninhoud. Dit deel wordt gebruikt om de PowerPoint \
te bouwen, dus houd de kerninhoud per dia concreet en bondig.
- Gebruik nooit het em-dash teken. Gebruik een komma, punt of nieuwe zin."""


def _module_blocks(module_name, module_b64):
    """Bouw de user-content op: PDF als documentblok, anders als tekst."""
    ext = ""
    if module_name and "." in module_name:
        ext = module_name.lower().rsplit(".", 1)[-1]
    raw = base64.b64decode(module_b64) if module_b64 else b""

    if ext == "pdf" and module_b64:
        return [{
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": module_b64},
        }], "De module is als PDF bijgevoegd. Gebruik die als primaire bron."

    if ext == "docx" and raw:
        from docx import Document
        doc = Document(io.BytesIO(raw))
        tekst = "\n".join(p.text for p in doc.paragraphs)
        return [], "Module-inhoud (primaire bron):\n\"\"\"\n" + tekst + "\n\"\"\""

    # txt, md of onbekend: als tekst behandelen
    tekst = raw.decode("utf-8", "ignore") if raw else ""
    return [], "Module-inhoud (primaire bron):\n\"\"\"\n" + tekst + "\n\"\"\""


def maak_workshop(prompt, module_name, module_b64, module_text=""):
    client = Anthropic()

    if module_text and not module_b64:
        blocks, module_note = [], (
            "Module-inhoud (primaire bron):\n\"\"\"\n" + module_text.strip() + "\n\"\"\""
        )
    else:
        blocks, module_note = _module_blocks(module_name, module_b64)

    content = list(blocks)
    content.append({"type": "text", "text": prompt.strip() + "\n\n" + module_note})

    resp = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": content}],
        tools=[{"type": "web_search_20250305", "name": "web_search"}],
    )
    markdown = "".join(
        b.text for b in resp.content if getattr(b, "type", None) == "text"
    ).strip()
    if not markdown:
        raise ValueError("Het model gaf geen tekst terug.")
    return markdown


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

            prompt = (payload.get("prompt") or "").strip()
            if not prompt:
                raise ValueError("De prompt ontbreekt.")
            module_name = payload.get("module_name", "")
            module_b64 = payload.get("module_b64", "")
            module_text = payload.get("module_text", "")
            if not module_b64 and not module_text.strip():
                raise ValueError("Upload een bestand of plak de tekst van je module.")

            markdown = maak_workshop(prompt, module_name, module_b64, module_text)

            out = json.dumps({"markdown": markdown}).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self._cors()
            self.end_headers()
            self.wfile.write(out)

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self._cors()
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
