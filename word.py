"""
Vercel serverless functie: zet de (bewerkbare) workshoptekst in Markdown om naar
een Word-document in de KW1C-huisstijl. Volledig deterministisch, geen AI-aanroep.
Hergebruikt het build-script en de template van de kw1c-huisstijl skill.
"""

import json
import re
from pathlib import Path
from http.server import BaseHTTPRequestHandler

from build_kw1c_report import build_report

TEMPLATE = Path(__file__).resolve().parent / "kw1c_report_template.docx"


def _clean_heading(t):
    return re.sub(r"[*_`]+", "", t).strip()


def md_to_content(md, doc_title="Workshop", titel=None):
    lines = md.split("\n")
    sections = []
    cur = None
    para = []
    bullets = []
    numbered = []

    def flush_para():
        nonlocal para
        if para and cur is not None:
            cur["blocks"].append({"type": "paragraph", "text": " ".join(para).strip()})
        para = []

    def flush_bullets():
        nonlocal bullets
        if bullets and cur is not None:
            cur["blocks"].append({"type": "bullets", "items": bullets[:]})
        bullets = []

    def flush_numbered():
        nonlocal numbered
        if numbered and cur is not None:
            cur["blocks"].append({"type": "numbered", "items": numbered[:]})
        numbered = []

    def flush_all():
        flush_para()
        flush_bullets()
        flush_numbered()

    def ensure_section():
        nonlocal cur
        if cur is None:
            cur = {"heading": "Workshop", "level": 1, "blocks": []}
            sections.append(cur)

    for ln in lines:
        s = ln.rstrip()

        if re.match(r"^#\s+", s):
            flush_all()
            cur = {"heading": _clean_heading(re.sub(r"^#\s+", "", s)), "level": 1, "blocks": []}
            sections.append(cur)
            continue
        if re.match(r"^##\s+", s):
            flush_all()
            ensure_section()
            cur["blocks"].append({"type": "heading", "level": 2, "text": _clean_heading(re.sub(r"^##\s+", "", s))})
            continue
        if re.match(r"^###+\s+", s):
            flush_all()
            ensure_section()
            cur["blocks"].append({"type": "heading", "level": 3, "text": _clean_heading(re.sub(r"^###+\s+", "", s))})
            continue

        mb = re.match(r"^\s*[-*]\s+(.*)", s)
        if mb:
            flush_para()
            flush_numbered()
            ensure_section()
            bullets.append(mb.group(1).strip())
            continue

        mn = re.match(r"^\s*\d+[.)]\s+(.*)", s)
        if mn:
            flush_para()
            flush_bullets()
            ensure_section()
            numbered.append(mn.group(1).strip())
            continue

        if s.strip() == "":
            flush_all()
            continue

        flush_bullets()
        flush_numbered()
        ensure_section()
        para.append(s.strip())

    flush_all()

    if not sections:
        sections = [{"heading": "Workshop", "level": 1,
                     "blocks": [{"type": "paragraph", "text": md.strip()}]}]

    content = {
        "doc_title": (doc_title or "Workshop")[:50],
        "sections": sections,
        "toc": len(sections) > 3,
    }
    if titel:
        content["title_page"] = {
            "title": titel.upper(),
            "subtitle": "Workshopontwerp",
            "author": "Studio MC",
            "version": "",
        }
    return content


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
                raise ValueError("Er is geen tekst om een Word van te maken.")
            titel = payload.get("titel", "") or ""
            doc_title = payload.get("doc_title") or (titel or "Workshop")

            content = md_to_content(markdown, doc_title=doc_title, titel=titel or None)

            import tempfile, os
            with tempfile.TemporaryDirectory() as tmp:
                cpath = os.path.join(tmp, "content.json")
                opath = os.path.join(tmp, "out.docx")
                with open(cpath, "w", encoding="utf-8") as f:
                    json.dump(content, f, ensure_ascii=False)
                build_report(content, TEMPLATE, Path(opath))
                with open(opath, "rb") as f:
                    data = f.read()

            naam = _slug(titel) + ".docx"
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{naam}"')
            self._cors()
            self.end_headers()
            self.wfile.write(data)

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self._cors()
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
