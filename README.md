# Studio MC LesAtelier

Een webtool waarmee collega's tekst plakken en er een PowerPoint van krijgen in de
KW1C-huisstijl. Geen Claude-account nodig voor de gebruiker. De huisstijl komt uit
het bestaande `kw1c-pptx`-buildscript en template.

## Twee modi

De tool heeft twee tabbladen:

- **Snel een presentatie**: plak tekst, krijg direct een PowerPoint in huisstijl.
- **Workshop ontwerpen**: upload een module plus opdracht en een aanpasbare prompt.
  Claude (Sonnet, met websearch) maakt een volledig workshopontwerp. Daaruit maak je
  altijd een PowerPoint, en optioneel een Word in de KW1C-huisstijl.

## Hoe het werkt

1. Collega plakt tekst en vult eventueel naam, e-mail en telefoon in. De omvang bepaalt
   de tool zelf op basis van de tekst, en er wordt altijd een inhoudsopgave-slide toegevoegd.
2. De server vraagt Claude om er een slide-structuur (JSON) van te maken volgens het
   kw1c-pptx-schema, inclusief de huisstijlregels (titels in hoofdletters, geen
   em-dash, variant rood spaarzaam).
3. Het bestaande `build_kw1c_pptx.py` plus `kw1c_template.pptx` maakt de .pptx.
4. De collega downloadt het bestand.

## Mappenstructuur

```
.
├── index.html              # de pagina (KW1C-huisstijl)
├── api/
│   ├── generate.py         # serverless functie: tekst -> JSON -> pptx
│   ├── build_kw1c_pptx.py  # jouw bestaande renderer
│   └── kw1c_template.pptx  # jouw huisstijl-template
├── requirements.txt
└── vercel.json
```

## Deployen op Vercel

1. Zet deze map in een Git-repo (GitHub) of gebruik de Vercel CLI.
2. Maak een nieuw project aan op vercel.com en koppel de repo.
3. Stel onder **Settings > Environment Variables** in:
   - Naam: `ANTHROPIC_API_KEY`
   - Waarde: jouw Anthropic API-sleutel (zie hieronder)
4. Deploy. Je krijgt een URL die je met collega's kunt delen.

## API-sleutel aanmaken

Dit staat los van je Claude-abonnement.

1. Ga naar console.anthropic.com en log in (of maak een account).
2. Voeg onder **Billing** een klein tegoed toe (een paar euro is genoeg voor honderden decks).
3. Ga naar **API Keys** en maak een nieuwe sleutel aan.
4. Plak die als `ANTHROPIC_API_KEY` in Vercel (stap 3 hierboven).

De sleutel staat alleen op de server. Collega's zien hem nooit.

## Aanpassen

- **Model**: in `api/generate.py` staat `MODEL`. Standaard het goedkope, snelle Haiku.
  Voor rijkere decks kun je er `claude-sonnet-4-6` van maken (iets duurder per deck).
- **Contactgegevens**: de velden naam, e-mail en telefoon vullen samen de balk onderaan
  de titelslide en de afsluiting. Laat ze leeg en er verschijnt geen contactbalk.
- **Slide-types en huisstijl**: die zitten in `build_kw1c_pptx.py` en het template,
  precies zoals in je skill.

## Lokaal testen (optioneel)

```bash
npm i -g vercel
vercel dev
```
Zet `ANTHROPIC_API_KEY` in een `.env`-bestand of in je shell.
