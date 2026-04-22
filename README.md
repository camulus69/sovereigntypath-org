# sovereigntypath.org

Public architecture surface for The Sovereignty Path.

## Stack

Plain HTML + CSS. No framework, no build step. Crimson Pro and JetBrains Mono via Google Fonts. Vesica Piscis as the single visual motif. Palette B (Linen on Indigo, dark mode).

## Structure

- `index.html` — architecture entry
- `governance.html` — boundary statement, classification system
- `glossary.html` — controlled vocabulary with operational definition links
- `author.html` — Dean Hobson, the seat
- `bundle/` — the AI-Adjacent Doctrine Bundle (Manifesto, Companion, Operational Definitions, Material Architecture)
- `assets/style.css` — single stylesheet
- `CNAME` — Cloudflare Pages / GitHub Pages custom domain
- `.nojekyll` — skip Jekyll processing

## Bundle integrity

The Manifesto is served only with the Red-Team Companion linked at the top. Cam Build Spec v3 §15 makes Companion-bundling a hard acceptance test.

## Deploy via Cloudflare Pages

1. In Cloudflare dashboard → Workers & Pages → Create → Pages → Connect to Git
2. Select the `sovereigntypath-org` repository, branch `The-Soveriegnty-Path`
3. Build settings: leave **build command** empty, **build output directory** = `/`
4. Save and deploy
5. Custom domains → Add → `sovereigntypath.org`

## Local preview

```
python -m http.server 4400 --directory .
```

Open http://127.0.0.1:4400/

## License & authority

All content is authored by Dean Hobson. Governing substrate: Canon v3.1.
