# sovereigntypath.org

Public architecture surface for The Sovereignty Path.

## Stack

Plain HTML + CSS. No framework, no build step. Crimson Pro and JetBrains Mono via Google Fonts. Vesica Piscis as the single visual motif. Palette A (Parchment on Oxblood, dark mode).

## Structure

- `index.html` — architecture entry
- `governance.html` — boundary statement, classification system
- `glossary.html` — controlled vocabulary with operational definition links
- `author.html` — Dean Hobson, the seat
- `bundle/` — the AI-Adjacent Doctrine Bundle (Manifesto, Companion, Operational Definitions, Material Architecture)
- `api/v1/` — static JSON consultation API for public AI-readable access
- `llms.txt` — AI discovery file for public retrieval/citation guidance
- `.well-known/ai-readme.json` — public AI-readable site metadata
- `.well-known/ai-plugin.json` — read-only OpenAPI adapter metadata for compatible agents
- `assets/style.css` — single stylesheet
- `CNAME` — Cloudflare Pages / GitHub Pages custom domain
- `.nojekyll` — skip Jekyll processing

## Bundle integrity

The Manifesto is served only with the Red-Team Companion linked at the top. Both the Manifesto template and the Bundle landing enforce this. Build Spec v3 §15 makes Companion-bundling a hard acceptance test.

## Deploy via Cloudflare Pages

1. In Cloudflare dashboard → Workers & Pages → Create → Pages → Connect to Git
2. Select the `sovereigntypath-org` repository
3. Build settings: leave **build command** empty, **build output directory** = `/`
4. Save and deploy
5. Custom domains → Add → `sovereigntypath.org` (Cloudflare proxies the existing DNS)

For GitHub Pages instead: Settings → Pages → Source = main branch → Custom domain = sovereigntypath.org. The `CNAME` and `.nojekyll` files are already in place.

## Local preview

```
python -m http.server 4400 --directory .
```

Open http://127.0.0.1:4400/

## License & authority

All content is authored by Dean Hobson. Governing substrate: Canon v3.1.

## Public AI boundary

The site is intentionally readable by AI systems as a public surface. Public retrieval, search, citation, and consultation are allowed when canonical URLs, boundary statements, and bundle integrity are preserved. Training/fine-tuning permission, private document access, local file access, credentials, and write access are not granted.
