# Build modes and the manifest — the localhost-revert trap

## The short version

`npm run build:dev` silently rewrites `dist/manifest.xml` back to
**localhost URLs**, even if a production manifest was already sideloaded.
If you run it after deploying/sideloading for real use (a pilot, a demo,
anything outside your own dev machine), the next Excel launch will try to
load the task pane from `https://localhost:3000` and fail or show stale
content.

**Rule of thumb: once `dist/manifest.xml` is sideloaded anywhere other than
your own dev machine, only ever rebuild it with `npm run build` (production
mode). Never `npm run build:dev` after that point — not even to sanity-check
an unrelated JS change.**

## Why this happens

`webpack.config.js`'s `CopyWebpackPlugin` config for `manifest*.xml` does
the dev→prod URL swap only when the build is NOT in dev mode:

```js
transform(content) {
  if (dev) {
    return content;
  } else {
    return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
  }
},
```

`npm run build:dev` sets `dev = true`, so this branch copies
`manifest.xml` through **unmodified** — i.e. with whatever URLs are in the
source `manifest.xml` at the repo root, which are the localhost dev-server
ones (`urlDev = "https://localhost:3000"`). `npm run build` (production
mode) takes the other branch and rewrites every occurrence of the dev
origin to the production one (`urlProd`, the stable Vercel alias).

This is a `dist/`-only file — `dist/` is gitignored, so this never shows up
as a git diff. The only way to notice it happened is to actually check
`dist/manifest.xml`'s contents (or watch Excel fail to load the add-in).

## How this bit us once already

Same bug class as an earlier incident: `dist/manifest.xml` had reverted to
localhost URLs after a `build:dev` run done purely to sanity-check unrelated
JS changes, while a production manifest was still meant to be sideloaded for
pilot testing. Found only by explicitly checking "which manifest is
currently sideloaded" and reading `dist/manifest.xml`'s actual URLs — not
from any build error or git diff, since nothing about the build fails and
nothing tracked in git changes.

## How to check whether this is currently a problem

```powershell
# What's actually in dist/manifest.xml right now?
Select-String -Path dist\manifest.xml -Pattern "https://"

# What manifest does Excel currently have registered, and where does it point?
Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"
```

If `dist/manifest.xml` shows `localhost` and you expected it to be serving
production, rebuild with `npm run build` (not `build:dev`) and re-sideload
if needed.

## Quick reference

| Command | Mode | Manifest URLs written to `dist/manifest.xml` |
|---|---|---|
| `npm run build:dev` | development | Unmodified — whatever's in the source `manifest.xml` (localhost) |
| `npm run build` | production | Dev origin replaced with the production Vercel alias |
| `npm run watch` | development | Same as `build:dev` — also unsafe post-sideload |
| `npm run dev-server` | development (webpack-dev-server) | Doesn't write `dist/` at all — irrelevant to this issue |
