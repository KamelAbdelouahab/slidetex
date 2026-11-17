# SlideTeX — équations LaTeX pour PowerPoint 2016+

Add-in PowerPoint moderne (Office.js) dans l'esprit de TeX4PPT : tapez du LaTeX, insérez une
équation **vectorielle (SVG)**, re-cliquez dessus plus tard pour la **modifier**. Fonctionne sur
PowerPoint 2016 / 2019 / 2021 / 2024, Microsoft 365 (Windows & Mac) et PowerPoint sur le web, avec
des fichiers `.pptx` standard (rien à installer chez les personnes qui ouvrent la présentation).

## Fonctionnalités

- **Deux moteurs**
  - *MathJax intégré* : aucune installation, hors ligne. AMS, `aligned`, `cases`, matrices, mhchem (`\ce{}`), physics, cancel…
  - *LaTeX local* : votre vrai LaTeX (TeX Live ou MiKTeX + `dvisvgm`) → TikZ, pgfplots, chemfig, tout paquet.
- **Édition en place** : sélectionnez une équation sur la diapo → la source se recharge → *Mettre à jour*.
  Position, redimensionnement manuel et ordre d'empilement sont conservés.
- Source LaTeX stockée **dans le fichier .pptx** (tag de la forme) + texte alternatif (accessibilité).
- Aperçu en direct, palette de symboles, exemples, historique, préambule / macros (`\newcommand`…).
- Taille en points (cohérente avec le texte de la diapo), couleur, mode display / inline / texte.
- SVG vectoriel (recolorable, « Convertir en forme » possible) ou PNG 150–600 dpi.
- **Régénérer toutes les équations** de la présentation (ex. changer couleur/taille d'un coup).

| Version | Insertion | Édition |
|---|---|---|
| Microsoft 365, PowerPoint 2021/2024, Mac, Web | SVG ou PNG | ✅ (PowerPointApi 1.5+) |
| PowerPoint 2016 / 2019 perpétuel (Windows, webview IE11) | PNG (SVG si dispo) | ❌ — source gardée dans les paramètres du document |

## Installation la plus simple (Windows, sans Node.js ni npm)

1. Téléchargez `SlideTeX-Windows.zip`, extrayez `SlideTeX.exe`, double-cliquez.
   (SmartScreen « Windows a protégé votre ordinateur » → *Informations complémentaires* → *Exécuter quand même* : l'exe n'est pas signé.)
2. Acceptez l'avertissement « Installer ce certificat ? » → **Oui** (certificat HTTPS local, limité à `localhost`).
3. Fermez / rouvrez PowerPoint → **Accueil › LaTeX › Équation LaTeX**.

Tout est installé pour l'utilisateur courant dans `%LOCALAPPDATA%\SlideTeX` (pas de droits admin) ; le
petit serveur local démarre automatiquement (invisible) à l'ouverture de session. Désinstallation :
Paramètres › Applications › SlideTeX, ou `SlideTeX.exe --uninstall`. État : `SlideTeX.exe --status`.

Pour reconstruire l'exe : `npm i -g bun` puis `npm run build:exe` (→ `dist-exe/SlideTeX.exe`).

## Installation développeur (Node.js)

Prérequis : **Node.js 18+** (https://nodejs.org), PowerPoint. Optionnel : **MiKTeX** ou **TeX Live**
pour le moteur LaTeX local (avec MiKTeX, réglez « Install missing packages » sur *Always*).

```bash
npm install
```

```bash
npm start
```

`npm start` :
1. crée et installe un certificat HTTPS local (accepter la fenêtre de sécurité Windows la 1ʳᵉ fois),
2. lance le serveur `https://localhost:3000`,
3. enregistre l'add-in et ouvre PowerPoint.

Dans PowerPoint : onglet **Accueil › groupe LaTeX › Équation LaTeX**.
Pour arrêter : `npm run stop`. Ou double-cliquez sur `start-windows.cmd`.

### Sideload manuel (si `npm start` n'ouvre pas PowerPoint)

- **Windows (dossier partagé)** : partagez le dossier contenant `manifest.xml`, puis PowerPoint ›
  Fichier › Options › Centre de gestion de la confidentialité › Paramètres › Catalogues de compléments
  approuvés → ajoutez `\\VOTRE-PC\partage`, cochez *Afficher dans le menu*. Redémarrez PowerPoint, puis
  Insertion › Mes compléments › DOSSIER PARTAGÉ › SlideTeX. Le serveur doit tourner (`npm run serve`).
- **Mac** : copiez `manifest.xml` dans
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (créez `wef`), lancez `npm run serve`,
  redémarrez PowerPoint › Insertion › Mes compléments.
- **PowerPoint sur le web** : `npm run serve`, puis Insertion › Compléments › Mes compléments ›
  *Charger mon complément* › `manifest.xml`.

## Déploiement pour une équipe

```bash
npm run build:static -- https://mon-serveur.example.com/slidetex
```

Copiez `dist/` sur n'importe quel hébergement HTTPS (IIS, GitHub Pages, SharePoint…) et distribuez
`dist/manifest.xml` (centre d'administration Microsoft 365 › Applications intégrées, ou catalogue partagé).
Le moteur MathJax marche partout ; le moteur LaTeX local nécessite `npm run serve` sur le poste de
l'utilisateur (le volet appelle `https://localhost:3000`, modifiable dans *Outils*). Ajoutez alors
l'origine du site à `SLIDETEX_ALLOWED_ORIGINS`.

## Scripts

| Commande | Rôle |
|---|---|
| `npm start` / `npm run stop` | serveur + sideload PowerPoint bureau |
| `npm run serve` | serveur HTTPS seul (`--http` pour HTTP, `--port N`) |
| `npm run certs` | (ré)installe le certificat localhost |
| `npm run manifest -- https://hôte` | régénère `manifest.xml` pour une autre URL |
| `npm run build:legacy` | régénère le bundle IE11 (`public/legacy/`) après modification du code |
| `npm run build:static -- https://hôte/chemin` | version statique dans `dist/` |
| `npm test` | tests de bout en bout (Chromium + simulateur Office.js ; nécessite `npm i -D playwright` et `npx playwright install chromium`) |

Variables : `SLIDETEX_PORT`, `SLIDETEX_ALLOWED_ORIGINS`, `SLIDETEX_LATEX`, `SLIDETEX_DVISVGM`.

## Fonctionnement

- Le volet (`public/`) rend la formule en SVG (MathJax, ou `latex` → DVI → `dvisvgm --no-fonts` via le
  serveur), puis la normalise pour PowerPoint : glyphes en `<path>` (pas de `<use>`), couleurs explicites,
  dimensions en points.
- Insertion via `setSelectedDataAsync` (SVG `XmlSvg`, sinon PNG). La source + paramètres sont écrits en
  JSON dans le tag `SLIDETEX` de la forme.
- Mise à jour : nouvelle image insérée au même endroit, mise à l'échelle et z-order restaurés, ancienne supprimée.
- Sécurité du serveur LaTeX : écoute uniquement sur la boucle locale, origines autorisées (CORS),
  `-no-shell-escape`, `openin_any/openout_any=p`, délai max 30 s, dossier temporaire par compilation.

## Dépannage

- **Le volet est blanc / « contenu bloqué »** : le serveur ne tourne pas ou le certificat n'est pas
  approuvé → `npm run certs`, puis `npm run serve`, ouvrez https://localhost:3000/taskpane.html dans
  Edge : la page doit s'afficher sans alerte.
- **« LaTeX local — indisponible »** : `latex` et `dvisvgm` doivent être dans le PATH (testez
  `latex --version` et `dvisvgm --version` dans un terminal), puis *Outils › Tester le serveur*.
- **Erreur TeX** : le message s'affiche sous l'aperçu (ex. `Missing close brace`).
- Vider le cache des add-ins Office (Windows) : supprimez `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
