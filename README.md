
# Label Generator - 100% online

Deploie sur GitHub Pages.

## Fichiers
- index.html
- app.js
- styles.css
- config.js
- presets.json

## Configuration
Dans config.js, remplace :

GOOGLE_CLIENT_ID: "REMPLACE_PAR_TON_CLIENT_ID_OAUTH"

par ton vrai OAuth Client ID Google.

## Google Cloud
Activer :
- Google Drive API

Creer un OAuth Client ID de type Web application.

Ajouter comme origine autorisee :
https://TON-USER.github.io

Et pour test local :
http://localhost:8080

## Test local
python -m http.server 8080
