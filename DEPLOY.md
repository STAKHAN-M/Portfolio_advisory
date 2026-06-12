# Déploiement sur Render

Guide pas-à-pas pour publier le dashboard (app Dash) sur [Render](https://render.com) — tier gratuit.

## 0. Pré-requis de sécurité (déjà fait dans le code)

- ✅ Les clés API sont lues depuis des **variables d'environnement** (`FRED_API_KEY`, `FMP_API_KEY`), plus aucune clé en dur.
- ✅ `.gitignore` exclut `.env`, les fichiers Excel personnels (`*.xlsx`, `*.xlsm`) et les caches.
- ⚠️ **Vérifie avant le 1er push** que `.env`, `base_transactions_propre.xlsx` et `Portfolio_Le_M.xlsm` ne sont PAS suivis par git :
  ```bash
  git status            # ces fichiers ne doivent PAS apparaître
  git check-ignore .env base_transactions_propre.xlsx Portfolio_Le_M.xlsm
  ```

## 1. Pousser le code sur GitHub

```bash
git init
git add .
git commit -m "Portfolio dashboard"
git branch -M main
git remote add origin https://github.com/<ton-compte>/<ton-repo>.git
git push -u origin main
```

> Si tu avais déjà committé `.env` ou les fichiers Excel par le passé, ils restent dans l'historique git.
> Dans ce cas, **révoque/régénère ta clé FRED** et purge l'historique (`git filter-repo`) avant de publier.

## 2. Créer le service sur Render

Option A — **Blueprint (recommandé)** : le repo contient déjà `render.yaml`.
1. Sur Render → **New** → **Blueprint** → connecte ton repo GitHub.
2. Render lit `render.yaml` et configure tout automatiquement.

Option B — **Manuel** : New → **Web Service** → repo GitHub, puis :
- **Build Command** : `pip install -r requirements.txt`
- **Start Command** : `gunicorn app:server --workers 1 --threads 8 --timeout 180 --bind 0.0.0.0:$PORT`
- **Instance Type** : Free

## 3. Définir les variables d'environnement (Render → Environment)

| Clé | Valeur |
|-----|--------|
| `FRED_API_KEY` | ta clé FRED |
| `FMP_API_KEY`  | (optionnel) ta clé FMP |
| `PYTHON_VERSION` | `3.11.9` |

## 4. Déployer

Render build + lance le service. La 1ʳᵉ requête peut être lente (téléchargements yfinance/FRED).
URL publique : `https://<nom>.onrender.com`.

## Notes & limites du tier gratuit

- **Mise en veille** : le service s'endort après ~15 min d'inactivité ; le réveil prend 30–60 s.
- **RAM 512 Mo** : on tourne en `--workers 1` pour rester dans l'enveloppe.
- **yfinance** : les IP cloud partagées peuvent être *rate-limited* par Yahoo Finance. Si les cours ne se chargent pas, c'est généralement temporaire.
- **Caches** (`.macro_cache.json`, etc.) : stockés sur un disque éphémère → réinitialisés à chaque redéploiement (re-téléchargement automatique, sans impact fonctionnel).
- **Données** : aucune donnée personnelle n'est stockée ; le portefeuille est uploadé par l'utilisateur à chaque session.
