# Déploiement de l'interface SLK Manager sur Render (gratuit)

L'interface (ce dossier) est une application React construite avec Vite. Une
fois déployée, elle appelle l'API backend déjà en ligne
(https://slk-manager-api.onrender.com) pour la connexion et les données.

## Résumé

- Type de service Render : **Static Site** (site statique) — gratuit, pas de
  mise en veille contrairement au backend.
- Build Command : `npm install && npm run build`
- Publish Directory : `dist`
- Variable d'environnement : `VITE_API_URL` = l'adresse de votre API.

## Étapes

### 1. Créer un dépôt GitHub

Comme pour le backend : créer un nouveau dépôt privé, par exemple
`slk-manager-frontend`, et y téléverser **tout le contenu de ce dossier**
(index.html, package.json, vite.config.js, .gitignore, et le dossier `src/`
avec ses fichiers).

⚠️ Ne pas téléverser `node_modules` s'il existe (le .gitignore l'exclut déjà).

### 2. Créer le service sur Render

1. Sur dashboard.render.com : **New +** → **Static Site**.
2. Connecter le dépôt `slk-manager-frontend`.
3. Renseigner :
   - **Name** : `slk-manager`
   - **Branch** : `main`
   - **Build Command** : `npm install && npm run build`
   - **Publish Directory** : `dist`
4. Section **Environment Variables**, ajouter :
   - **Key** : `VITE_API_URL`
   - **Value** : `https://slk-manager-api.onrender.com`
     (ou l'adresse exacte de votre backend si elle diffère)
5. Cliquer sur **Create Static Site**.

Render construit l'application (2-4 minutes) et fournit une adresse du type
`https://slk-manager.onrender.com`.

### 3. Tester

1. Ouvrir l'adresse fournie par Render.
2. L'écran de connexion s'affiche.
3. Se connecter avec l'email et le mot de passe du compte Direction créé
   précédemment (via la page creer-premier-compte.html).
4. Premier appel : le backend gratuit peut mettre ~50 s à se réveiller — c'est
   normal, le message d'attente s'affiche.

## Note sur la connexion frontend ↔ backend

Le backend autorise déjà les appels entre services (CORS ouvert). Pour la vraie
mise en production, on resserrera cette autorisation à la seule adresse du
frontend, et on retirera la variable BOOTSTRAP_SECRET du backend.

## Important : ce qui est branché, et ce qui ne l'est pas encore

Dans cette version, seule la **connexion** est réellement branchée sur l'API.
Une fois connecté, les autres écrans (devis, chantiers, comptabilité...)
affichent encore des données de démonstration en mémoire. Le branchement de
chaque module sur l'API se fera ensuite, un par un, en validant à chaque étape.
