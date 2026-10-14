# Relier le site au Google Sheet

## 1. Ajouter le script au Google Sheet
1. Ouvrez votre Google Sheet.
2. Menu **Extensions → Apps Script**.
3. Effacez le contenu de `Code.gs` et collez tout le contenu du fichier `Code.gs` de ce dossier.
4. Cliquez sur l'icône **Enregistrer** (disquette).

## 2. Créer l'onglet « Réservations »
1. En haut, choisissez la fonction **setup**, puis cliquez sur **Exécuter**.
2. Google demande une autorisation : **Examiner les autorisations** → votre compte → **Paramètres avancés** → **Accéder au projet (non sécurisé)** → **Autoriser**.
   (C'est normal : c'est votre propre script.)
3. Un onglet **Réservations** apparaît dans le Sheet, avec un récapitulatif (total des passagers confirmés).

## 3. Publier en application Web
1. Bouton **Déployer → Nouveau déploiement**.
2. Icône engrenage → **Application Web**.
3. Réglages :
   - **Exécuter en tant que** : Moi
   - **Qui a accès** : **Tout le monde**
4. **Déployer**, puis copiez l'**URL de l'application Web** (elle se termine par `/exec`).

## 4. Brancher le site
Envoyez cette URL à Claude, ou collez-la dans `index.html` :

```js
const CONFIG = {
  sheetUrl: "https://script.google.com/macros/s/XXXX/exec",
  ...
```

Pour vérifier : ouvrez l'URL `/exec` dans le navigateur, le message « enregistrement des réservations : OK » doit s'afficher.

> Si vous modifiez le script plus tard : **Déployer → Gérer les déploiements → crayon → Version : Nouvelle version → Déployer**. L'URL reste la même.
