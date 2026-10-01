# Classer dans Drive (v2) : guide d'installation

Tout se fait dans le navigateur, sans ligne de commande. Compte environ 30 minutes, une seule fois. Garde trois onglets ouverts : GitHub, Google Cloud, Outlook.

## Étape 1 : mettre le complément en ligne (GitHub Pages)

Outlook a besoin que les fichiers du complément soient accessibles sur internet en https. GitHub Pages les héberge gratuitement. Tu n'as rien à programmer : tu déposes les fichiers, GitHub les publie.

1. Va sur https://github.com et clique **Sign up**. Choisis ton **pseudo** avec soin : il apparaîtra dans l'adresse (par exemple `xcelle-avocat`).
2. Une fois connecté, clique sur le **+** en haut à droite, puis **New repository**.
   - Repository name : `classer-drive` (exactement, en minuscules)
   - Coche **Public**
   - Coche **Add a README file**
   - Clique **Create repository**
3. Décompresse le zip `classer-drive.zip` sur ton ordinateur.
4. Dans la page du dépôt sur GitHub, clique **Add file**, puis **Upload files**. Ouvre le dossier décompressé et fais glisser **son contenu** (les fichiers et le dossier `assets`), pas le dossier lui-même. Clique **Commit changes** en bas.
5. Toujours dans le dépôt : **Settings** (onglet en haut), puis **Pages** (menu de gauche).
   - Source : **Deploy from a branch**
   - Branch : **main**, dossier **/ (root)**, puis **Save**
6. Attends 1 à 2 minutes et recharge la page : l'adresse apparaît, du type `https://xcavocat.github.io/classer-drive/`.
7. Vérifie : ouvre `https://xcavocat.github.io/classer-drive/taskpane.html`. Tu dois voir le titre « Classer dans Drive ». C'est normal que le reste soit vide en dehors d'Outlook.

## Étape 2 : autoriser l'accès à Google Drive

Sur https://console.cloud.google.com, avec ton compte Google (celui du Drive) :

1. En haut, sélecteur de projet, **Nouveau projet**, nom « Classer dans Drive », **Créer**, puis sélectionne-le.
2. Menu, **API et services**, **Bibliothèque** : cherche **Google Drive API**, clique **Activer**.
3. **API et services**, **Écran de consentement OAuth** (ou « Google Auth Platform ») :
   - Type : **Externe** (ou **Interne** si ton Drive est un Google Workspace du cabinet)
   - Nom de l'application, ton adresse, puis enregistre
   - Section **Audience** ou **Utilisateurs test** : ajoute ton adresse Google. Laisse l'application en mode **Test** : pas de validation Google nécessaire pour ton usage.
4. **Identifiants**, **Créer des identifiants**, **ID client OAuth**, type **Application Web** :
   - Origines JavaScript autorisées : `https://xcavocat.github.io`
   - URI de redirection autorisés : `https://xcavocat.github.io/classer-drive/auth.html`
   - **Créer**, puis copie l'**ID client** (il finit par `.apps.googleusercontent.com`). Tu le colleras dans Outlook à l'étape 4.

## Étape 3 : adapter le manifeste

Ouvre `manifest.xml` (celui du zip, sur ton ordinateur) avec le Bloc-notes. Menu **Édition**, **Remplacer** : remplace `xcavocat` par ton pseudo GitHub, **Remplacer tout**, enregistre.

## Étape 4 : installer dans Outlook

- **Outlook web ou nouvel Outlook** : ouvre https://aka.ms/olksideload, **Mes compléments**, **Ajouter un complément personnalisé**, **Ajouter à partir d'un fichier**, choisis `manifest.xml`. Outlook signale que le complément peut lire et modifier les mails : c'est nécessaire pour poser la catégorie de classement.
- **Outlook classique (Windows)** : Accueil, **Obtenir des compléments**, puis même chemin.

Ouvre un mail, clique **Classer dans Drive** dans le ruban. Au premier lancement, colle l'ID client Google, puis **Connecter Google Drive**.

## Ce que fait la v2

- **Bandeau « Déjà classé dans ... »** en haut du volet, avec le lien vers le dossier Drive et la date du classement. L'information est enregistrée sur le mail lui-même : tu la retrouves depuis n'importe quel poste.
- **Catégorie Outlook au nom du dossier** (par exemple `Drive - LASIS`), avec une couleur propre à chaque dossier. Elle se voit directement dans la liste des mails, sans les ouvrir. Pour retrouver tous les mails d'un dossier : tape `catégorie:"Drive - LASIS"` (ou `category:`) dans la recherche Outlook. Si le dossier choisi porte un nom générique (Correspondance, Pièces, Procédure...), la catégorie reprend le dossier parent : `Drive - LASIS / Correspondance`.
- Le marquage se désactive dans les options du volet si tu ne le veux pas sur un mail donné.

## Bon à savoir

- Session Google : une heure ; une fenêtre de reconnexion s'ouvre brièvement si besoin.
- .eml : Outlook récent requis (Microsoft 365 à jour, Outlook web, nouvel Outlook).
- PDF : les images intégrées au corps (logos, signatures) n'y figurent pas ; le .eml est intégral.
- Si le Microsoft 365 du cabinet bloque les compléments personnalisés, l'ajout du manifeste sera refusé : il faudra le faire déployer par l'administrateur.
- Pour mettre à jour le complément plus tard : dans GitHub, **Add file**, **Upload files**, dépose les nouveaux fichiers (ils remplacent les anciens).
