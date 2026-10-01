# Classer dans OneDrive : guide d'installation

Le complément enregistre un mail Outlook (.eml et/ou PDF) et ses pièces jointes dans un dossier de ton OneDrive, puis marque le mail dans Outlook avec une catégorie au nom du dossier (par exemple `Classé - LASIS`) et affiche « Déjà classé dans ... » quand tu le rouvres.

## Étape 1 : héberger les fichiers (déjà fait)

Ton dépôt `xcavocat/classer-drive` est en ligne. Il suffit d'y remplacer les anciens fichiers par ceux de ce zip :
dans le dépôt, **Add file**, **Upload files**, glisse le contenu du dossier décompressé (fichiers + dossier `assets`), **Commit changes**. Les fichiers du même nom sont remplacés.

## Étape 2 : déclarer l'application chez Microsoft

Va sur https://entra.microsoft.com et connecte-toi avec ton compte Microsoft 365 (celui d'Outlook).

1. Menu de gauche : **Applications**, puis **Inscriptions d'applications**, puis **+ Nouvelle inscription**.
   - Nom : `Classer dans OneDrive`
   - Types de comptes pris en charge : **Comptes dans un annuaire organisationnel (tout locataire Microsoft Entra ID) et comptes Microsoft personnels**
   - URI de redirection : plateforme **Application monopage (SPA)**, valeur `https://xcavocat.github.io/classer-drive/auth.html`
   - **S'inscrire**
2. Sur la page qui s'affiche, copie l'**ID d'application (client)** (36 caractères).
3. Menu **Authentification** : sous « Application monopage », clique **Ajouter un URI** et ajoute `brk-multihub://xcavocat.github.io`, puis **Enregistrer**. Cette adresse permet à Outlook de te connecter sans fenêtre.
4. Menu **Autorisations de l'API** : **+ Ajouter une autorisation**, **Microsoft Graph**, **Autorisations déléguées**, coche **Files.ReadWrite.All**, **Ajouter des autorisations**. (`User.Read` est déjà présent.)

Si le portail te refuse la création d'une inscription, c'est que le compte Microsoft 365 du cabinet l'a réservée à l'administrateur : il faudra lui demander de faire ces quatre étapes (et éventuellement de « donner le consentement administrateur » sur la page des autorisations).

## Étape 3 : installer dans Outlook

Si tu avais installé une version précédente, retire-la d'abord (même écran, « Supprimer »).

- **Outlook web ou nouvel Outlook** : https://aka.ms/olksideload, **Mes compléments**, **Ajouter un complément personnalisé**, **Ajouter à partir d'un fichier**, choisis `manifest.xml`.
- **Outlook classique (Windows)** : Accueil, **Obtenir des compléments**, puis même chemin.

Ouvre un mail, clique **Classer dans OneDrive**. Au premier lancement, colle l'ID d'application, puis **Connecter OneDrive** et accepte les autorisations demandées.

## Utilisation

- Le volet propose le dossier probable (même conversation, même correspondant, même organisation, mots de l'objet) et le présélectionne quand il l'a déjà appris.
- Tu peux chercher un dossier, naviguer dans ton OneDrive et dans les dossiers « Partagés avec moi », ou créer un dossier.
- Nommage par défaut : `AAAA-MM-JJ - Expéditeur - Objet`, modifiable.
- Pour retrouver tous les mails d'un dossier dans Outlook : recherche `catégorie:"Classé - LASIS"`.

## Bon à savoir

- Le PDF est produit par OneDrive à partir du .eml (images comprises) ; à défaut, à partir du corps du mail.
- La recherche par nom couvre ton OneDrive ; les dossiers partagés par d'autres se trouvent par la navigation et sont ensuite mémorisés.
- .eml : Outlook récent requis (Microsoft 365 à jour, Outlook web, nouvel Outlook).
- Le projet Google Cloud créé précédemment ne sert plus : tu peux le supprimer ou l'ignorer.
