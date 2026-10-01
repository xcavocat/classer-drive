// Configuration du complément "Classer dans Drive".
// GOOGLE_CLIENT_ID peut rester tel quel : le volet le demande au premier lancement.
window.CONFIG = {
  GOOGLE_CLIENT_ID: "__CLIENT_ID__",
  // Accès complet nécessaire pour voir et parcourir tes dossiers existants.
  SCOPE: "https://www.googleapis.com/auth/drive",
  // Facultatif : ID d'un dossier racine (ex. "Dossiers clients"). Vide = Mon Drive.
  // Tu peux aussi le définir depuis le volet ("Définir comme racine").
  ROOT_FOLDER_ID: "",
  ROOT_FOLDER_NAME: "Mon Drive"
};
