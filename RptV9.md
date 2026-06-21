// ============================================================
//  Rapporto Congregazione Bologna FR — Google Apps Script v9
//  Sécurité maximale :
//  - Mot de passe stocké dans Script Properties (jamais dans le code)
//  - Vérification via hash SHA-256 (le code ne contient aucun secret)
//  - Clé AES dérivée du mot de passe côté client (PBKDF2)
// ============================================================
//
//  SETUP INITIAL (une seule fois) :
//  1. Dans l'éditeur Apps Script : icône ⚙ (engrenage) → "Propriétés du script"
//  2. Ajoute une propriété :
//       Nom  : EDITOR_PWD
//       Valeur : ton mot de passe (ex: @BOFra2046#)
//  3. Clique "Enregistrer"
//  4. Le mot de passe n'apparaît PLUS JAMAIS dans le code.
//
//  Pour changer le mot de passe :
//  1. Change la valeur dans Script Properties
//  2. Re-génère le hash en appelant doGet avec action=getPasswordHash
//     (temporaire, à désactiver ensuite)
//  3. Depuis l'app HTML : re-migre les noms avec le nouveau mot de passe
// ============================================================