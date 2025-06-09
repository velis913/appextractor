# Structure générale 
- C'est une application web qui permet d'extraire des données d'un fichier Excel et de les convertir en format JSON
- L'interface utilise un design moderne avec des champs de saisie et des boutons
- Les couleurs et le style sont définis avec des variables CSS pour une cohérence visuelle

1. **Fonctionnalités principales** :
- Upload d'un fichier Excel (.xlsx ou .xls)
- Configuration des paramètres de lecture :
  - Nom de la feuille Excel
  - Position des titres de cours
  - Position de début des données étudiants
  - Colonnes pour différentes informations (numéro, nom, moyenne, etc.)
- Conversion des données en JSON
- Téléchargement du résultat en fichier JSON

2. **Traitement des données** :
- Le script lit le fichier Excel de manière asynchrone
- Il extrait les informations selon les paramètres spécifiés
- Il ignore certaines colonnes selon les mots-clés définis
- Pour chaque étudiant, il collecte :
  - Les informations générales (numéro, nom)
  - Les notes pour chaque cours
  - La moyenne générale
  - Les crédits
  - La décision

3. **Points d'attention** :
- La bibliothèque XLSX doit être chargée (via js.js)
- Le script inclut des logs de débogage dans la console
- Les valeurs vides sont converties en null
- Le format de sortie est structuré et indenté

4. **Sécurité et validation** :
- Vérification de la présence du fichier
- Validation de l'existence de la feuille Excel
- Gestion des erreurs basique avec des alertes

C'est essentiellement un outil de conversion qui permet de transformer des données d'un bulletin de notes Excel en format JSON structuré, avec une interface utilisateur conviviale permettant de configurer tous les paramètres nécessaires.
