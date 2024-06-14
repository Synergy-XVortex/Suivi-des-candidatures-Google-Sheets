# Projet Google Apps Script : Ajouter Candidature
Ce projet contient des scripts Google Apps Script permettant de gérer des candidatures directement depuis une feuille de calcul Google Sheets. Les utilisateurs peuvent ajouter de nouvelles candidatures via une interface utilisateur personnalisée et dynamique.

# Scripts inclus :
1. Code.gs
    - Description : Gestion des candidatures avec Google Apps Script.
    - Fonctionnalités :
        - Ajout de candidatures dans une feuille de calcul.
        - Ajout de nouvelles entreprises à une liste déroulante.
        - Mise à jour automatique des listes déroulantes et mise en forme conditionnelle.

2. FormulaireCandidature.html
    - Description : Interface utilisateur HTML pour l'ajout de nouvelles candidatures.
    - Fonctionnalités :
        - Formulaire pour saisir les informations de la candidature.
        - Sélection dynamique d'entreprises existantes ou ajout de nouvelles entreprises.
        - Validation de la saisie utilisateur.

# Utilisation
Pour utiliser ces scripts avec votre propre feuille de calcul Google Sheets, suivez les étapes ci-dessous :

1. Configuration initiale :
    - Créez une nouvelle feuille de calcul Google Sheets ou utilisez une existante.
    - Copiez le contenu de chaque script (Code.gs et FormulaireCandidature.html) dans l'éditeur de script associé à votre feuille de calcul :
        - Ouvrez votre feuille de calcul.
        - Allez dans le menu "Extensions" -> "Apps Script".
        - Collez le script correspondant dans l'éditeur de script et sauvegardez-le.

2. Installation du menu personnalisé :
    - Lorsque vous ouvrez la feuille de calcul, un menu personnalisé "Candidatures" est ajouté automatiquement. Cliquez sur ce menu pour afficher l'option "Ajouter une candidature".

3. Ajout d'une nouvelle candidature :
    - Cliquez sur "Ajouter une candidature" dans le menu "Candidatures".
    - Remplissez le formulaire qui s'affiche avec les détails de la candidature (entreprise, titre de l'offre, lien).
    - Cliquez sur "Ajouter Candidature" pour soumettre le formulaire. Les informations seront automatiquement ajoutées à la feuille de calcul.

# Exemple de configuration
Un exemple de configuration type de feuille de calcul est fourni dans ce dépôt GitHub pour vous aider à démarrer rapidement avec les scripts et leur utilisation. Assurez-vous d'avoir les autorisations nécessaires pour accéder et modifier la feuille de calcul associée à ces scripts.
Lien vers la feuille type:
- https://docs.google.com/spreadsheets/d/11Sz7P_bk9WTBEnWX4phSnCvWJVZwe-19hquaYakiXXw/edit?usp=sharing

# Auteur
Ce projet a été développé par Clément Vongsanga. Pour toute question, suggestion ou contribution, n'hésitez pas à ouvrir une issue ou à proposer une pull request sur GitHub.
