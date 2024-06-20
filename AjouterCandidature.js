function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Candidatures')
    .addItem('Ajouter une candidature', 'afficherDialogue')
    .addToUi();
}

function afficherDialogue() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('FormulaireCandidature')
      .setWidth(550)
      .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Ajouter une candidature');
}

function ajouterCandidatureAvecValeurs(entreprise, titre, lien) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  try {
    // Trouver la première cellule vide dans la colonne E
    var range = sheet.getRange('E:E');
    var values = range.getValues();
    var nextRow = values.findIndex(row => !row[0]) + 1; // +1 car les index sont basés sur 0
    
    // Ajouter l'entreprise dans la colonne B
    sheet.getRange(nextRow, 2).setValue(entreprise); // Colonne B pour l'entreprise
    
    // Copie de la mise en forme
    copyFormatting();
    
    // Ajouter le titre, le lien et la date dans les colonnes masquées E, F et G
    sheet.getRange(nextRow, 5).setValue(titre); // Colonne E pour le titre
    sheet.getRange(nextRow, 6).setValue(lien);  // Colonne F pour le lien
    sheet.getRange(nextRow, 7).setValue(new Date()); // Colonne G pour la date
    
    // Mettre à jour la colonne D avec "En Attente"
    sheet.getRange(nextRow, 4).setValue("En Attente");
  } catch (error) {
    Logger.log("Erreur lors de l'ajout de la candidature : " + error);
    // Vous pouvez ajouter ici une notification ou un traitement spécifique en cas d'erreur
    throw new Error("Erreur lors de l'ajout de la candidature : " + error);
  }
}

// Fonction pour ajouter une nouvelle entreprise
function ajouterNouvelleEntreprise(nouvelleEntreprise) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var entreprisesRange = sheet.getRange('A2:A' + sheet.getLastRow());
  var entreprises = entreprisesRange.getValues().flat().filter(e => e.trim() !== ''); // Récupère toutes les valeurs dans un tableau à une dimension
  
  // Vérifie si l'entreprise existe déjà dans la colonne A
  if (!entreprises.includes(nouvelleEntreprise)) {
    // Trouve la première ligne vide dans la colonne A
    var colonneARange = sheet.getRange('A2:A' + sheet.getLastRow());
    var colonneAValues = colonneARange.getValues().flat();
    var lastEmptyRowInA = colonneAValues.findIndex(value => value === '') + 2; // +2 pour l'index de ligne basé sur 1
    
    // Si aucune ligne vide n'est trouvée dans A, utilise la ligne suivante après la dernière ligne non vide
    if (lastEmptyRowInA === 1) {
      lastEmptyRowInA = sheet.getLastRow() + 1;
    }
    
    // Ajoute la nouvelle entreprise à la feuille de calcul dans la colonne A
    sheet.getRange(lastEmptyRowInA, 1).setValue(nouvelleEntreprise);

    // Ajoute la nouvelle entreprise à la liste déroulante
    var cell = SpreadsheetApp.getActive().getRange('A1:B1000');
    var range = SpreadsheetApp.getActive().getRange('A1:B1000');
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
    cell.setDataValidation(rule);
    
    // Forcer la mise à jour immédiate de la feuille de calcul
    SpreadsheetApp.flush();
    
    // Appliquer la mise en forme après l'ajout de la nouvelle entreprise
    applyColors();
    copyFormatting();
    
    // Retourne toutes les entreprises pour mettre à jour le menu déroulant dans Google Sheets
    return getEntreprises();
  } else {
    // Si l'entreprise existe déjà, retourne simplement les entreprises actuelles
    return entreprises;
  }
}

// Fonction pour récupérer la liste des entreprises
function getEntreprises() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var entreprisesRange = sheet.getRange('A2:A' + sheet.getLastRow());
  var entreprises = entreprisesRange.getValues().flat().filter(e => e.trim() !== ''); // Récupère toutes les valeurs dans un tableau à une dimension
  
  return entreprises;
}

// Fonction pour appliquer des couleurs aux entreprises dans la colonne A
function applyColors() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("A2:A"); // Colonne A à partir de la ligne 2
  var values = range.getValues();
  
  // Liste de 15 couleurs de fond
  var backgroundColors = [
    "#ffcfc9", "#ffc8aa", "#ffe5a0", "#d4edbc", "#bfe1f6",
    "#c6dbe1", "#e6cff2", "#3d3d3d", "#b10202", "#753800",
    "#473822", "#11734b", "#0a53a8", "#215a6c", "#5a3286"
  ];

  // Liste de 15 couleurs de texte
  var textColors = [
    "#b10202", "#753800", "#473821", "#11734b", "#0a53a8", "#215a6c",
    "#5a3286", "#e5e5e5", "#ffcfc9", "#ffc8aa", "#ffe5a0", "#d4edbc",
    "#bfe0f6", "#c6dbe1", "#e5cff2"
  ];

  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) { // Vérifiez si la cellule n'est pas vide
      var colorIndex = i % backgroundColors.length; // Cycle à travers les couleurs
      sheet.getRange(i + 2, 1).setBackground(backgroundColors[colorIndex]); // +2 pour commencer à la ligne 2
      sheet.getRange(i + 2, 1).setFontColor(textColors[colorIndex]); // Utiliser une couleur différente pour la police
    }
  }
}

// Fonction pour copier la mise en forme de la colonne A vers la colonne B
function copyFormatting() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeA = sheet.getRange("A2:A");
  var rangeB = sheet.getRange("B2:B");
  var valuesA = rangeA.getValues();
  var valuesB = rangeB.getValues();
  
  var formattingMap = {};

  // Créez une carte de la mise en forme de la colonne A
  for (var i = 0; i < valuesA.length; i++) {
    if (valuesA[i][0]) { // Vérifiez si la cellule n'est pas vide
      var cellA = sheet.getRange(i + 2, 1); // +2 pour commencer à la ligne 2
      formattingMap[valuesA[i][0]] = {
        background: cellA.getBackground(),
        fontColor: cellA.getFontColor(),
        fontWeight: cellA.getFontWeight(),
        fontStyle: cellA.getFontStyle(),
        fontSize: cellA.getFontSize()
      };
    }
  }

  // Appliquez la mise en forme aux cellules correspondantes de la colonne B
  for (var j = 0; j < valuesB.length; j++) {
    var companyName = valuesB[j][0];
    if (companyName && formattingMap[companyName]) {
      var cellB = sheet.getRange(j + 2, 2); // +2 pour commencer à la ligne 2
      cellB.setBackground(formattingMap[companyName].background);
      cellB.setFontColor(formattingMap[companyName].fontColor);
      cellB.setFontWeight(formattingMap[companyName].fontWeight);
      cellB.setFontStyle(formattingMap[companyName].fontStyle);
      cellB.setFontSize(formattingMap[companyName].fontSize);
    }
  }
}

// Fonction de déclenchement onEdit pour actualiser la mise en forme de la colonne B
function onEdit(e) {
  var range = e.range;
  var sheet = e.source.getActiveSheet();
  
  // Vérifiez si l'édition a eu lieu dans la colonne B
  if (range.getColumn() == 2) {
    copyFormatting();
  }
}
