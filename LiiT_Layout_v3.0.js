var doc = app.activeDocument;
var w = doc.documentPreferences.pageWidth;
var h = doc.documentPreferences.pageHeight;

// Funzione per arrotondare al multiplo di 4 più vicino
function roundToNearestMultiple(number, multiple) {
  return Math.round(number / multiple) * multiple;
}

doc.textFramePreferences.firstBaselineOffset = FirstBaseline.CAP_HEIGHT;

// Crea una finestra di dialogo
var myDialog = app.dialogs.add({ name: "Inserire punti base testo" });

// Aggiungi un campo di testo alla finestra di dialogo
var myDialogText = myDialog.dialogColumns.add().staticTexts.add({
  staticLabel: "Inserire punti base testo:",
});

// Aggiungi un campo di input numerico alla finestra di dialogo
var myNumberField = myDialog.dialogColumns.add().textEditboxes.add({
  editContents: "0",
});

// Mostra la finestra di dialogo
if (myDialog.show() == true) {
  // Salva il numero inserito dall'utente in una variabile
  var userNumber = parseInt(myNumberField.editContents);

  // Esegui le azioni desiderate con la variabile salvata
  // Esempio: Stampa il numero nella console
  $.writeln("Il numero inserito è: " + userNumber);

  // Arrotonda il numero inserito all'intero più vicino
  userNumber = Math.round(userNumber);
}

// Chiudi la finestra di dialogo
myDialog.destroy();

function getUnitString(unit) {
  switch (unit) {
    case MeasurementUnits.MILLIMETERS:
      return "mm";
    case MeasurementUnits.INCHES:
      return "inch";
    case MeasurementUnits.POINTS:
      return "point";
    case MeasurementUnits.PIXELS:
      return "px";
    case MeasurementUnits.AGATES:
      return "agates";
    case MeasurementUnits.CENTIMETERS:
      return "cm";
    default:
      return "";
  }
}

// Funzione per la conversione di millimetri in un'altra unità di misura
function convertToUnit(value, unit) {
  // Puoi aggiungere altri casi a seconda delle tue esigenze
  switch (unit) {
    case "mm":
      return value;
    case "inch":
      return value * 0.0394;
    case "point":
      return value * 2.83465; // 1 mm ≈ 2.83465 punti
    case "cm":
      return value * 0.1; // 1 mm ≈ 2.83465 punti
    case "agates":  
      return value * 0.5512; // 1 mm ≈ 2.83465 punti
    case "px":
      return value * 2.835; // Conversione in pixel considerando la risoluzione
    // Aggiungi altri casi se necessario
    default:
      return value;
  }
}

// Otteniamo l'unità di misura e stimiamo la risoluzione del documento in pixel per millimetro
var unit = doc.viewPreferences.horizontalMeasurementUnits;

var l = Math.min(w, h);

var master = doc.masterSpreads[0];

for (var i = 0; i < master.pages.length; i++) {
  var page = master.pages.item(i);
  var marginSize = convertToUnit(l / 15, unit);

  if (l/42 > convertToUnit(2, unit)) {
    var columnNumber = 12;
    var gutterSize = convertToUnit(l / 42, unit);
  } else {
    var columnNumber = 6;
    var gutterSize = convertToUnit(l / 21, unit);
  }

  var marginPrefs = page.marginPreferences;
  marginPrefs.properties = {
    top: marginSize,
    left: marginSize,
    right: marginSize,
    bottom: marginSize,
    columnCount: columnNumber,
    columnGutter: gutterSize,
  };
}



var colors = [
  { name: 'Grigio', hex: ['1b1b1b', '333333', '3f3f3f', '525252', '686868', '878787', 'b0b0b0', 'e3e3e3'] },
  { name: 'Cartaria', hex: ['00305b', '004784', '005cac', '0073D4', '0a8dff', '50aeff', '93ccff', 'd7edff'] },
  { name: 'Pop', hex: ['ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff'] },
  { name: 'Acqua', hex: ['033f36', '04584b', '057160', '068c78', '07a990', '08c9ad', '0aeecb', 'e4fefa'] },
  { name: 'Pesca', hex: ['591806', '822503', 'a92e0c', 'd2390f', 'f0572c', 'f4886b', 'f8b6a3', 'fce3db'] },
  { name: 'Rosa', hex: ['580e37', '811550', 'a91b69', 'd22282', 'e354a3', 'eb84bc', 'f2b2d5', 'fae1ee'] },

  { name: 'Luxury', hex: ['ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff'] },
  { name: 'Salvia', hex: ['263328', '384b3a', '4a624c', '5c7b5f', '719575', '9fb6a2', 'b9ccbb', 'e5ece5'] },
  { name: 'PEsca', hex: ['591806', '822503', 'a92e0c', 'd2390f', 'f0572c', 'f4886b', 'f8b6a3', 'fce3db'] },
  { name: 'Sabbia', hex: ['3d2e1a', '594326', '745832', '906e3e', 'af864b', 'c4a374', 'd6bfa0', 'f2e9df'] },

  { name: 'GH', hex: ['ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff', 'ffffff'] },
  { name: 'Rosso', hex: ['511d1a', '782b26', '9d3932', 'c24940', 'c95b56', 'db938d', 'e8bab6', 'f6e4e2'] },
  { name: 'Blu', hex: ['213045', '304765', '3f5d85', '4f74a6', '6585b4', '90aacb', 'b7c8dd', 'e3eaf2'] },
  { name: 'Giallo', hex: ['3f320c', '5a4811', '755d16', '91741b', 'af8c21', 'deb84e', 'e5c973', 'f7efd5'] },
];

for (var i = 0; i < colors.length; i++) {
  var color = colors[i];
  var colorGroup = doc.colorGroups.add(color.name);

  for (var j = 0; j < color.hex.length; j++) {
    var hex = color.hex[j];
    var c = [
      parseInt(hex.substring(0, 2), 16),
      parseInt(hex.substring(2, 4), 16),
      parseInt(hex.substring(4, 6), 16),
    ];

    var index = j + 1;
    var name = color.name + ' ' + index.toString();

    var newColor = doc.colors.add({
      space: ColorSpace.RGB,
      name: name,
      colorValue: c,
    });

    colorGroup.colorGroupSwatches.add(newColor);
  }
}

var headingLevels = [
  { name: 'Step 8', pointSize: userNumber * 12, leading: userNumber * 12 * 1.1, tracking: -25, appliedFont: 'Inter', kerningMethod: "Optical" },
  { name: 'Step 7', pointSize: userNumber * 8, leading: userNumber * 8 * 1.1, tracking: -25, appliedFont: 'Inter', kerningMethod: "Optical" },
  { name: 'Step 6', pointSize: userNumber * 6, leading: userNumber * 6 * 1.1, tracking: -10, appliedFont: 'Inter', kerningMethod: "Optical" },
  { name: 'Step 5', pointSize: userNumber * 4, leading: userNumber * 4 * 1.1, tracking: -10, appliedFont: 'Inter', kerningMethod: "Optical" },
  { name: 'Step 4', pointSize: userNumber * 3, leading: userNumber * 3 * 1.1, tracking: -10, appliedFont: 'Inter', kerningMethod: "Optical" },
  { name: 'Step 3', pointSize: userNumber * 2, leading: userNumber * 2 * 1.1, tracking: userNumber * 0.1, appliedFont: 'Inter', kerningMethod: "Metrics" },
  { name: 'Step 2', pointSize: userNumber * 1.5, leading: userNumber * 1.5 * 1.1, tracking: userNumber * 0.1, appliedFont: 'Inter', kerningMethod: "Metrics" },
  { name: 'Step 1', pointSize: userNumber * 1.25, leading: userNumber * 1.25 * 1.25, tracking: userNumber * 0.1, appliedFont: 'Inter', kerningMethod: "Metrics" },
  { name: 'Step 0', pointSize: userNumber, leading: userNumber * 1.25, tracking: userNumber * 0.1, appliedFont: 'Inter', kerningMethod: "Metrics" },
  { name: 'Step -1', pointSize: userNumber * 0.75, leading: userNumber * 0.75 * 1.25, tracking: userNumber * 0.1, appliedFont: 'Inter', kerningMethod: "Metrics" },
  { name: 'Step -2', pointSize: userNumber * 0.66, leading: userNumber * 0.66 * 1.25, tracking: userNumber * 0.1, appliedFont: 'Inter', kerningMethod: "Metrics" },
];

// Funzione per eliminare gli stili di paragrafo esistenti
function deleteExistingParagraphStyles() {
  for (var i = doc.paragraphStyles.length - 1; i >= 0; i--) {
    doc.paragraphStyles.item(i).remove();
  }
}

// Arrotonda i valori pointSize nell'array headingLevels al multiplo di 4 più vicino
for (var i = 0; i < headingLevels.length; i++) {
  headingLevels[i].pointSize = roundToNearestMultiple(headingLevels[i].pointSize, 2);
}

// Crea gli stili di paragrafo
for (var i = 0; i < headingLevels.length; i++) {
  var level = headingLevels[i];
  var paragraphStyle = doc.paragraphStyles.add({
    name: level.name,
    fontStyle: 'Regular',
    appliedFont: level.appliedFont
  });

  paragraphStyle.pointSize = level.pointSize;
  paragraphStyle.leading = level.leading;
  paragraphStyle.tracking = level.tracking;

  // Set kerning method
  if (level.kerningMethod !== undefined) {
    paragraphStyle.kerningMethod = level.kerningMethod;
  }

  paragraphStyle.hyphenation = false;
}

function deleteExistingCharactertyles() {
  for (var i = doc.characterStyles.length - 1; i >= 0; i--) {
    doc.characterStyles.item(i).remove();
  }
}

function createCharacterStyle(styleName, appliedFont, appliedStyle) {
  var characterStyle = doc.characterStyles.add({
    name: styleName,
    fontStyle: appliedStyle,
    appliedFont: appliedFont,
  });


  return characterStyle;
}


var characterStyles = [
  { name: 'Pop - Intestazione', appliedFont: 'Inter', fontStyle: 'Black'},
  { name: 'Pop - Sottotitolo', appliedFont: 'Inter', fontStyle: 'Bold'},
  { name: 'Pop - Highlight', appliedFont: 'Inter', fontStyle: 'Bold'},
  { name: 'LiiT/GH - Intestazione', appliedFont: 'Inter', fontStyle: 'Bold'},
  { name: 'LiiT/GH - Sottotitolo', appliedFont: 'Inter', fontStyle: 'Regular'},
  { name: 'LiiT/GH - Highlight', appliedFont: 'Inter', fontStyle: 'Bold'},
  { name: 'Luxury - Intestazione', appliedFont: 'Inter', fontStyle: 'Regular'},
  { name: 'Luxury - Sottotitolo', appliedFont: 'P22 Mackinac Pro', fontStyle: 'Book'},
  { name: 'Luxury - Highlight', appliedFont: 'P22 Mackinac Pro', fontStyle: 'Book Italic'},
];


for (var i = 0; i < characterStyles.length; i++) {
  var style = characterStyles[i];
  var charStyle = createCharacterStyle(style.name, style.appliedFont, style.fontStyle);  
}