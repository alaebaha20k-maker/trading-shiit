/**
 * ═══════════════════════════════════════════════════════════════════
 *  JOURNAL DE TRADING — MODÈLE COMPLET EN FRANÇAIS
 *  Fichier : trading_template_francais.gs
 *
 *  INSTALLATION :
 *    1. Extensions > Apps Script dans votre Google Sheet
 *    2. Collez ce fichier complet
 *    3. Sauvegardez puis lancez  creerModeleTrading()
 *    4. Accordez les autorisations demandées
 *
 *  FEUILLES CRÉÉES :
 *    • Statistiques    — tableau de bord (métriques, graphiques)
 *    • Tous les trades — saisie des trades
 *    • Vue Calendrier  — calendrier mensuel P/L
 * ═══════════════════════════════════════════════════════════════════
 */

/* ─────────────────────────────────────────────────────────────────
   CONFIGURATION GLOBALE
   ───────────────────────────────────────────────────────────────── */
var CFG = {
  // Noms des feuilles
  STATS      : 'Statistiques',
  TRADES     : 'Tous les trades',
  CALENDRIER : 'Vue Calendrier',

  // Structure feuille trades
  LIGNE_DEBUT   : 3,   // première ligne de données
  COL_DATE      : 2,   // colonne B  — date
  COL_PL        : 12,  // colonne L  — P/L
  MAX_LIGNES    : 500, // nombre de lignes de trades gérées

  // Structure feuille calendrier
  CAL_MOIS      : 'B1',
  CAL_ANNEE     : 'B2',
  CAL_GRILLE    : 5,   // ligne où commence la grille

  // Couleurs
  NOIR          : '#000000',
  BLANC         : '#ffffff',
  BLEU_FONCE    : '#1f4e79',
  BLEU_CLAIR    : '#d9e1f2',
  GRIS          : '#f2f2f2',
  VERT_BG       : '#e6ffe6',
  ROUGE_BG      : '#ffe6e6',
  JAUNE_BG      : '#fff2cc',
  ALT_BG        : '#f8f9fa',
  VERT_TXT      : '#1a7a1a',
  ROUGE_TXT     : '#cc0000',
};

var MOIS  = ['Janvier','Février','Mars','Avril','Mai','Juin',
             'Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
var JOURS = ['Lun','Mar','Mer','Jeu','Ven','Sam','Dim'];


/* ─────────────────────────────────────────────────────────────────
   MENU PERSONNALISÉ
   ───────────────────────────────────────────────────────────────── */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Trading FR')
    .addItem('🏗️  Initialiser le modèle',   'creerModeleTrading')
    .addSeparator()
    .addItem('📅  Générer le calendrier',   'genererCalendrier')
    .addItem('🗑️  Effacer le calendrier',   'effacerCalendrier')
    .addToUi();
}


/* ═══════════════════════════════════════════════════════════════════
   POINT D'ENTRÉE — creerModeleTrading()
   ═══════════════════════════════════════════════════════════════════ */
function creerModeleTrading() {
  var ui  = SpreadsheetApp.getUi();
  var rep = ui.alert(
    '🏗️ Initialisation du modèle de trading',
    'Cette action va créer / réinitialiser les 3 feuilles :\n' +
    '  • Statistiques\n  • Tous les trades\n  • Vue Calendrier\n\nContinuer ?',
    ui.ButtonSet.YES_NO
  );
  if (rep !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Création en cours, veuillez patienter…', '⏳', 30);

  var shSt  = _getOrCreate(ss, CFG.STATS);
  var shTr  = _getOrCreate(ss, CFG.TRADES);
  var shCal = _getOrCreate(ss, CFG.CALENDRIER);

  _setupStats(ss, shSt, shTr);
  _setupTrades(shTr);
  _setupCalendrier(shCal);

  // Supprimer la feuille vide par défaut si présente
  ['Sheet1', 'Feuille 1', 'Feuille1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && ss.getNumSheets() > 3) { try { ss.deleteSheet(s); } catch (e) {} }
  });

  shSt.activate();
  ui.alert(
    '✅ Modèle prêt !\n\n' +
    '1️⃣  Renseignez votre capital initial en B2 (feuille Statistiques)\n' +
    '2️⃣  Ajoutez vos trades dans "Tous les trades" à partir de la ligne 3\n' +
    '3️⃣  Générez le calendrier via  📊 Trading FR → Générer le calendrier'
  );
}


/* ═══════════════════════════════════════════════════════════════════
   FEUILLE : STATISTIQUES
   ═══════════════════════════════════════════════════════════════════ */
function _setupStats(ss, sh, shTr) {
  sh.clear();
  sh.clearFormats();
  sh.clearNotes();
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });

  /* ── Dimensions ── */
  var colW = {1:170, 2:130, 3:20, 4:100, 5:100, 6:100, 7:100, 8:20,
              9:115, 10:80, 11:80, 12:80, 13:120};
  Object.keys(colW).forEach(function (c) { sh.setColumnWidth(Number(c), colW[c]); });
  sh.setRowHeight(1, 42);
  for (var r = 2; r <= 30; r++) sh.setRowHeight(r, 30);
  sh.setRowHeight(17, 36);

  var annee = new Date().getFullYear();
  var T     = "'" + CFG.TRADES + "'";   // référence feuille (avec espaces)

  /* ── A1:B1  TITRE PRINCIPAL ── */
  sh.getRange('A1:B1').merge()
    .setValue(annee + ' BILAN GLOBAL')
    .setBackground(CFG.NOIR).setFontColor(CFG.BLANC)
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  /* ── A2:B12  MÉTRIQUES ── */
  var labels = [
    'Capital Initial',    // A2  → B2  saisie manuelle
    'Trades',             // A3  → B3
    'Gains',              // A4  → B4
    'Pertes',             // A5  → B5
    'Taux de Réussite',   // A6  → B6
    'Espérance (EV)',      // A7  → B7
    'Facteur de Profit',  // A8  → B8
    'Perte Moyenne',      // A9  → B9
    'Gain Moyen',         // A10 → B10
    'Profit Total',       // A11 → B11
    'ROI',                // A12 → B12
  ];
  sh.getRange(2, 1, labels.length, 1)
    .setValues(labels.map(function (l) { return [l]; }))
    .setFontWeight('bold').setBackground(CFG.BLEU_CLAIR)
    .setVerticalAlignment('middle');

  // Formules (B3:B12)
  var formules = [
    ['=COUNTA('+T+'!B3:B5000)'],
    ['=COUNTIF('+T+'!L3:L5000,">0")'],
    ['=COUNTIF('+T+'!L3:L5000,"<0")'],
    ['=IFERROR(B4/B3,0)'],
    ['=IFERROR(B10*B6+B9*(1-B6),0)'],
    ['=IFERROR(SUMIF('+T+'!L3:L5000,">0")/ABS(SUMIF('+T+'!L3:L5000,"<0")),0)'],
    ['=IFERROR(AVERAGEIF('+T+'!L3:L5000,"<0"),0)'],
    ['=IFERROR(AVERAGEIF('+T+'!L3:L5000,">0"),0)'],
    ['=SUM('+T+'!L3:L5000)'],
    ['=IFERROR(B11/B2,0)'],
  ];
  sh.getRange(3, 2, formules.length, 1).setValues(formules);

  // Formats numériques
  sh.getRange('B2').setNumberFormat('"€ "#,##0.00');
  sh.getRange('B3:B5').setNumberFormat('#,##0');
  sh.getRange('B6').setNumberFormat('0.00%');
  sh.getRange('B7').setNumberFormat('"€ "#,##0.00');
  sh.getRange('B8').setNumberFormat('0.00"x"');
  sh.getRange('B9:B11').setNumberFormat('"€ "#,##0.00');
  sh.getRange('B12').setNumberFormat('0.00%');

  // Mise en forme conditionnelle Profit + ROI
  var cfRules = [];
  ['B11', 'B12'].forEach(function (addr) {
    var rng = sh.getRange(addr);
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setFontColor(CFG.VERT_TXT).setRanges([rng]).build());
    cfRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor(CFG.ROUGE_TXT).setRanges([rng]).build());
  });
  sh.setConditionalFormatRules(cfRules);

  /* ── ROW 17  EN-TÊTES DE SECTION ── */
  sh.getRange('A17:G17').merge()
    .setValue('Détail par Jour de Semaine')
    .setBackground(CFG.NOIR).setFontColor(CFG.BLANC)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.getRange('I17:M17').merge()
    .setValue('Informations Mensuelles')
    .setBackground(CFG.NOIR).setFontColor(CFG.BLANC)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  /* ── ROW 18  EN-TÊTES DES TABLEAUX ── */
  sh.getRange(18, 1, 1, 7)
    .setValues([['Jour','Trades','Gains','Pertes','% Réussite','P/L Total','P/L Moyen']])
    .setFontWeight('bold').setBackground(CFG.BLEU_CLAIR).setHorizontalAlignment('center');

  sh.getRange(18, 9, 1, 5)
    .setValues([['Mois','Trades','Gains','Pertes','P/L (€)']])
    .setFontWeight('bold').setBackground(CFG.BLEU_CLAIR).setHorizontalAlignment('center');

  /* ── ROWS 19-25  DÉTAIL PAR JOUR (WEEKDAY type 2 : Lun=1 … Dim=7) ── */
  var joursFull = ['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche'];
  joursFull.forEach(function (jour, i) {
    var r  = 19 + i;
    var wd = i + 1; // type 2 : Lun=1
    sh.getRange(r, 1).setValue(jour).setFontWeight('bold');
    sh.getRange(r, 2).setFormula(
      '=SUMPRODUCT(('+T+'!B3:B5000<>"")*' +
      '(WEEKDAY('+T+'!B3:B5000,2)='+wd+'))');
    sh.getRange(r, 3).setFormula(
      '=SUMPRODUCT(('+T+'!L3:L5000>0)*' +
      '(WEEKDAY('+T+'!B3:B5000,2)='+wd+')*' +
      '('+T+'!B3:B5000<>""))');
    sh.getRange(r, 4).setFormula(
      '=SUMPRODUCT(('+T+'!L3:L5000<0)*' +
      '(WEEKDAY('+T+'!B3:B5000,2)='+wd+')*' +
      '('+T+'!B3:B5000<>""))');
    sh.getRange(r, 5).setFormula('=IFERROR(C'+r+'/B'+r+',0)')
      .setNumberFormat('0.00%');
    sh.getRange(r, 6).setFormula(
      '=SUMPRODUCT(('+T+'!B3:B5000<>"")*' +
      '(WEEKDAY('+T+'!B3:B5000,2)='+wd+')*' +
      T+'!L3:L5000)').setNumberFormat('"€ "#,##0.00');
    sh.getRange(r, 7).setFormula('=IFERROR(F'+r+'/B'+r+',0)')
      .setNumberFormat('"€ "#,##0.00');
    if (i % 2 === 0) sh.getRange(r, 1, 1, 7).setBackground(CFG.ALT_BG);
  });
  sh.getRange(19, 1, 7, 7).setBorder(true, true, true, true, true, true);

  /* ── ROWS 19-30  INFORMATIONS MENSUELLES (cols I-M) ── */
  MOIS.forEach(function (moisNom, m) {
    var mr  = 19 + m;
    var mn  = m + 1;
    sh.getRange(mr, 9).setValue(moisNom).setFontWeight('bold');
    sh.getRange(mr, 10).setFormula(
      '=SUMPRODUCT(('+T+'!B3:B5000<>"")*' +
      '(MONTH('+T+'!B3:B5000)='+mn+')*' +
      '(YEAR('+T+'!B3:B5000)='+annee+'))');
    sh.getRange(mr, 11).setFormula(
      '=SUMPRODUCT(('+T+'!L3:L5000>0)*' +
      '(MONTH('+T+'!B3:B5000)='+mn+')*' +
      '(YEAR('+T+'!B3:B5000)='+annee+')*' +
      '('+T+'!B3:B5000<>""))');
    sh.getRange(mr, 12).setFormula(
      '=SUMPRODUCT(('+T+'!L3:L5000<0)*' +
      '(MONTH('+T+'!B3:B5000)='+mn+')*' +
      '(YEAR('+T+'!B3:B5000)='+annee+')*' +
      '('+T+'!B3:B5000<>""))');
    sh.getRange(mr, 13).setFormula(
      '=SUMPRODUCT(('+T+'!B3:B5000<>"")*' +
      '(MONTH('+T+'!B3:B5000)='+mn+')*' +
      '(YEAR('+T+'!B3:B5000)='+annee+')*' +
      T+'!L3:L5000)').setNumberFormat('"€ "#,##0.00');
    if (m % 2 === 0) sh.getRange(mr, 9, 1, 5).setBackground(CFG.ALT_BG);
  });
  sh.getRange(19, 9, 12, 5).setBorder(true, true, true, true, true, true);

  /* ── GRAPHIQUES ── */
  _graphiqueEquite(sh, shTr);
  _graphiqueJoursSemaine(sh);
}

/* ── Graphique : Courbe d'Équité (cols D-G, rows 1-16) ── */
function _graphiqueEquite(shStats, shTrades) {
  var chart = shStats.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(shTrades.getRange(3, CFG.COL_DATE, CFG.MAX_LIGNES, 1))  // dates (axe X)
    .addRange(shTrades.getRange(3, 14, CFG.MAX_LIGNES, 1))             // équité cumulée col N
    .setPosition(1, 4, 5, 5)   // ligne 1, col D
    .setNumHeaders(0)
    .setOption('title', 'Courbe d\'Équité')
    .setOption('titleTextStyle', {fontSize: 13, bold: true})
    .setOption('hAxis', {title: 'Trades', textStyle: {fontSize: 10}})
    .setOption('vAxis', {title: 'P/L Cumulé (€)', textStyle: {fontSize: 10},
                         format: '€#,##0.00'})
    .setOption('legend', {position: 'none'})
    .setOption('colors', [CFG.BLEU_FONCE])
    .setOption('lineWidth', 2)
    .setOption('pointSize', 3)
    .setOption('backgroundColor', '#fafafa')
    .setOption('chartArea', {left: 60, top: 40, width: '80%', height: '75%'})
    .setOption('width', 460)
    .setOption('height', 300)
    .build();
  shStats.insertChart(chart);
}

/* ── Graphique : Performance par Jour de Semaine (cols I-M, rows 1-16) ── */
function _graphiqueJoursSemaine(sh) {
  var chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(18, 1, 8, 1))   // A18:A25 — labels (en-tête + 7 jours)
    .addRange(sh.getRange(18, 6, 8, 1))   // F18:F25 — P/L Total
    .setPosition(1, 9, 5, 5)             // ligne 1, col I
    .setNumHeaders(1)
    .setOption('title', 'Performance par Jour de Semaine')
    .setOption('titleTextStyle', {fontSize: 13, bold: true})
    .setOption('hAxis', {textStyle: {fontSize: 10}})
    .setOption('vAxis', {title: 'P/L (€)', textStyle: {fontSize: 10},
                         format: '€#,##0.00'})
    .setOption('legend', {position: 'none'})
    .setOption('colors', [CFG.BLEU_FONCE])
    .setOption('backgroundColor', '#fafafa')
    .setOption('chartArea', {left: 65, top: 40, width: '80%', height: '75%'})
    .setOption('width', 460)
    .setOption('height', 300)
    .build();
  sh.insertChart(chart);
}


/* ═══════════════════════════════════════════════════════════════════
   FEUILLE : TOUS LES TRADES
   ═══════════════════════════════════════════════════════════════════ */
function _setupTrades(sh) {
  sh.clear();
  sh.clearFormats();

  // Largeurs colonnes
  var cw = {1:45, 2:105, 3:95, 4:80, 5:70, 6:100, 7:100,
            8:95, 9:95, 10:80, 11:90, 12:110, 13:160, 14:120};
  Object.keys(cw).forEach(function (c) { sh.setColumnWidth(Number(c), cw[c]); });

  sh.setRowHeight(1, 38);
  sh.setRowHeight(2, 32);

  /* ── Ligne 1 : Titre ── */
  sh.getRange(1, 1, 1, 14).merge()
    .setValue('TOUS LES TRADES — JOURNAL DE TRADING')
    .setBackground(CFG.NOIR).setFontColor(CFG.BLANC)
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  /* ── Ligne 2 : En-têtes ── */
  var headers = [
    'N°', 'Date', 'Actif', 'Sens', 'Quantité',
    'Prix Entrée', 'Prix Sortie', 'Stop Loss', 'Take Profit',
    'Durée', 'Résultat Pts', 'P/L (€)', 'Commentaires', 'Équité Cumulée'
  ];
  sh.getRange(2, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold').setBackground(CFG.BLEU_CLAIR)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  var N = CFG.MAX_LIGNES;

  /* ── Formats ── */
  sh.getRange(3, 2,  N, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(3, 6,  N, 4).setNumberFormat('"€ "#,##0.00000');   // Prix E/S, SL, TP
  sh.getRange(3, 12, N, 1).setNumberFormat('"€ "#,##0.00');       // P/L
  sh.getRange(3, 14, N, 1).setNumberFormat('"€ "#,##0.00');       // Équité cumulée

  /* ── Numérotation auto colonne A ── */
  var numForms = [];
  for (var i = 0; i < N; i++) {
    numForms.push(['=IF(B' + (3 + i) + '="","",ROW()-2)']);
  }
  sh.getRange(3, 1, N, 1).setFormulas(numForms);

  /* ── Équité cumulée colonne N (helper pour le graphique) ── */
  var equiteForms = [['=IF(L3="","",L3)']];
  for (var j = 1; j < N; j++) {
    var r = 3 + j;
    equiteForms.push(['=IF(L' + r + '="","",N' + (r - 1) + '+L' + r + ')']);
  }
  sh.getRange(3, 14, N, 1).setFormulas(equiteForms);
  sh.getRange(2, 14).setNote('Colonne auxiliaire — courbe d\'équité. Ne pas modifier.');

  /* ── Validation Sens (ACHAT / VENTE) ── */
  var sensDV = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ACHAT', 'VENTE'], true)
    .setAllowInvalid(false).build();
  sh.getRange(3, 4, N, 1).setDataValidation(sensDV);

  /* ── Mise en forme conditionnelle P/L ── */
  var plRange = sh.getRange(3, 12, N, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(CFG.VERT_BG).setFontColor(CFG.VERT_TXT)
      .setRanges([plRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(CFG.ROUGE_BG).setFontColor(CFG.ROUGE_TXT)
      .setRanges([plRange]).build(),
  ]);

  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
}


/* ═══════════════════════════════════════════════════════════════════
   FEUILLE : VUE CALENDRIER
   ═══════════════════════════════════════════════════════════════════ */
function _setupCalendrier(sh) {
  sh.clear();
  sh.clearFormats();

  sh.setRowHeight(1, 38);
  sh.setRowHeight(2, 28);
  sh.setRowHeight(3, 28);
  sh.setRowHeight(4, 28);

  /* Titre */
  sh.getRange('A1:M1').merge()
    .setValue('VUE CALENDRIER MENSUEL')
    .setBackground(CFG.NOIR).setFontColor(CFG.BLANC)
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  /* Saisie mois / année (B1 et B2 requis par genererCalendrier) */
  sh.getRange('A2').setValue('Mois (1-12) :')
    .setFontWeight('bold').setBackground(CFG.BLEU_CLAIR).setHorizontalAlignment('right');
  sh.getRange('B2').setValue(new Date().getMonth() + 1);

  sh.getRange('A3').setValue('Année :')
    .setFontWeight('bold').setBackground(CFG.BLEU_CLAIR).setHorizontalAlignment('right');
  sh.getRange('B3').setValue(new Date().getFullYear());

  /* Note instruction */
  sh.getRange('A4:M4').merge()
    .setValue('▶  Renseignez le mois et l\'année ci-dessus, puis : menu  📊 Trading FR → Générer le calendrier')
    .setFontStyle('italic').setFontColor('#777777').setHorizontalAlignment('left');

  /* Largeurs colonnes jours */
  for (var c = 1; c <= 7; c++) sh.setColumnWidth(c, 125);
  sh.setColumnWidth(8, 155);   // totaux hebdo
  sh.setColumnWidth(11, 155);  // colonne K (totaux mensuels)

  sh.setFrozenRows(4);

  /* La grille est générée dynamiquement par genererCalendrier() */
}


/* ═══════════════════════════════════════════════════════════════════
   GÉNÉRER LE CALENDRIER MENSUEL
   ═══════════════════════════════════════════════════════════════════ */
function genererCalendrier() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var feuilleData = ss.getSheetByName(CFG.TRADES);
  var feuilleCal  = ss.getSheetByName(CFG.CALENDRIER);

  if (!feuilleData || !feuilleCal) {
    SpreadsheetApp.getUi().alert(
      'Feuilles introuvables.\nLancez d\'abord "🏗️ Initialiser le modèle" depuis le menu.'
    );
    return;
  }

  // Lire mois / année depuis B2 / B3 (ou B1/B2 si ancienne config)
  var mois  = feuilleCal.getRange('B2').getValue();
  var annee = feuilleCal.getRange('B3').getValue();
  if (!mois || !annee || mois < 1 || mois > 12 || annee < 2000 || annee > 2100) {
    SpreadsheetApp.getUi().alert('Mois ou année invalide.\nMois : 1-12 | Année : 2000-2100');
    return;
  }
  mois  = parseInt(mois,  10);
  annee = parseInt(annee, 10);

  // Effacer l'ancienne grille (rows 5-70)
  var DEBUT = CFG.CAL_GRILLE;
  feuilleCal.getRange('A' + DEBUT + ':M70').clear();
  feuilleCal.getRange('A' + DEBUT + ':M70').setBackground(null).setFontColor(null);

  var tradeMap          = _construireTradeMap(feuilleData);
  var premierJour       = new Date(annee, mois - 1, 1);
  var dernierJourNum    = new Date(annee, mois, 0).getDate();
  var jourDepartSemaine = (premierJour.getDay() + 6) % 7; // Lun=0 … Dim=6
  var jourCourant       = 1;
  var ligne             = DEBUT;

  /* Titre mois */
  feuilleCal.getRange(ligne, 1, 1, 8).merge()
    .setValue(MOIS[mois - 1].toUpperCase() + ' ' + annee)
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setBackground(CFG.BLEU_CLAIR);
  ligne++;

  /* En-têtes jours */
  feuilleCal.getRange(ligne, 1, 1, 8)
    .setValues([JOURS.concat(['Totaux\nHebdo'])])
    .setFontWeight('bold').setBackground(CFG.BLEU_CLAIR)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ligne++;

  var totalMoisTrades = 0;
  var totalMoisPL     = 0;

  /* Corps du calendrier */
  while (jourCourant <= dernierJourNum) {
    var res = _semaine(
      feuilleCal, tradeMap,
      jourCourant, dernierJourNum,
      jourDepartSemaine, ligne, mois, annee
    );
    jourCourant       = res.jourCourant;
    totalMoisTrades  += res.trades;
    totalMoisPL      += res.pl;
    jourDepartSemaine = 0;
    ligne++;
  }

  /* Totaux mensuels colonne K */
  _totauxMensuels(feuilleCal, mois, annee, totalMoisTrades, totalMoisPL);

  feuilleCal.setColumnWidth(8, 155);
  feuilleCal.setColumnWidth(11, 155);
  feuilleCal.autoResizeColumn(12);

  SpreadsheetApp.getUi().alert(
    '✅ Calendrier généré — ' + MOIS[mois - 1] + ' ' + annee
  );
}

/* ─────────────────────────────────────────────────────────────────
   Effacer la grille du calendrier
   ───────────────────────────────────────────────────────────────── */
function effacerCalendrier() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName(CFG.CALENDRIER);
  if (!sh) return;
  sh.getRange('A' + CFG.CAL_GRILLE + ':M70')
    .clear().setBackground(null).setFontColor(null);
}


/* ═══════════════════════════════════════════════════════════════════
   HELPERS PRIVÉS
   ═══════════════════════════════════════════════════════════════════ */

/** Construit { 'yyyy-MM-dd': {count, totalPL} } depuis la feuille trades. */
function _construireTradeMap(sh) {
  var last = sh.getLastRow();
  if (last < CFG.LIGNE_DEBUT) return {};
  var rows = last - CFG.LIGNE_DEBUT + 1;
  var dates = sh.getRange(CFG.LIGNE_DEBUT, CFG.COL_DATE, rows, 1).getValues();
  var pl    = sh.getRange(CFG.LIGNE_DEBUT, CFG.COL_PL,   rows, 1).getValues();
  var map   = {};
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    if (d instanceof Date && !isNaN(d)) {
      var k = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
      if (!map[k]) map[k] = {count: 0, totalPL: 0};
      map[k].count++;
      map[k].totalPL += Number(pl[i][0]) || 0;
    }
  }
  return map;
}

/** Affiche une semaine du calendrier. Retourne {jourCourant, trades, pl}. */
function _semaine(sh, map, jourCourant, dernierJour, offset, ligne, mois, annee) {
  var trades = 0, pl = 0, vals = [];
  for (var col = 0; col < 7; col++) {
    var weekend = (col === 5 || col === 6);
    if (col < offset || jourCourant > dernierJour) {
      vals.push('');
    } else {
      var ds    = annee + '-' + ('0'+mois).slice(-2) + '-' + ('0'+jourCourant).slice(-2);
      var t     = map[ds] || {count: 0, totalPL: 0};
      var texte = jourCourant + '\nTrades : ' + t.count + '\nP/L : ' + _euro(t.totalPL);
      vals.push(texte);
      var cell  = sh.getRange(ligne, col + 1);
      cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment('top').setHorizontalAlignment('left');
      if (weekend)         cell.setBackground(CFG.GRIS);
      else if (t.count > 0) cell.setBackground(t.totalPL >= 0 ? CFG.VERT_BG : CFG.ROUGE_BG);
      if      (t.totalPL > 0) cell.setFontColor(CFG.VERT_TXT);
      else if (t.totalPL < 0) cell.setFontColor(CFG.ROUGE_TXT);
      trades += t.count; pl += t.totalPL; jourCourant++;
    }
  }
  vals.push('Trades : ' + trades + '\nP/L : ' + _euro(pl));
  sh.getRange(ligne, 1, 1, 8).setValues([vals]);
  var cellH = sh.getRange(ligne, 8);
  cellH.setBackground(CFG.JAUNE_BG).setFontWeight('bold')
       .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top');
  if      (pl > 0) cellH.setFontColor(CFG.VERT_TXT);
  else if (pl < 0) cellH.setFontColor(CFG.ROUGE_TXT);
  else             cellH.setFontColor(CFG.NOIR);
  return {jourCourant: jourCourant, trades: trades, pl: pl};
}

/** Écrit le bloc totaux mensuels à partir de la colonne K. */
function _totauxMensuels(sh, mois, annee, totalTrades, totalPL) {
  var col = 11, r = 7;
  sh.getRange(r,   col).setValue('TOTAUX MENSUELS')
    .setFontWeight('bold').setFontSize(12).setBackground(CFG.BLEU_CLAIR);
  sh.getRange(r+1, col).setValue(MOIS[mois-1] + ' ' + annee).setFontStyle('italic');
  sh.getRange(r+2, col).setValue('──────────────');
  sh.getRange(r+3, col).setValue('Total trades :');
  sh.getRange(r+3, col+1).setValue(totalTrades).setFontWeight('bold');
  sh.getRange(r+4, col).setValue('Total P/L :');
  var cellPL = sh.getRange(r+4, col+1);
  cellPL.setValue(_euro(totalPL)).setFontWeight('bold');
  if      (totalPL > 0) cellPL.setFontColor(CFG.VERT_TXT);
  else if (totalPL < 0) cellPL.setFontColor(CFG.ROUGE_TXT);
  sh.getRange(r+5, col).setValue('Performance :');
  sh.getRange(r+5, col+1).setValue(
    totalPL > 0 ? '🟢 Mois positif' : totalPL < 0 ? '🔴 Mois négatif' : '⚪ Neutre'
  );
}

/** Formate un montant en euros : 1234.5 → "1 234,50 €" */
function _euro(v) {
  var s = v < 0 ? '-' : '';
  var a = Math.abs(v).toFixed(2).split('.');
  return s + a[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ',' + a[1] + ' €';
}

/** Retourne la feuille existante ou en crée une nouvelle. */
function _getOrCreate(ss, nom) {
  var sh = ss.getSheetByName(nom);
  if (!sh) sh = ss.insertSheet(nom);
  return sh;
}
