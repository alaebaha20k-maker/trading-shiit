/**
 * MODÈLE DE TRADING - CALENDRIER MENSUEL (Google Sheets Apps Script)
 *
 * Feuilles requises :
 *   - "Tous les trades"  : données brutes (dates en colonne B, P/L en colonne L, à partir de la ligne 3)
 *   - "Vue Calendrier"   : affichage du calendrier (mois en B1, année en B2)
 *
 * Comment utiliser :
 *   1. Copiez ce script dans Outils > Éditeur de script de votre Google Sheet.
 *   2. Créez les deux feuilles ci-dessus si elles n'existent pas.
 *   3. Saisissez le mois (1-12) en B1 et l'année (ex : 2024) en B2 dans "Vue Calendrier".
 *   4. Lancez genererCalendrier() depuis l'éditeur ou via le menu personnalisé.
 */

// ─────────────────────────────────────────────
//  Menu personnalisé
// ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 Trading FR')
    .addItem('Générer le calendrier', 'genererCalendrier')
    .addItem('Effacer le calendrier', 'effacerCalendrier')
    .addToUi();
}

// ─────────────────────────────────────────────
//  Constantes de configuration
// ─────────────────────────────────────────────
var CONFIG = {
  FEUILLE_TRADES     : 'Tous les trades',
  FEUILLE_CALENDRIER : 'Vue Calendrier',
  CELLULE_MOIS       : 'B1',
  CELLULE_ANNEE      : 'B2',
  LIGNE_DEBUT_DONNEES: 3,       // première ligne de données dans "Tous les trades"
  COL_DATE           : 2,       // colonne B  (1-indexed)
  COL_PL             : 12,      // colonne L  (1-indexed)
  LIGNE_DEBUT_GRILLE : 5,       // ligne où commence l'en-tête du calendrier
  COL_TOTAL_MENSUEL  : 11,      // colonne K  (totaux mensuels)

  // Couleurs
  COULEUR_EN_TETE    : '#d9e1f2',
  COULEUR_WEEK_END   : '#f2f2f2',
  COULEUR_GAIN       : '#e6ffe6',
  COULEUR_PERTE      : '#ffe6e6',
  COULEUR_TOTAL      : '#fff2cc',
  COULEUR_TEXTE_GAIN : '#1a7a1a',
  COULEUR_TEXTE_PERTE: '#cc0000',
};

// Jours en français (Lundi en premier — semaine ISO européenne)
var JOURS_SEMAINE = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'];

// Mois en français
var NOMS_MOIS = [
  'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
  'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'
];

// ─────────────────────────────────────────────
//  Fonction principale : genererCalendrier()
// ─────────────────────────────────────────────
function genererCalendrier() {
  var ss             = SpreadsheetApp.getActiveSpreadsheet();
  var feuilleData    = ss.getSheetByName(CONFIG.FEUILLE_TRADES);
  var feuilleCal     = ss.getSheetByName(CONFIG.FEUILLE_CALENDRIER);

  // ── Validation des feuilles ──
  if (!feuilleData || !feuilleCal) {
    SpreadsheetApp.getUi().alert(
      'Feuilles introuvables.\n' +
      'Veuillez créer "' + CONFIG.FEUILLE_TRADES + '" et "' + CONFIG.FEUILLE_CALENDRIER + '".'
    );
    return;
  }

  var mois  = feuilleCal.getRange(CONFIG.CELLULE_MOIS).getValue();
  var annee = feuilleCal.getRange(CONFIG.CELLULE_ANNEE).getValue();

  // ── Validation du mois / année ──
  if (!mois || !annee || mois < 1 || mois > 12 || annee < 2000 || annee > 2100) {
    SpreadsheetApp.getUi().alert(
      'Mois ou année invalide.\n' +
      'Mois : 1-12 | Année : 2000-2100'
    );
    return;
  }

  mois  = parseInt(mois,  10);
  annee = parseInt(annee, 10);

  // ── Effacement de l'ancienne grille ──
  feuilleCal.getRange('A' + CONFIG.LIGNE_DEBUT_GRILLE + ':M60').clear();
  feuilleCal.getRange('A' + CONFIG.LIGNE_DEBUT_GRILLE + ':M60').setBackground(null).setFontColor(null);

  // ── Construction du tradeMap ──
  var tradeMap = _construireTradeMap(feuilleData);

  // ── Paramètres du mois ──
  var premierJour    = new Date(annee, mois - 1, 1);
  var dernierJourNum = new Date(annee, mois, 0).getDate();

  // Lundi = 0 ... Dimanche = 6  (décalage ISO)
  var jourDepartSemaine = (premierJour.getDay() + 6) % 7;

  var jourCourant      = 1;
  var ligneActuelle    = CONFIG.LIGNE_DEBUT_GRILLE;

  // ── En-tête du titre ──
  _ecrireTitreMois(feuilleCal, ligneActuelle, mois, annee);
  ligneActuelle++;

  // ── En-têtes des jours ──
  var enTetes = JOURS_SEMAINE.concat(['Totaux\nHebdo']);
  feuilleCal.getRange(ligneActuelle, 1, 1, 8)
    .setValues([enTetes])
    .setFontWeight('bold')
    .setBackground(CONFIG.COULEUR_EN_TETE)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  ligneActuelle++;

  // ── Corps du calendrier ──
  var totalMoisTrades = 0;
  var totalMoisPL     = 0;

  while (jourCourant <= dernierJourNum) {
    var semaineResult = _calculerSemaine(
      feuilleCal, tradeMap,
      jourCourant, dernierJourNum,
      jourDepartSemaine, ligneActuelle,
      mois, annee
    );

    jourCourant       = semaineResult.jourCourant;
    totalMoisTrades  += semaineResult.tradesSemaine;
    totalMoisPL      += semaineResult.plSemaine;
    jourDepartSemaine = 0; // après la 1ʳᵉ semaine, on repart toujours du lundi
    ligneActuelle++;
  }

  // ── Totaux mensuels (colonne K) ──
  _ecrireTotauxMensuels(feuilleCal, mois, annee, totalMoisTrades, totalMoisPL);

  // ── Mise en forme finale des colonnes ──
  feuilleCal.setColumnWidth(8, 150);                          // Totaux hebdo
  feuilleCal.setColumnWidth(CONFIG.COL_TOTAL_MENSUEL, 160);   // Libellé K
  feuilleCal.autoResizeColumn(CONFIG.COL_TOTAL_MENSUEL + 1);  // Valeur L

  SpreadsheetApp.getUi().alert(
    '✅ Calendrier généré avec succès pour ' + NOMS_MOIS[mois - 1] + ' ' + annee + '.'
  );
}

// ─────────────────────────────────────────────
//  Effacer uniquement la grille
// ─────────────────────────────────────────────
function effacerCalendrier() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var feuilleCal = ss.getSheetByName(CONFIG.FEUILLE_CALENDRIER);
  if (!feuilleCal) return;
  feuilleCal.getRange('A' + CONFIG.LIGNE_DEBUT_GRILLE + ':M60')
    .clear()
    .setBackground(null)
    .setFontColor(null);
}

// ─────────────────────────────────────────────
//  Helpers privés
// ─────────────────────────────────────────────

/**
 * Construit un dictionnaire { 'yyyy-MM-dd': { count, totalPL } }
 * depuis la feuille de trades.
 */
function _construireTradeMap(feuilleData) {
  var derniereColonne = feuilleData.getLastRow();
  if (derniereColonne < CONFIG.LIGNE_DEBUT_DONNEES) return {};

  var dates = feuilleData
    .getRange(CONFIG.LIGNE_DEBUT_DONNEES, CONFIG.COL_DATE,
              derniereColonne - CONFIG.LIGNE_DEBUT_DONNEES + 1, 1)
    .getValues();
  var pl = feuilleData
    .getRange(CONFIG.LIGNE_DEBUT_DONNEES, CONFIG.COL_PL,
              derniereColonne - CONFIG.LIGNE_DEBUT_DONNEES + 1, 1)
    .getValues();

  var map = {};
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    if (d instanceof Date && !isNaN(d)) {
      var cle = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
      if (!map[cle]) map[cle] = { count: 0, totalPL: 0 };
      map[cle].count++;
      map[cle].totalPL += Number(pl[i][0]) || 0;
    }
  }
  return map;
}

/**
 * Écrit la ligne de titre du mois (fusionnée sur 8 colonnes).
 */
function _ecrireTitreMois(feuilleCal, ligne, mois, annee) {
  var plage = feuilleCal.getRange(ligne, 1, 1, 8);
  plage.merge();
  plage.setValue(NOMS_MOIS[mois - 1].toUpperCase() + ' ' + annee);
  plage.setFontWeight('bold')
       .setFontSize(14)
       .setHorizontalAlignment('center')
       .setBackground(CONFIG.COULEUR_EN_TETE);
}

/**
 * Calcule et affiche une semaine du calendrier.
 * Retourne { jourCourant, tradesSemaine, plSemaine }.
 */
function _calculerSemaine(
  feuilleCal, tradeMap,
  jourCourant, dernierJourNum,
  offsetDebut, ligne,
  mois, annee
) {
  var tradesSemaine = 0;
  var plSemaine     = 0;
  var valeursCells  = [];

  for (var col = 0; col < 7; col++) {
    var estWeekEnd = (col === 5 || col === 6); // Sam = 5, Dim = 6

    if (col < offsetDebut || jourCourant > dernierJourNum) {
      // Cellule vide (avant le 1ᵉʳ jour ou après la fin du mois)
      valeursCells.push('');
    } else {
      var dateStr = annee + '-' +
                    ('0' + mois).slice(-2) + '-' +
                    ('0' + jourCourant).slice(-2);
      var trade   = tradeMap[dateStr] || { count: 0, totalPL: 0 };
      var contenu = jourCourant +
                    '\nTrades : ' + trade.count +
                    '\nP/L : ' + _formatEuro(trade.totalPL);

      valeursCells.push(contenu);

      // Style de la cellule
      var cell = feuilleCal.getRange(ligne, col + 1);
      cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .setVerticalAlignment('top')
          .setHorizontalAlignment('left');

      if (estWeekEnd) {
        cell.setBackground(CONFIG.COULEUR_WEEK_END);
      } else if (trade.count > 0) {
        cell.setBackground(trade.totalPL >= 0 ? CONFIG.COULEUR_GAIN : CONFIG.COULEUR_PERTE);
      }

      if (trade.totalPL > 0)      cell.setFontColor(CONFIG.COULEUR_TEXTE_GAIN);
      else if (trade.totalPL < 0) cell.setFontColor(CONFIG.COULEUR_TEXTE_PERTE);

      tradesSemaine += trade.count;
      plSemaine     += trade.totalPL;
      jourCourant++;
    }
  }

  // Colonne 8 : totaux hebdomadaires
  var contenuHebdo = 'Trades : ' + tradesSemaine +
                     '\nP/L : '  + _formatEuro(plSemaine);
  valeursCells.push(contenuHebdo);

  feuilleCal.getRange(ligne, 1, 1, 8).setValues([valeursCells]);

  var cellHebdo = feuilleCal.getRange(ligne, 8);
  cellHebdo.setBackground(CONFIG.COULEUR_TOTAL)
           .setFontWeight('bold')
           .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
           .setVerticalAlignment('top');
  if (plSemaine > 0)      cellHebdo.setFontColor(CONFIG.COULEUR_TEXTE_GAIN);
  else if (plSemaine < 0) cellHebdo.setFontColor(CONFIG.COULEUR_TEXTE_PERTE);
  else                    cellHebdo.setFontColor('black');

  return { jourCourant: jourCourant, tradesSemaine: tradesSemaine, plSemaine: plSemaine };
}

/**
 * Affiche le bloc "TOTAUX MENSUELS" dans la colonne K.
 */
function _ecrireTotauxMensuels(feuilleCal, mois, annee, totalTrades, totalPL) {
  var col     = CONFIG.COL_TOTAL_MENSUEL;
  var ligneD  = 7; // démarre à K7 pour laisser de l'espace

  // Titre
  feuilleCal.getRange(ligneD, col)
    .setValue('TOTAUX MENSUELS')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground(CONFIG.COULEUR_EN_TETE);

  // Nom du mois
  feuilleCal.getRange(ligneD + 1, col)
    .setValue(NOMS_MOIS[mois - 1] + ' ' + annee)
    .setFontStyle('italic');

  // Séparateur visuel
  feuilleCal.getRange(ligneD + 2, col).setValue('──────────────');

  // Total trades
  feuilleCal.getRange(ligneD + 3, col).setValue('Total trades :');
  feuilleCal.getRange(ligneD + 3, col + 1).setValue(totalTrades).setFontWeight('bold');

  // Total P/L
  feuilleCal.getRange(ligneD + 4, col).setValue('Total P/L :');
  var cellPL = feuilleCal.getRange(ligneD + 4, col + 1);
  cellPL.setValue(_formatEuro(totalPL))
        .setFontWeight('bold');
  if (totalPL > 0)      cellPL.setFontColor(CONFIG.COULEUR_TEXTE_GAIN);
  else if (totalPL < 0) cellPL.setFontColor(CONFIG.COULEUR_TEXTE_PERTE);
  else                  cellPL.setFontColor('black');

  // Performance
  feuilleCal.getRange(ligneD + 5, col).setValue('Performance :');
  var perf = totalPL > 0 ? '🟢 Mois positif' : (totalPL < 0 ? '🔴 Mois négatif' : '⚪ Neutre');
  feuilleCal.getRange(ligneD + 5, col + 1).setValue(perf);
}

/**
 * Formate un montant en euros avec 2 décimales.
 * Ex : 1234.5 → "1 234,50 €"
 */
function _formatEuro(montant) {
  var signe  = montant < 0 ? '-' : '';
  var abs    = Math.abs(montant).toFixed(2);
  var parts  = abs.split('.');
  var entier = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  return signe + entier + ',' + parts[1] + ' €';
}
