const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, PageBreak, VerticalAlign, TableLayoutType
} = require('docx');
const fs = require('fs');

const W = 9360;
const COL2 = [4680, 4680];
const COL3 = [3120, 3120, 3120];

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const cellMargin = { top: 100, bottom: 100, left: 140, right: 140 };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size: 36, font: 'Arial', color: '1a1a1a' })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, font: 'Arial', color: '2c2c2c' })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, font: 'Arial', color: '3d3d3d' })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 100 },
    children: [new TextRun({ text, size: 22, font: 'Arial', color: '444444', ...opts })]
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: 'Arial', color: '444444' })]
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'DDDDDD' } },
    children: []
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function label(text, color = '888888') {
  return new TextRun({ text, size: 18, font: 'Arial', color, bold: true });
}

function cell(content, opts = {}) {
  const { fill = 'FFFFFF', w = 4680, bold = false, color = '333333', fontSize = 22, header = false } = opts;
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: cellMargin,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: [new TextRun({ text: content, size: fontSize, font: 'Arial', bold: bold || header, color })]
    })]
  });
}

function table2(rows, headers, fills = []) {
  const tableRows = [];
  if (headers) {
    tableRows.push(new TableRow({
      children: headers.map((h, i) => cell(h, { fill: 'F0F0F0', w: COL2[i], bold: true, color: '222222' }))
    }));
  }
  rows.forEach((row, ri) => {
    tableRows.push(new TableRow({
      children: row.map((c, i) => cell(c, { fill: fills[ri] || 'FFFFFF', w: COL2[i] }))
    }));
  });
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: COL2,
    rows: tableRows
  });
}

function table3(rows, headers) {
  const tableRows = [];
  if (headers) {
    tableRows.push(new TableRow({
      children: headers.map((h, i) => cell(h, { fill: 'F0F0F0', w: COL3[i], bold: true, color: '222222' }))
    }));
  }
  rows.forEach(row => {
    tableRows.push(new TableRow({
      children: row.map((c, i) => cell(c, { w: COL3[i] }))
    }));
  });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: COL3, rows: tableRows });
}

function tableN(rows, colWidths, headers) {
  const tableRows = [];
  if (headers) {
    tableRows.push(new TableRow({
      children: headers.map((h, i) => cell(h, { fill: 'F0F0F0', w: colWidths[i], bold: true, color: '222222' }))
    }));
  }
  rows.forEach(row => {
    tableRows.push(new TableRow({
      children: row.map((c, i) => cell(c, { w: colWidths[i] }))
    }));
  });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: colWidths, rows: tableRows });
}

function noteBox(text, fillColor = 'FFF8E6') {
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [W],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: W, type: WidthType.DXA },
        shading: { fill: fillColor, type: ShadingType.CLEAR },
        margins: cellMargin,
        children: [new Paragraph({ children: [new TextRun({ text, size: 20, font: 'Arial', color: '5c4a00', italics: true })] })]
      })]
    })]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 260 } } }
        }, {
          level: 1, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 900, hanging: 260 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: 'Arial', color: '1a1a1a' },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: '2c2c2c' },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [

      // ─── COVER ───────────────────────────────────────────────────────────────
      new Paragraph({ spacing: { before: 1200, after: 80 }, children: [new TextRun({ text: 'GROW CREATURES', bold: true, size: 56, font: 'Arial', color: '1a1a1a' })] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: 'Game Design Document', size: 28, font: 'Arial', color: '666666' })] }),
      new Paragraph({ spacing: { before: 0, after: 400 }, children: [new TextRun({ text: 'Roblox — Grow / Idle + Collection — 8 joueurs / serveur', size: 22, font: 'Arial', color: '888888' })] }),
      divider(),
      new Paragraph({ spacing: { before: 80, after: 40 }, children: [new TextRun({ text: 'Version  1.0  |  Document de conception complet', size: 20, font: 'Arial', color: 'AAAAAA' })] }),
      pageBreak(),

      // ─── 1. VISION ───────────────────────────────────────────────────────────
      h1('1. Vision & Positionnement'),
      h2('1.1 Concept central'),
      p('Grow Creatures est un jeu Roblox de type Grow / Idle + Collection dans lequel 8 joueurs partagent un serveur. Le joueur produit des ressources, ameliore sa production, debloque de nouvelles dimensions et collecte des creatures de rarete croissante.'),
      p(''),
      h2('1.2 Pourquoi ce genre fonctionne'),
      table3(
        [
          ['Progression visible', 'Les chiffres montent en temps reel. Le cerveau valide chaque seconde.', 'Retention immédiate — D1'],
          ['Collection', 'Completer un roster de creatures = satisfaction profonde.', 'Retention long terme — D7+'],
          ['Retour quotidien', 'Une amelioration en attente = raison de se reconnecterle lendemain.', 'Driver du D7 retention'],
        ],
        ['Pilier', 'Mecanique', 'Impact KPI']
      ),
      p(''),
      h2('1.3 KPI cibles'),
      table2(
        [
          ['D1 Retention', '> 40% — fun immediat en moins de 30 secondes'],
          ['D7 Retention', '> 20% — objectif de collection visible et atteignable'],
          ['D30 Retention', '> 8% — systeme de rebirth et pass saisonnier actifs'],
          ['ARPU mois 1', '0.15 - 0.30 USD par joueur actif (standard Roblox idle)'],
        ],
        ['Metrique', 'Cible & logique']
      ),
      p(''),
      divider(),

      // ─── 2. BOUCLE DE JEU ────────────────────────────────────────────────────
      h1('2. Boucle de Jeu'),
      h2('2.1 Macro-loop (session a session)'),
      table2(
        [
          ['Etape 1 — Commencer pauvre', 'Production lente, creature de base. Le joueur comprend vite qu\'il faut s\'ameliorer.'],
          ['Etape 2 — Produire', 'Recolte passive ou clics. Compteur visible en temps reel.'],
          ['Etape 3 — Ameliorer', 'Vitesse, multiplicateur, capacite. Effet visible dans les 2 secondes.'],
          ['Etape 4 — Produire plus vite', 'Boucle de retroaction positive. Le joueur sent la difference immediatement.'],
          ['Etape 5 — Debloquer une dimension', 'Nouveau biome, nouvelles creatures. Pic d\'excitation.'],
          ['Etape 6 — Collectionner', 'RNG controle. Chance augmentee avec multiplicateurs.'],
          ['Etape 7 — Rebirth / Prestige', 'Reset avec multiplicateur permanent. La boucle repart plus vite.'],
        ],
        ['Etape', 'Description']
      ),
      p(''),
      h2('2.2 Micro-loop (en session)'),
      p('Clic > ressource > achat > feedback visuel > nouveau palier > decision. Cycle cible : 45-90 secondes par palier au debut du jeu, jamais plus de 10 minutes au mid-game sans possibilite de rebirth.'),
      p(''),
      h2('2.3 Regle des 30 secondes'),
      noteBox('CRITIQUE : le joueur doit voir un nombre changer dans les 30 premieres secondes. Si l\'ecran est statique au spawn, il part. Pas d\'exception.'),
      p(''),
      h2('2.4 Rebirth system'),
      p('Chaque rebirth donne des Essence Tokens permanents. Ces tokens achetent des multiplicateurs globaux ou debloquent des creatures exclusives. Objectif de design : au moins 1 rebirth par session longue (45-60 min).'),
      p(''),
      divider(),

      // ─── 3. PREMIERE SESSION ─────────────────────────────────────────────────
      h1('3. Premiere Session — Minute par Minute'),
      p('Le jeu doit creer une addiction en 10 minutes, pas en 1 heure. Voici la sequence precise avec les wow moments et les portes de monetisation naturelles.'),
      p(''),
      tableN(
        [
          ['0:00 – 0:30', 'Spawn & premier choc visuel', 'Les creatures des autres joueurs sont visibles immediatement. Premiere creature donnee automatiquement — 0 menu, 0 friction.', 'Wow #1 — 0 friction'],
          ['0:30 – 2:00', 'Premiere production', 'Compteur monte en temps reel. Tutoriel contextuel : une fleche vers "Ameliorer". Un seul bouton visible.', 'Progression visible'],
          ['2:00 – 3:30', 'Premier upgrade', 'La creature change visuellement : grandit, animation differente, production double. Son satisfaisant. Animation 1.5 sec minimum.', 'Wow #2 — feedback sensoriel'],
          ['3:30 – 5:00', 'Premiere creature rare', 'Une creature Uncommon droppee. Couleur distincte. Notification serveur : "[Joueur] a obtenu Crystal Pup !"', 'Wow #3 — jalousie sociale'],
          ['5:00 – 8:00', 'Objectif moyen terme', 'Zone 2 affichee avec distance visible. Temps estime affiche. Le joueur sait ou il va.', 'Tension narrative'],
          ['8:00 – 12:00', 'Friction douce + porte mono', 'Progression 30% plus lente. Icone VIP discrete avec tooltip. Auras VIP visibles sur les autres joueurs.', 'Porte de monetisation naturelle'],
          ['12:00 – 18:00', 'Deblocage Zone 2', 'Transition animee : sol, musique, creatures changes. Premier objectif long terme accompli.', 'Wow #4 — D1 retention hook'],
          ['18:00 – 25:00', 'Premiere fusion', 'Animation 3 secondes, spectaculaire. Les autres joueurs voient le resultat. L\'addiction collection demarre.', 'Wow #5 — hook collection'],
          ['25:00 – 30:00', 'Rebirth visible + D2 hook', 'Joueur Rebirth visible dans le serveur avec aura x10. Tooltip : "Le Rebirth donne un multiplicateur permanent".', 'D2 retention hook'],
        ],
        [1800, 2200, 3960, 1400],
        ['Temps', 'Etape', 'Description', 'Tag design']
      ),
      p(''),
      noteBox('Les 5 wow moments DOIVENT etre spectaculaires visuellement et sonores. C\'est 50% de la sensation d\'addiction. Ne pas economiser sur les animations de ces moments cles.'),
      p(''),
      divider(),

      // ─── 4. DIMENSIONS & CREATURES ───────────────────────────────────────────
      h1('4. Dimensions & Creatures'),
      h2('4.1 Vue d\'ensemble'),
      p('Chaque dimension est un biome distinct avec une esthetique, une musique et un type de creature propres. La desirabilite augmente avec chaque dimension, creant une hierarchie sociale naturelle visible entre les joueurs.'),
      p(''),
      table2(
        [
          ['Dimension 1 — Meadow Brainrot', 'Starter. Creatures meme & brainrot internet. Accessible des le depart. Desirabilite : 55%.'],
          ['Dimension 2 — Football Arena', 'Mid-game. Creatures inspirees de joueurs de foot (archétypes, pas licences reelles). Factions Rouge/Bleu. Desirabilite : 72%.'],
          ['Dimension 3 — Anime Nexus', 'Zone avancee. Creatures style anime et manga (archétypes originaux). Communaute la plus fidelisee. Desirabilite : 88%.'],
          ['Dimension 4 — Void Realm', 'Post-rebirth uniquement. Creatures cosmiques avec auras permanentes visibles depuis toute la map. Desirabilite : 98%.'],
        ],
        ['Dimension', 'Description']
      ),
      p(''),
      h2('4.2 Roster MVP — Dimension 1 (Meadow Brainrot)'),
      tableN(
        [
          ['Skibidi Pup', 'Common', '70%', '1 energie/s', 'Chien tete WC. Creature de depart.'],
          ['Rizz Cat', 'Common', '70%', '2 energie/s', 'Chat avec lunettes de soleil.'],
          ['Gyatt Frog', 'Uncommon', '18%', '6 energie/s', 'Grenouille meme — premier rare a 5 min.'],
          ['NPC Golem', 'Rare', '8%', '15 energie/s', 'Golem visage NPC. Notification discrete.'],
          ['Sigma Wolf', 'Epic', '2.5%', '40 energie/s', 'Loup sigma. Annonce serveur + animation.'],
          ['Mr Beast Beast', 'Legendary', '0.5%', '150 energie/s', 'Creature style MrBeast. Annonce globale.'],
          ['Trollface Ultimate', 'Secret', '0.01%', '600 energie/s', 'Notifie TOUT le serveur. Mecanique unique.'],
        ],
        [1800, 1400, 800, 1400, 3960],
        ['Creature', 'Rarete', 'Taux drop', 'Production', 'Notes']
      ),
      p(''),
      h2('4.3 Roster MVP — Dimension 2 (Football Arena)'),
      tableN(
        [
          ['Rookie Striker', 'Common', '70%', '5 energie/s', 'Attaquant debutant.'],
          ['Golden Boot', 'Uncommon', '18%', '18 energie/s', 'Botte doree.'],
          ['El Toro', 'Rare', '8%', '45 energie/s', 'Taureau en crampons.'],
          ['CR7 Goat', 'Epic', '2.5%', '120 energie/s', 'Chevre avec jersey 7. Faction Rouge.'],
          ['The Ballon God', 'Legendary', '0.5%', '400 energie/s', 'Dieu du ballon. Annonce globale.'],
        ],
        [1800, 1400, 800, 1400, 3960],
        ['Creature', 'Rarete', 'Taux drop', 'Production', 'Notes']
      ),
      p('Mecanique unique Dimension 2 : 2 factions (Rouge / Bleu). La faction majoritaire dans le serveur donne un bonus de production a tous ses membres.'),
      p(''),
      h2('4.4 Roster MVP — Dimension 3 (Anime Nexus)'),
      tableN(
        [
          ['Chibi Blade', 'Common', '70%', '20 energie/s', 'Guerrier chibi.'],
          ['Shadow Fox', 'Uncommon', '18%', '60 energie/s', 'Renard demon.'],
          ['Spirit Kitsune', 'Rare', '8%', '150 energie/s', 'Kitsune 3 queues.'],
          ['Demon King', 'Epic', '2.5%', '400 energie/s', 'Roi demon avec cornes.'],
          ['Celestial Dragon', 'Legendary', '0.5%', '1200 energie/s', 'Dragon celeste.'],
          ['??? (Secret)', 'Secret', '0.005%', '5000 energie/s', 'Inconnu jusqu\'au premier drop mondial.'],
        ],
        [1800, 1400, 800, 1400, 3960],
        ['Creature', 'Rarete', 'Taux drop', 'Production', 'Notes']
      ),
      p(''),
      h2('4.5 Dimension 4 — Void Realm (post-rebirth)'),
      tableN(
        [
          ['Void Pup', 'Rare', 'Post-rebirth', '500 energie/s', 'Version cosmique de Skibidi Pup.'],
          ['Dark Ballon God', 'Epic', 'Post-rebirth', '1500 energie/s', 'Version corrompue du Ballon God.'],
          ['Void Dragon', 'Legendary', 'Post-rebirth', '5000 energie/s', 'Dragon cosmique avec particules permanentes.'],
          ['The One', 'Secret', '0.001%', '20 000 energie/s', '1 seul par monde simultanement. Si quelqu\'un drop le sien, le tien disparait.'],
        ],
        [1800, 1400, 1200, 1400, 3560],
        ['Creature', 'Rarete', 'Acces', 'Production', 'Notes']
      ),
      p(''),
      h2('4.6 Systeme de rarete — signaux visuels'),
      table2(
        [
          ['Common', 'Pas d\'effet visuel'],
          ['Uncommon', 'Contour subtil de couleur'],
          ['Rare', 'Particules legeres autour de la creature'],
          ['Epic', 'Aura permanente animee'],
          ['Legendary', 'Aura + effet lumineux au sol'],
          ['Secret', 'Visible depuis toute la map + annonce mondiale'],
        ],
        ['Rarete', 'Signal visuel']
      ),
      p(''),
      h2('4.7 Fusion system'),
      p('2 creatures identiques + ressources suffisantes = fusion vers le tier superieur. Animation : 3 secondes, sons impactants, particules, visible pour les autres joueurs du serveur. La fusion est le moment le plus satisfaisant visuellement du jeu. Ne pas economiser sur cette animation.'),
      p(''),
      h2('4.8 Collection completee — bonus permanents'),
      table2(
        [
          ['Dimension 1 complete', 'Aura "Brainrot Master" + x1.5 production permanente'],
          ['Dimension 2 complete', 'Badge "Football God" sur le profil + x2 production'],
          ['Dimension 3 complete', 'Creature bonus exclusive "Anime Lord" + x3 production'],
          ['Dimension 4 complete', 'Titre "The Collector" — 1 seul par serveur'],
        ],
        ['Condition', 'Recompense']
      ),
      p(''),
      divider(),

      // ─── 5. ZONES / PROGRESSION ──────────────────────────────────────────────
      h1('5. Structure des Zones'),
      tableN(
        [
          ['Zone 1', 'Meadow Starter', 'Depart', '—', 'Tutoriel. 5 creatures. Debloque par defaut.'],
          ['Zone 2', 'Crystal Caves', 'Min 12-18', '500 energie', 'Premiers rares brillants. Partage social naturel.'],
          ['Zone 3', 'Lava Realm', 'Min 40-60', '5 000 cristaux', 'Fusion uniquement. Milestone naturel pour l\'upsell VIP.'],
          ['Zone 4', 'Void Realm', 'Post-rebirth', '1 Rebirth', 'Exclusif post-rebirth. Creatures cosmiques.'],
          ['Zone 5+', 'Seasonal Zones', 'Evenements', 'Passe saison.', 'Zones temporaires (Halloween, Noel, ete). Driver du passe.'],
        ],
        [700, 1600, 1200, 1400, 4460],
        ['#', 'Nom', 'Acces (session 1)', 'Cout deblocage', 'Contenu & role']
      ),
      p(''),
      noteBox('Regle de timing : Zone 2 doit etre accessible entre 12 et 20 minutes de jeu en session 1. Trop tot (<8 min) = pas assez de satisfaction. Trop tard (>30 min) = le joueur est deja parti.'),
      p(''),
      divider(),

      // ─── 6. MULTIJOUEUR ──────────────────────────────────────────────────────
      h1('6. Multijoueur — 8 Joueurs / Serveur'),
      h2('6.1 Philosophie'),
      p('Competitif doux, jamais PvP direct. Les 8 joueurs partagent le meme espace mais ne se bloquent pas. La competition passe par la comparaison sociale (leaderboard, creatures visibles), pas par le sabotage. Resultat attendu : moins de toxicite, sessions plus longues.'),
      p(''),
      h2('6.2 Features multijoueur'),
      table2(
        [
          ['Leaderboard serveur', 'Classement temps reel par richesse ou tier de creature. Mise a jour toutes les 10 secondes. Visible en permanence dans le HUD.'],
          ['Shared Boost Events', 'Si le serveur atteint un objectif collectif (ex : 1M produit ensemble), tous les joueurs recoivent un bonus x2 pendant 2 minutes.'],
          ['Creatures visibles', 'Les creatures rares des autres joueurs s\'affichent dans le monde. Animation speciale quand une Legendary apparait.'],
          ['Factions (D2)', 'Faction Rouge vs Bleu en Dimension 2. La faction majoritaire dans le serveur donne un bonus de production passif.'],
          ['Trade (V2)', 'Echange de creatures entre joueurs du meme serveur. Limite : 1 trade par heure. A implementer apres le MVP.'],
        ],
        ['Feature', 'Description']
      ),
      p(''),
      h2('6.3 Gestion de la disparite de niveaux'),
      p('Probleme classique : un nouveau joueur rejoint un serveur avec des joueurs en Dimension 4. Il se sent ecrase et part avant 10 minutes.'),
      p('Solution implementee : 3 couches visuelles toujours presentes dans chaque serveur :'),
      bullet('Creatures de D1 (communes) — le nouveau joueur se voit dedans'),
      bullet('Creatures de D2/D3 (cool) — le nouvel objectif visible'),
      bullet('Creatures de D4 (show-off uniquement) — la carotte lointaine'),
      p('L\'auto-matchmaking place les nouveaux joueurs en priorite dans des serveurs avec d\'autres debutants.'),
      p(''),
      h2('6.4 Architecture serveur'),
      bullet('Serveurs prives optionnels (Roblox natif) pour groupes d\'amis'),
      bullet('DataStore sauvegarde cote serveur uniquement — pas de client-side trust'),
      bullet('Anti-exploit obligatoire des le MVP : validation ServerSide de chaque transaction'),
      bullet('Les jeux idle sont les plus hackes de Roblox — le leaderboard est corrompu en 48h sans protection'),
      p(''),
      divider(),

      // ─── 7. MONETISATION ─────────────────────────────────────────────────────
      h1('7. Monetisation'),
      h2('7.1 Philosophie'),
      p('Objectif realiste : 2-5% des joueurs achetent quelque chose. Priorite absolue = volume de joueurs, pas taux de conversion. Un jeu bien equilibre avec 100 000 joueurs/mois genere plus qu\'un jeu P2W avec 10 000 joueurs.'),
      p(''),
      h2('7.2 Gamepasses (lancement)'),
      tableN(
        [
          ['VIP Pass', '399 R$', 'Gamepass one-time', 'x1.5 production, tag VIP, aura visible, salon VIP', 'Priorite 1'],
          ['x2 Money', '299 R$', 'Gamepass one-time', 'Double la production passive', 'Priorite 1'],
          ['Auto Collect', '199 R$', 'Gamepass one-time', 'Recolte auto sans clic (QoL fort)', 'Priorite 2'],
        ],
        [1800, 1000, 1600, 3000, 1160] ,
        ['Produit', 'Prix', 'Type', 'Avantage', 'Priorite']
      ),
      p(''),
      h2('7.3 Contenu post-lancement'),
      tableN(
        [
          ['Passe Saisonnier', '499 R$', 'Mensuel', '30 paliers, creatures exclusives, cosmétiques, emotes'],
          ['Skins creatures', '149-299 R$', 'Shop cosmetique', 'Apparence uniquement, 0 avantage gameplay'],
          ['Skip timer', '49-99 R$', 'Microtransaction', 'A doser — max 1 timer par zone. Jamais obligatoire.'],
        ],
        [1800, 1000, 1200, 5360],
        ['Produit', 'Prix', 'Type', 'Detail']
      ),
      p(''),
      h2('7.4 A NE JAMAIS faire'),
      bullet('Vendre des creatures rares directement (detruit l\'economie de collection)'),
      bullet('Plus de 5 gamepasses visibles simultanement (confusion = non-achat)'),
      bullet('Progression trop lente pour forcer l\'achat (frustration = quit)'),
      bullet('Popups de monetisation interrompant le gameplay'),
      bullet('Avantages P2W qui rendent le jeu injouable en F2P'),
      p(''),
      noteBox('La monetisation se vend elle-meme : les auras VIP et les skins rares sont visibles par tous les joueurs depuis le spawn. La jalousie sociale fait le travail de vente — pas les popups.'),
      p(''),
      divider(),

      // ─── 8. ROADMAP ──────────────────────────────────────────────────────────
      h1('8. Roadmap de Developpement'),
      h2('8.1 MVP — Semaines 1 a 6'),
      p('Objectif : valider que le jeu retient les joueurs 7 jours avant de developper davantage de contenu.'),
      table2(
        [
          ['Zone 1 complete', '5 creatures avec animations, systeme de production, 3 upgrades par creature'],
          ['Leaderboard serveur', 'Classement basique par richesse, mis a jour toutes les 10 secondes'],
          ['2 gamepasses', 'VIP Pass + x2 Money uniquement au lancement'],
          ['Systeme de sauvegarde', 'DataStore robuste avec anti-exploit ServerSide'],
          ['Shared Boost Events', 'Evenement collectif simple pour creer de la cohesion serveur'],
        ],
        ['Feature MVP', 'Detail']
      ),
      p(''),
      h2('8.2 V1 — Semaines 6 a 10'),
      table2(
        [
          ['Zone 2 — Crystal Caves', 'Nouveau biome + 5 nouvelles creatures + systeme de fusion'],
          ['Rebirth system', 'Essence Tokens + multiplicateurs permanents'],
          ['Auto Collect gamepass', 'Troisieme gamepass + skip timers optionnels'],
          ['Factions Football', 'Systeme Rouge/Bleu en Dimension 2'],
        ],
        ['Feature V1', 'Detail']
      ),
      p(''),
      h2('8.3 V2 — Semaines 10 a 16'),
      table2(
        [
          ['Premier passe saisonnier', 'Event saisonnier + zone temporaire + 30 paliers de recompenses'],
          ['Shop cosmetique', 'Skins pour creatures existantes, 0 avantage gameplay'],
          ['Zone 3 + Void Realm', 'Dimension 3 Anime + Dimension 4 post-rebirth'],
          ['Systeme de trade', 'Echange entre joueurs du meme serveur, limite 1 trade/heure'],
        ],
        ['Feature V2', 'Detail']
      ),
      p(''),
      divider(),

      // ─── 9. CE QUI N'A PAS ETE COUVERT ──────────────────────────────────────
      h1('9. Sujets Non Couverts — A Traiter'),
      p('Les elements suivants ont ete identifies comme manquants dans la conception actuelle et necessitent un travail de design supplementaire.'),
      p(''),
      table2(
        [
          ['Viralite Roblox', 'Thumbnail, titre, tags SEO. 80% des joueurs viennent de la. Levier le plus sous-estime.'],
          ['Strategie influenceurs', 'Early access a 5 YouTubeurs Roblox mid-tier. Souvent plus efficace qu\'un budget pub.'],
          ['Daily quests & missions', 'Objectifs quotidiens pour garantir le retour J2-J7. Sans ca, le D7 depend uniquement du contenu.'],
          ['Economie numerique', 'Chiffres concrets : production creature par niveau, cout des upgrades, ratio progression/temps.'],
          ['Anti-exploit detail', 'Validation ServerSide de chaque transaction. Les jeux idle sont les plus hackes de Roblox.'],
          ['Design sonore', 'Sons d\'upgrade, fusion, drop rare. 30% de la sensation d\'addiction.'],
          ['Systeme de quetes', 'Quetes journalieres, hebdomadaires, de collection. Driver de session sans nouveau contenu.'],
          ['Balancing curve', 'Courbe mathematique de progression. A modeliser sur tableur avant de coder.'],
          ['Thumbnail & SEO Roblox', 'Titre, description, thumbnail A/B testing. Separe du game design mais critique pour la decouverte.'],
        ],
        ['Sujet', 'Pourquoi c\'est important']
      ),
      p(''),
      divider(),

      // ─── 10. RESUME EXECUTIF ─────────────────────────────────────────────────
      h1('10. Resume Executif'),
      table2(
        [
          ['Genre', 'Grow / Idle + Collection'],
          ['Plateforme', 'Roblox'],
          ['Joueurs / serveur', '8 joueurs en parallele'],
          ['Dimensions au MVP', '1 (Meadow Brainrot)'],
          ['Dimensions totales', '4 + zones saisonnieres'],
          ['Creatures au MVP', '8 (Zone 1 complete)'],
          ['Creatures totales (V2)', '30+'],
          ['Gamepasses au lancement', '2 (VIP + x2 Money)'],
          ['KPI prioritaire', 'D7 Retention > 20%'],
          ['Mecanique virale cle', 'Creatures rares visibles par tout le serveur'],
          ['Driver de monetisation', 'Jalousie sociale par les auras et skins'],
        ],
        ['Parametre', 'Valeur']
      ),
      p(''),
      p('La regle d\'or du MVP : ne pas construire la Zone 3 avant de savoir si la Zone 1 retient les joueurs 7 jours. Les analytics D1/D7 sont la seule metrique qui compte lors du lancement.'),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/GrowCreatures_GDD.docx', buf);
  console.log('Done');
});
