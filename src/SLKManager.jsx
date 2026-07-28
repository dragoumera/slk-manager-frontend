import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import { Plus, Trash2, LayoutDashboard, FileText, Library, AlertTriangle, Clock, Wallet, ChevronRight, Building2, Settings, CheckCircle2, XCircle, ShieldCheck, UserCog, Users, Truck, UserPlus, Bell, MapPin, TrendingUp, Camera, BellRing, MessageCircle, Send, Receipt, Menu, X } from "lucide-react";
import { api } from "./api.js";

const LOGO_SRC = "data:image/png;base64,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";

// ---------------------------------------------------------------------------
// Design tokens
// ---------------------------------------------------------------------------
const INK = "#0F1B2E";        // near-black navy, primary text on light
const NAVY = "#12203D";       // sidebar / brand dark
const NAVY_SOFT = "#1B2E52";  // sidebar hover
const STEEL = "#3E5C82";      // secondary text on navy
const ACCENT = "#E8622C";     // safety-orange accent (construction cue)
const ACCENT_DEEP = "#C2501F";
const BG = "#F4F5F7";
const SURFACE = "#FFFFFF";
const BORDER = "#E5E8EE";
const MUTE = "#8891A0";
const GOOD = "#1F7A4D";
const BAD = "#B23A1A";
const SUBMENU_ACCENT = "#4FA3D1"; // bleu clair — distingue visuellement les sous-menus des menus principaux (orange)

// ---------------------------------------------------------------------------
// Seed data — bibliotheque de prix reelle (devis SIG, chantier CPSSD)
// ---------------------------------------------------------------------------
const seedLibrary = [
  { id: "l1", code: "01-05", categorie: "Gaine rectangulaire", designation: "Gaine rectangulaire galva (Fo+Mo)", unite: "KG", prix: 6.8, tempsPose: null },
  { id: "l2", code: "01-05b", categorie: "Gaine rectangulaire", designation: "Gaine rectangulaire tole noire (Fo+Mo)", unite: "KG", prix: 9.8, tempsPose: null },
  { id: "l3", code: "01-01", categorie: "Gaine circulaire", designation: "Gaine circulaire galva D125 (Fo+Mo)", unite: "ML", prix: 36.5, tempsPose: null },
  { id: "l4", code: "01-02", categorie: "Gaine circulaire", designation: "Gaine circulaire galva D160 (Fo+Mo)", unite: "ML", prix: 41.4, tempsPose: null },
  { id: "l5", code: "01-03", categorie: "Gaine circulaire", designation: "Gaine circulaire galva D200 (Fo+Mo)", unite: "ML", prix: 44.85, tempsPose: null },
  { id: "l6", code: "01-06", categorie: "Gaine circulaire", designation: "Gaine circulaire galva D250 (Fo+Mo)", unite: "ML", prix: 54.05, tempsPose: 0.2 },
  { id: "l7", code: "01-07", categorie: "Gaine circulaire", designation: "Gaine circulaire galva D315 (Fo+Mo)", unite: "ML", prix: 63.2, tempsPose: null },
  { id: "l8", code: "01-08", categorie: "Gaine circulaire", designation: "Gaine circulaire galva D400 (Fo+Mo)", unite: "ML", prix: 82.6, tempsPose: null },
  { id: "l9", code: "01-04", categorie: "Accessoires", designation: "Gaine souple (fourniture + pose)", unite: "ML", prix: 25.0, tempsPose: null },
  { id: "l10", code: "05-01", categorie: "Pose", designation: "Pose registre circulaire", unite: "UTE", prix: 50.0, tempsPose: null },
  { id: "l11", code: "04-08", categorie: "Pose", designation: "Pose plenum + grille", unite: "UTE", prix: 70.0, tempsPose: null },
  { id: "l12", code: "10-03", categorie: "Pose", designation: "Pose de bouche", unite: "UTE", prix: 45.0, tempsPose: null },
  { id: "l13", code: "05-03", categorie: "Pose", designation: "Pose CCF circulaire", unite: "UTE", prix: 70.0, tempsPose: null },
  { id: "l14", code: "05-05", categorie: "Pose", designation: "Pose BDV", unite: "UTE", prix: 45.0, tempsPose: null },
  { id: "l15", code: "05-04", categorie: "Pose", designation: "Pose CCF rectangulaire", unite: "UTE", prix: 70.0, tempsPose: null },
  { id: "l16", code: "05-02", categorie: "Pose", designation: "Pose registre rectangulaire", unite: "UTE", prix: 55.0, tempsPose: null },
  { id: "l17", code: "03-11", categorie: "Accessoires", designation: "Fourniture et pose trappe de visite", unite: "UTE", prix: 45.0, tempsPose: null },
  { id: "l18", code: "04-10", categorie: "Pose", designation: "Pose batterie", unite: "UTE", prix: 70.0, tempsPose: null },
  { id: "l19", code: "03-12", categorie: "Accessoires", designation: "Fourniture et pose grand PAS rectangulaire", unite: "ENS", prix: 360.0, tempsPose: null },
  { id: "l20", code: "01-09", categorie: "Accessoires", designation: "Pose canal air", unite: "UTE", prix: 90.0, tempsPose: null },
  { id: "l21", code: "05-08", categorie: "Pose", designation: "Pose clapet anti-retour", unite: "UTE", prix: 110.0, tempsPose: null },
];

const seedClients = [
  { id: "cl1", nom: "AXIMA Concept Equans", adresse: "Nanterre, France", telephone: "01 41 XX XX XX", contact: "Service travaux" },
  { id: "cl2", nom: "SCI Rocher Vienne", adresse: "Paris, France", telephone: "01 42 XX XX XX", contact: "M. Vienne" },
  { id: "cl3", nom: "SCI Frepillon Invest", adresse: "Frepillon, France", telephone: "01 34 XX XX XX", contact: "Mme Dubois" },
  { id: "cl4", nom: "Logistia SAS", adresse: "Zone industrielle des Linandes, Cergy", telephone: "01 30 XX XX XX", contact: "M. Rousseau" },
  { id: "cl5", nom: "COGET Batiment", adresse: "Osny, France", telephone: "01 34 XX XX XX", contact: "Mme Kacem" },
];

const seedFournisseurs = [
  { id: "f1", nom: "SRV GAINE", specialite: "Fabrication gaine de ventilation", telephone: "03 XX XX XX XX" },
];

const LOGIN_USERS = [
  { id: "u1", nom: "Sankanou Lasano", poste: "Gerant", role: "direction", roleLabel: "Direction", initiales: "SL", email: "sankanou.lasano@slkclim.fr" },
  { id: "u2", nom: "Joaquim Ribeiro", poste: "Co-gerant", role: "direction", roleLabel: "Direction", initiales: "JR", email: "joaquim.ribeiro@slkclim.fr" },
  { id: "u3", nom: "Semega Bakaty", poste: "Chef de chantier", role: "chef_chantier", roleLabel: "Chef de chantier", initiales: "SB", email: "semega.bakaty@slkclim.fr" },
  { id: "u4", nom: "Fatoumata Coulibaly", poste: "Chef de chantier", role: "chef_chantier", roleLabel: "Chef de chantier", initiales: "FC", email: "fatoumata.coulibaly@slkclim.fr" },
  { id: "u5", nom: "Ibrahim Traore", poste: "Ouvrier", role: "ouvrier", roleLabel: "Ouvrier", initiales: "IT", email: "ibrahim.traore@slkclim.fr" },
  { id: "u6", nom: "Moussa Diarra", poste: "Ouvrier", role: "ouvrier", roleLabel: "Ouvrier", initiales: "MD", email: "moussa.diarra@slkclim.fr" },
];

// Un chantier n'est visible que par la Direction et par les agents listes
// dans `affectations` (leur id LOGIN_USERS) — c'est la meme liste d'agents
// que celle utilisee pour la connexion, plus de decalage possible entre
// "qui est affecte" et "qui se connecte".
const ANNEE_EN_COURS = 2026;

const seedChantiers = [
  { id: "c1", nom: "Chantier SIG - AXIMA Concept Equans", clientId: "cl1", affectations: ["u3", "u5"], devisStatut: "accepte", montantHT: 327117.83, heuresConsommees: 480, statut: "en_cours", anneeCreation: 2026 },
  { id: "c2", nom: "Chantier CPSSD - Bat QF", clientId: "cl2", affectations: ["u3", "u6"], devisStatut: "accepte", montantHT: 196047.3, heuresConsommees: 1420, statut: "en_cours", anneeCreation: 2026 },
  { id: "c3", nom: "Rehabilitation bureaux - Frepillon", clientId: "cl3", affectations: ["u4"], devisStatut: "accepte", montantHT: 84500, heuresConsommees: 610, statut: "en_difficulte", anneeCreation: 2026 },
  { id: "c4", nom: "Extension entrepot - Cergy", clientId: "cl4", affectations: ["u3"], devisStatut: "accepte", montantHT: 152300, heuresConsommees: 210, statut: "en_cours", anneeCreation: 2026 },
  { id: "c5", nom: "Reprise reseau VMC - Osny", clientId: "cl5", affectations: [], devisStatut: "accepte", montantHT: 38900, heuresConsommees: 0, statut: "en_cours", anneeCreation: 2026 },
  // Demonstration de la regle "annee en cours" : ce chantier date de 2025 mais
  // n'est pas termine -> il continue d'apparaitre partout jusqu'a sa cloture.
  { id: "c6", nom: "Renovation CVC - Bureaux Levallois", clientId: "cl2", affectations: ["u4"], devisStatut: "accepte", montantHT: 61200, heuresConsommees: 540, statut: "en_cours", anneeCreation: 2025 },
  // Celui-ci est de 2025 ET termine -> il n'apparait plus dans les vues
  // "en cours" (dashboard, plan, productivite...), seulement dans le Bilan
  // annuel 2025 et l'historique de facturation.
  { id: "c7", nom: "Extension atelier - Argenteuil", clientId: "cl4", affectations: ["u3"], devisStatut: "accepte", montantHT: 45900, heuresConsommees: 720, statut: "termine", anneeCreation: 2025 },
];

// Un chantier "appartient" a l'annee en cours s'il a ete cree cette annee-la,
// OU s'il vient d'une annee anterieure mais n'est pas encore termine (il
// continue alors d'apparaitre dans le suivi courant jusqu'a sa cloture).
// Un chantier termine d'une annee anterieure bascule dans l'historique de
// cette annee-la (Bilan annuel) et ne s'affiche plus dans les vues courantes.
// Regle appliquee de facon identique partout (tableau de bord, plan,
// productivite, heures, messagerie, comptabilite analytique).
function chantierAppartientAnneeEnCours(chantier, anneeRef = ANNEE_EN_COURS) {
  const annee = chantier.anneeCreation || anneeRef;
  if (annee >= anneeRef) return true;
  return chantier.statut !== "termine";
}

function chantiersActifs(utilisateur, chantiers, anneeRef = ANNEE_EN_COURS) {
  return chantiersVisibles(utilisateur, chantiers).filter((c) => chantierAppartientAnneeEnCours(c, anneeRef));
}

// chantierId ajoute explicitement (en plus du nom, garde pour affichage) afin
// de pouvoir filtrer les saisies par chantier reellement affecte.
const seedHeures = [
  { id: "h1", chantierId: "c1", chantier: "Chantier SIG - AXIMA Concept Equans", ouvrier: "Ibrahim Traore", heures: 7.5, poste: "Gaine D250 - zone AS-Bureau", statut: "a_valider", commentaire: "", quand: "Il y a 12 min" },
  { id: "h2", chantierId: "c1", chantier: "Chantier SIG - AXIMA Concept Equans", ouvrier: "Moussa Diarra", heures: 8, poste: "Pose de bouche", statut: "a_valider", commentaire: "", quand: "Il y a 45 min" },
  { id: "h3", chantierId: "c4", chantier: "Extension entrepot - Cergy", ouvrier: "Ibrahim Traore", heures: 6, poste: "Gaine rectangulaire", statut: "acceptee", commentaire: "", quand: "Hier" },
  { id: "h4", chantierId: "c2", chantier: "Chantier CPSSD - Bat QF", ouvrier: "Moussa Diarra", heures: 8, poste: "Gaine souple isolee", statut: "acceptee", commentaire: "", quand: "Hier" },
  { id: "h5", chantierId: "c3", chantier: "Rehabilitation bureaux - Frepillon", ouvrier: "Ibrahim Traore", heures: 4, poste: "Pose registre circulaire", statut: "validee", commentaire: "", quand: "Il y a 3 jours" },
  { id: "h6", chantierId: "c1", chantier: "Chantier SIG - AXIMA Concept Equans", ouvrier: "Moussa Diarra", heures: 5, poste: "Gaine D250 - zone AR Bureau", statut: "refusee", commentaire: "Heures incoherentes avec le planning du jour, a corriger.", quand: "Il y a 4 jours" },
];

const seedPlanElements = [
  { id: "e1", chantierId: "c1", designation: "Gaine D250 - zone AS-Bureau", statut: "pose", quantitePrevue: 40, quantitePosee: 40 },
  { id: "e2", chantierId: "c1", designation: "Gaine D250 - zone AR-Bureau", statut: "en_cours", quantitePrevue: 35, quantitePosee: 18 },
  { id: "e3", chantierId: "c1", designation: "Pose de bouches - niveau 1", statut: "non_commence", quantitePrevue: 12, quantitePosee: 0 },
  { id: "e4", chantierId: "c1", designation: "CCF circulaire - gaine technique", statut: "bloque", quantitePrevue: 4, quantitePosee: 0 },
  { id: "e5", chantierId: "c2", designation: "Gaine rectangulaire - hall principal", statut: "en_cours", quantitePrevue: 800, quantitePosee: 520 },
  { id: "e6", chantierId: "c2", designation: "Registres circulaires", statut: "pose", quantitePrevue: 10, quantitePosee: 10 },
  { id: "e7", chantierId: "c4", designation: "Gaine souple isolee - stockage", statut: "non_commence", quantitePrevue: 45, quantitePosee: 0 },
];

// Module 7 — Gestion des plans PDF. Un meme "numero de plan" peut avoir
// plusieurs indices de revision dans le temps : une seule version est
// "active" a la fois, les precedentes restent consultables en historique.
const seedPlansPdf = [
  { id: "pl1", chantierId: "c2", planNumero: "100", indice: "IND D", actif: true, dateImport: "2026-06-02", importePar: "Semega Bakaty", fichierUrl: null },
  { id: "pl2", chantierId: "c2", planNumero: "100", indice: "IND C", actif: false, dateImport: "2026-05-12", importePar: "Sankanou Lasano", fichierUrl: null },
  { id: "pl3", chantierId: "c2", planNumero: "101", indice: "IND C", actif: true, dateImport: "2026-06-02", importePar: "Semega Bakaty", fichierUrl: null },
  { id: "pl4", chantierId: "c1", planNumero: "1", indice: "IND A", actif: true, dateImport: "2026-04-20", importePar: "Sankanou Lasano", fichierUrl: null },
];

const COULEUR_STATUT = {
  non_commence: { label: "Non commence", color: "#8891A0", bg: "#F1F2F4", dot: "#B7BCC5" },
  en_cours: { label: "En cours", color: "#B5710A", bg: "#FDF3E2", dot: "#E8622C" },
  pose: { label: "Pose", color: "#1F7A4D", bg: "#E9F7EF", dot: "#1F7A4D" },
  bloque: { label: "Bloque", color: "#B23A1A", bg: "#FBEBE5", dot: "#B23A1A" },
};

// chantierId ajoute pour permettre au chef de chantier de ne voir que les
// notifications de ses ouvriers SUR SES PROPRES chantiers affectes.
const seedNotifications = [
  { id: "n1", chantierId: "c1", pourRole: "chef_chantier", pourNom: "Semega Bakaty", type: "heure_a_valider", message: "Nouvelle saisie d'heures a valider sur Chantier SIG - AXIMA Concept Equans (7.5 h).", lu: false, quand: "Il y a 12 min" },
  { id: "n2", chantierId: "c1", pourRole: "chef_chantier", pourNom: "Semega Bakaty", type: "heure_a_valider", message: "Nouvelle saisie d'heures a valider sur Chantier SIG - AXIMA Concept Equans (8 h).", lu: false, quand: "Il y a 45 min" },
  { id: "n3", chantierId: "c3", pourRole: "direction", pourNom: null, type: "depassement_budget", message: "Rehabilitation bureaux - Frepillon a depasse 90% des heures prevues.", lu: false, quand: "Il y a 2 h" },
  { id: "n4", chantierId: "c2", pourRole: "direction", pourNom: null, type: "nouveau_plan", message: "Un nouveau plan a ete importe pour Chantier CPSSD - Bat QF.", lu: true, quand: "Hier" },
  { id: "n5", chantierId: "c1", pourRole: "ouvrier", pourNom: "Ibrahim Traore", type: "heure_refusee", message: "Votre saisie du 24/07 a ete refusee : heures incoherentes avec le planning du jour.", lu: false, quand: "Il y a 1 j" },
  { id: "n6", chantierId: "c2", pourRole: "ouvrier", pourNom: "Moussa Diarra", type: "heure_validee", message: "Votre saisie du 22/07 (8h) a ete validee par la Direction.", lu: true, quand: "Il y a 2 j" },
];

// Chantiers visibles pour un utilisateur donne : la Direction voit tout ;
// un chef de chantier ou un ouvrier ne voit QUE les chantiers ou il est
// explicitement affecte (voir DashboardTab / affecterAgent).
function chantiersVisibles(utilisateur, chantiers) {
  if (utilisateur.role === "direction") return chantiers;
  return chantiers.filter((c) => c.affectations.includes(utilisateur.id));
}

// Module Communication interne — messages hierarchises entre agents.
const seedMessages = [
  { id: "m1", deId: "u3", versId: "u5", chantierId: "c1", sujet: "Avancement gaine D250 - zone AR-Bureau",
    message: "J'ai vu que la pose de la gaine D250 zone AR-Bureau est bloquee a 18/35. Peux-tu m'expliquer ce qui retarde la suite ?", lu: true, quand: "Hier 16:20", necessiteReponse: true, repondu: true, repondA: null },
  { id: "m2", deId: "u5", versId: "u3", chantierId: "c1", sujet: "RE: Avancement gaine D250 - zone AR-Bureau",
    message: "On attend la livraison des coudes galva, prevue demain matin. On reprend des reception.", lu: false, quand: "Hier 17:05", necessiteReponse: false, repondu: false, repondA: "m1" },
  { id: "m3", deId: "u1", versId: "u4", chantierId: "c3", sujet: "Retard chantier Rehabilitation bureaux",
    message: "Le chantier est passe en alerte (avancement > 90% des heures prevues). Merci de faire un point precis sur la situation.", lu: false, quand: "Il y a 3 h", necessiteReponse: true, repondu: false, repondA: null },
];

// Communication hierarchisee (voir cahier des charges, matrice de droits) :
// - Direction peut ecrire a tout le monde
// - Un chef de chantier peut ecrire a la Direction et aux agents affectes
//   sur SES chantiers (son equipe), pas a l'equipe d'un autre chef
// - Un ouvrier ne peut ecrire qu'aux chefs de chantier de SES chantiers,
//   jamais directement a un autre ouvrier ni a la Direction
// Communication hierarchisee STRICTE, pour l'initiation de nouveaux
// messages : seule la Direction peut ECRIRE a un chef de chantier, et seul
// un chef de chantier peut ECRIRE a un ouvrier. Un chef ou un ouvrier ne
// peuvent pas initier de nouveau message vers le niveau au-dessus : ils
// REPONDENT (voir MessagerieTab, bouton "Repondre"), ce qui reste obligatoire
// pour les messages marques necessiteReponse.
function correspondantsAutorises(utilisateur, utilisateursSysteme, chantiers) {
  if (utilisateur.role === "direction") {
    return utilisateursSysteme.filter((u) => u.role === "chef_chantier");
  }
  if (utilisateur.role === "chef_chantier") {
    const mesChantiers = chantiersActifs(utilisateur, chantiers);
    const idsEquipe = new Set();
    mesChantiers.forEach((c) => c.affectations.forEach((id) => idsEquipe.add(id)));
    idsEquipe.delete(utilisateur.id);
    return utilisateursSysteme.filter((u) => idsEquipe.has(u.id) && u.role === "ouvrier");
  }
  // ouvrier : ne peut initier aucun nouveau message, seulement repondre
  return [];
}

// ---------------------------------------------------------------------------
// Module 20 — Comptabilite. Valeurs de depart INDICATIVES (coefficient de
// charges patronales, montants) — a confirmer avec un expert-comptable avant
// utilisation reelle (voir section 20.9 du CDC). Ce logiciel ne genere ni
// bulletin de paie ni declaration DSN : il sert au SUIVI, pas a la paie legale.
const seedFactures = [
  { id: "fa1", chantierId: "c1", numero: "FA-2026-0031", type: "acompte", montantHT: 98135.35, montantTVA: 19627.07, statutPaiement: "payee", emiseLe: "2026-06-02", echeanceLe: "2026-07-02" },
  { id: "fa2", chantierId: "c2", numero: "FA-2026-0042", type: "facture", montantHT: 58814.19, montantTVA: 0, statutPaiement: "en_retard", emiseLe: "2026-06-20", echeanceLe: "2026-07-20" },
  { id: "fa3", chantierId: "c3", numero: "FA-2026-0045", type: "facture", montantHT: 84500, montantTVA: 16900, statutPaiement: "emise", emiseLe: "2026-07-10", echeanceLe: "2026-08-10" },
];
const seedPaiements = [
  { id: "pa1", factureId: "fa1", montant: 117762.42, mode: "virement", date: "2026-06-15" },
];
const seedRemunerations = [
  { id: "re1", userId: "u3", mode: "salaire_fixe", salaireMensuelBrut: 2800, tauxHoraireBrut: null, coeffCharges: 42 },
  { id: "re2", userId: "u4", mode: "salaire_fixe", salaireMensuelBrut: 2700, tauxHoraireBrut: null, coeffCharges: 42 },
  { id: "re3", userId: "u5", mode: "taux_horaire", salaireMensuelBrut: null, tauxHoraireBrut: 18, coeffCharges: 42 },
  { id: "re4", userId: "u6", mode: "taux_horaire", salaireMensuelBrut: null, tauxHoraireBrut: 17.5, coeffCharges: 42 },
];
const seedCharges = [
  { id: "ch1", libelle: "Loyer atelier / depot", categorie: "Local", montantMensuel: 1800 },
  { id: "ch2", libelle: "Assurance flotte vehicules", categorie: "Assurance", montantMensuel: 620 },
  { id: "ch3", libelle: "Expert-comptable", categorie: "Services", montantMensuel: 450 },
  { id: "ch4", libelle: "Assurance decennale / RC pro", categorie: "Assurance", montantMensuel: 390 },
];
const MODE_REMUNERATION_LABEL = { salaire_fixe: "Salaire fixe", taux_horaire: "Taux horaire", forfait_projet: "Forfait par chantier" };

// Bilan annuel — repartition MENSUELLE illustrative (CA facture et cout de
// revient), pour demontrer le comparatif N / N-1. A brancher sur les vraies
// dates de facturation une fois le logiciel connecte a la base reelle : ces
// chiffres ne sont PAS derives des chantiers de demonstration ci-dessus.
const MOIS_LABELS = ["Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre"];
const seedBilanMensuel = {
  2025: [
    { ca: 58000, cout: 49800 }, { ca: 62000, cout: 53200 }, { ca: 71000, cout: 61500 }, { ca: 68000, cout: 58800 },
    { ca: 75000, cout: 64500 }, { ca: 82000, cout: 71200 }, { ca: 79000, cout: 68600 }, { ca: 54000, cout: 47200 },
    { ca: 88000, cout: 76800 }, { ca: 91000, cout: 78900 }, { ca: 85000, cout: 73600 }, { ca: 96000, cout: 82900 },
  ],
  2026: [
    { ca: 65000, cout: 55400 }, { ca: 70000, cout: 59600 }, { ca: 84000, cout: 71300 }, { ca: 91000, cout: 76900 },
    { ca: 98000, cout: 82500 }, { ca: 112000, cout: 93900 }, { ca: 105000, cout: 88300 },
    { ca: null, cout: null }, { ca: null, cout: null }, { ca: null, cout: null }, { ca: null, cout: null }, { ca: null, cout: null },
  ],
};

const STATUT_INFO = {
  a_valider: { label: "A valider", color: "#8891A0", bg: "#F1F2F4" },
  acceptee: { label: "Acceptee", color: "#1F6FB2", bg: "#E9F2FB" },
  validee: { label: "Validee", color: "#1F7A4D", bg: "#E9F7EF" },
  refusee: { label: "Refusee", color: "#B23A1A", bg: "#FBEBE5" },
};

const PCT_DEFAUT_INIT = { fourniture: 60, frais: 15, mainOeuvre: 25 };
const TAUX_HORAIRE_DEFAUT_INIT = 40;
const SEUIL_ALERTE_INIT = 90;
const COEFFICIENT_CONSOMMABLES_INIT = 2.5;
const REMISE_DEFAUT_INIT = 0;

// Hypothese "1 jour = 16h" — VALIDEE le 26/07/2026 a partir du chiffrage reel
// du chantier CPSSD Bat QF (188h / 11,75j = 16,0 h/jour, verifie sur plusieurs
// sections du document transmis par SLK Clim).
const HEURES_PAR_JOUR = 16;

const fmtEUR = (n) => n.toLocaleString("fr-FR", { style: "currency", currency: "EUR", maximumFractionDigits: 0 });
const fmtH = (n) => n.toLocaleString("fr-FR", { maximumFractionDigits: 1 }) + " h";
const fmtJours = (n) => n.toLocaleString("fr-FR", { maximumFractionDigits: 2 }) + " j";

// Convertit une couleur hex (#RRGGBB) en rgba avec une opacite donnee —
// utilise pour les fonds de badges/icones colores de facon coherente
// partout dans le logiciel (KPI, groupes de menu, etc.).
function hexAlpha(hex, alpha) {
  const h = hex.replace("#", "");
  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function heuresPrevues(montantHT, pct, taux) {
  const montantMO = (montantHT * pct.mainOeuvre) / 100;
  return taux > 0 ? montantMO / taux : 0;
}

// ---------------------------------------------------------------------------
export default function SLKManagerPrototype() {
  const [tab, setTab] = useState("dashboard");
  const [menuMobileOuvert, setMenuMobileOuvert] = useState(false);
  const [library, setLibrary] = useState(seedLibrary);
  const [chantiers, setChantiers] = useState(seedChantiers);
  const [clients, setClients] = useState(seedClients);
  const [fournisseurs, setFournisseurs] = useState(seedFournisseurs);
  const [heuresListe, setHeuresListe] = useState(seedHeures);
  const [planElements, setPlanElements] = useState(seedPlanElements);
  const [plansPdf, setPlansPdf] = useState(seedPlansPdf);

  // Import d'une nouvelle revision de plan : desactive l'ancien indice pour
  // ce meme numero de plan (une seule version active a la fois), garde
  // l'historique complet.
  function importerPlanPdf(chantierId, planNumero, indice, file) {
    const fichierUrl = file ? URL.createObjectURL(file) : null;
    setPlansPdf((prev) => [
      { id: "pl" + Date.now(), chantierId, planNumero, indice, actif: true, dateImport: new Date().toISOString().slice(0, 10), importePar: utilisateur.nom, fichierUrl },
      ...prev.map((p) => (p.chantierId === chantierId && p.planNumero === planNumero ? { ...p, actif: false } : p)),
    ]);
  }

  // Envoie la fiche de chantier (adresse, chef, equipe) a chaque agent
  // affecte, sous forme de notification prioritaire — independant de la
  // messagerie hierarchisee, puisqu'il s'agit d'une information officielle
  // diffusee par la Direction/le chef, pas d'un echange interpersonnel.
  function envoyerFicheAffectation(chantierId) {
    const chantier = chantiers.find((c) => c.id === chantierId);
    if (!chantier) return;
    const client = clients.find((cl) => cl.id === chantier.clientId);
    const agents = chantier.affectations.map((id) => utilisateursSysteme.find((u) => u.id === id)).filter(Boolean);
    const chefs = agents.filter((a) => a.role === "chef_chantier").map((a) => a.nom).join(", ") || "non affecte";
    const ouvriers = agents.filter((a) => a.role === "ouvrier").map((a) => a.nom).join(", ") || "aucun";
    const adresse = client && client.adresse ? client.adresse : "adresse non renseignee";

    setNotifications((prev) => [
      ...agents.map((a) => ({
        id: "n" + Date.now() + a.id,
        chantierId,
        pourRole: a.role,
        pourNom: a.nom,
        type: "affectation_chantier",
        message: "Fiche de chantier " + chantier.nom + " (" + adresse + ") — Chef : " + chefs + " — Equipe : " + ouvriers + ".",
        lu: false,
        quand: "a l'instant",
      })),
      ...prev,
    ]);
  }

  const [historiqueAvancement, setHistoriqueAvancement] = useState([]);

  function changerStatutElement(elementId, statut) {
    const element = planElements.find((e) => e.id === elementId);
    if (!element || element.statut === statut) return;
    setPlanElements((prev) => prev.map((e) => (e.id === elementId ? { ...e, statut } : e)));
    setHistoriqueAvancement((prev) => [
      { id: "hist" + Date.now(), chantierId: element.chantierId, designation: element.designation, avant: element.statut, apres: statut, par: utilisateur.nom, quand: new Date().toLocaleString("fr-FR") },
      ...prev,
    ]);
  }
  const [notifications, setNotifications] = useState(seedNotifications);
  const [messages, setMessages] = useState(seedMessages);
  const [factures, setFactures] = useState(seedFactures);
  const [paiements, setPaiements] = useState(seedPaiements);
  const [remunerations, setRemunerations] = useState(seedRemunerations);
  const [charges, setCharges] = useState(seedCharges);

  function enregistrerPaiement(factureId, montant, mode) {
    setPaiements((prev) => [...prev, { id: "pa" + Date.now(), factureId, montant, mode, date: new Date().toISOString().slice(0, 10) }]);
    setFactures((prev) => prev.map((f) => {
      if (f.id !== factureId) return f;
      const dejaPaye = paiements.filter((p) => p.factureId === factureId).reduce((s, p) => s + p.montant, 0) + montant;
      const totalDu = f.montantHT + f.montantTVA;
      const statutPaiement = dejaPaye >= totalDu ? "payee" : dejaPaye > 0 ? "payee_partiellement" : "emise";
      return { ...f, statutPaiement };
    }));
  }
  function ajouterCharge(charge) {
    setCharges((prev) => [...prev, { id: "ch" + Date.now(), ...charge }]);
  }
  function supprimerCharge(id) {
    setCharges((prev) => prev.filter((c) => c.id !== id));
  }
  function definirRemuneration(userId, remuneration) {
    setRemunerations((prev) => [...prev.filter((r) => r.userId !== userId), { id: "re" + Date.now(), userId, ...remuneration }]);
  }

  function envoyerMessage({ versId, chantierId, sujet, message, necessiteReponse, repondA }) {
    const id = "m" + Date.now();
    setMessages((prev) => [...prev, { id, deId: utilisateur.id, versId, chantierId: chantierId || null, sujet, message, lu: false, quand: "a l'instant", necessiteReponse: !!necessiteReponse, repondu: false, repondA: repondA || null }]);
    if (repondA) {
      setMessages((prev) => prev.map((m) => (m.id === repondA ? { ...m, repondu: true } : m)));
    }
    const destinataire = utilisateursSysteme.find((u) => u.id === versId);
    if (destinataire) {
      setNotifications((prev) => [
        { id: "n" + Date.now(), chantierId: chantierId || null, pourRole: destinataire.role, pourNom: destinataire.nom, type: "message_recu", message: (necessiteReponse ? "[Reponse requise] " : "") + "Nouveau message de " + utilisateur.nom + " : " + sujet, lu: false, quand: "a l'instant" },
        ...prev,
      ]);
    }
  }
  function marquerMessageLu(id) {
    setMessages((prev) => prev.map((m) => (m.id === id ? { ...m, lu: true } : m)));
  }
  // Clic sur une notification liee a un chantier : renvoie directement sur
  // la partie concernee (Plan & avancement de ce chantier), pas seulement
  // sur la messagerie — comme demande.
  const [planChantierCible, setPlanChantierCible] = useState(null);
  const [comptaSousOngletCible, setComptaSousOngletCible] = useState(null);
  function ouvrirPartieConcernee(chantierId) {
    if (chantierId) {
      setPlanChantierCible(chantierId);
      setTab("plan");
    } else {
      setTab("messagerie");
    }
  }
  // Connexion reelle active : l'utilisateur doit s'authentifier via l'API.
  // (verifie reste geree par l'OTP optionnel ; ici on la laisse a true car la
  // verification par email n'est pas activee dans cette version branchee API.)
  const [verifie, setVerifie] = useState(true);
  const [utilisateursSysteme, setUtilisateursSysteme] = useState(LOGIN_USERS);
  const [utilisateur, setUtilisateur] = useState(null);

  // Parametres globaux (Module 1/2/8) — modifiables uniquement ici, jamais
  // codes en dur dans la logique de calcul (voir section 6.1bis du CDC).
  const [pctDefaut, setPctDefaut] = useState(PCT_DEFAUT_INIT);
  const [tauxDefaut, setTauxDefaut] = useState(TAUX_HORAIRE_DEFAUT_INIT);
  const [seuilAlerte, setSeuilAlerte] = useState(SEUIL_ALERTE_INIT);
  const [coeffConsommables, setCoeffConsommables] = useState(COEFFICIENT_CONSOMMABLES_INIT);
  const [remiseDefaut, setRemiseDefaut] = useState(REMISE_DEFAUT_INIT);

  const [devisNom, setDevisNom] = useState("");
  const [devisClientId, setDevisClientId] = useState("");
  const [devisSociete, setDevisSociete] = useState("");
  const [taux, setTaux] = useState(tauxDefaut);
  const [pct, setPct] = useState(pctDefaut);
  const [remise, setRemise] = useState(remiseDefaut);
  const [lignes, setLignes] = useState([]);
  const [sections, setSections] = useState([]);
  const [sectionActiveId, setSectionActiveId] = useState(null);

  const montantTotal = useMemo(() => lignes.reduce((s, l) => s + l.quantite * l.prix, 0), [lignes]);
  const montantRemise = (montantTotal * remise) / 100;
  const montantApresRemise = montantTotal - montantRemise;
  const montantFourniture = (montantTotal * pct.fourniture) / 100;
  const montantFrais = (montantTotal * pct.frais) / 100;
  const montantMO = (montantTotal * pct.mainOeuvre) / 100;
  const heures = heuresPrevues(montantTotal, pct, taux);
  const jours = heures / HEURES_PAR_JOUR;

  function addLigne(item) {
    setLignes((prev) => [...prev, { id: item.id + "-" + Date.now(), itemId: item.id, code: item.code, designation: item.designation, unite: item.unite, prix: item.prix, quantite: 1, sectionId: sectionActiveId }]);
  }
  function updateQte(id, q) {
    setLignes((prev) => prev.map((l) => (l.id === id ? { ...l, quantite: Math.max(0, q) } : l)));
  }
  function removeLigne(id) {
    setLignes((prev) => prev.filter((l) => l.id !== id));
  }
  function ajouterSection(section) {
    const id = "sec" + Date.now();
    setSections((prev) => [...prev, { id, ...section }]);
    setSectionActiveId(id);
    return id;
  }
  function supprimerSection(id) {
    setSections((prev) => prev.filter((s) => s.id !== id));
    setLignes((prev) => prev.map((l) => (l.sectionId === id ? { ...l, sectionId: null } : l)));
    if (sectionActiveId === id) setSectionActiveId(null);
  }
  function enregistrerDevis() {
    if (!devisNom || !devisClientId || montantTotal <= 0) return;
    const chantierId = "c" + Date.now();
    setChantiers((prev) => [
      { id: chantierId, nom: devisNom, clientId: devisClientId, societe: devisSociete || null, affectations: [], devisStatut: "accepte", montantHT: montantTotal, heuresConsommees: 0, statut: "en_cours" },
      ...prev,
    ]);
    // Devis accepte -> facture d'acompte generee automatiquement (30%, voir
    // conditions standard SLK Clim : acompte a la commande, solde a reception).
    const montantAcompteHT = Math.round(montantTotal * 0.30 * 100) / 100;
    const tva = Math.round(montantAcompteHT * 0.20 * 100) / 100;
    setFactures((prev) => [
      { id: "fa" + Date.now(), chantierId, numero: "FA-" + new Date().getFullYear() + "-" + String(prev.length + 1).padStart(4, "0"),
        type: "acompte", montantHT: montantAcompteHT, montantTVA: tva, statutPaiement: "emise",
        emiseLe: new Date().toISOString().slice(0, 10), echeanceLe: new Date(Date.now() + 30 * 86400000).toISOString().slice(0, 10) },
      ...prev,
    ]);
    setDevisNom("");
    setDevisClientId("");
    setDevisSociete("");
    setLignes([]);
    setSections([]);
    setSectionActiveId(null);
    setTab("dashboard");
  }

  // Cloture d'un chantier : passage a "termine" declenche automatiquement la
  // facture de solde (montant restant = total - ce qui a deja ete facture).
  function changerStatutChantier(chantierId, statut) {
    setChantiers((prev) => prev.map((c) => (c.id === chantierId ? { ...c, statut } : c)));
    if (statut !== "termine") return;
    const chantier = chantiers.find((c) => c.id === chantierId);
    if (!chantier) return;
    const dejaFacture = factures.filter((f) => f.chantierId === chantierId).reduce((s, f) => s + f.montantHT, 0);
    const resteAFacturer = Math.round((chantier.montantHT - dejaFacture) * 100) / 100;
    if (resteAFacturer <= 0) return;
    const tva = Math.round(resteAFacturer * 0.20 * 100) / 100;
    setFactures((prev) => [
      { id: "fa" + Date.now(), chantierId, numero: "FA-" + new Date().getFullYear() + "-" + String(prev.length + 1).padStart(4, "0"),
        type: "solde", montantHT: resteAFacturer, montantTVA: tva, statutPaiement: "emise",
        emiseLe: new Date().toISOString().slice(0, 10), echeanceLe: new Date(Date.now() + 30 * 86400000).toISOString().slice(0, 10) },
      ...prev,
    ]);
  }

  function ajouterClient(client) {
    const id = "cl" + Date.now();
    setClients((prev) => [...prev, { id, ...client }]);
    return id;
  }
  function ajouterFournisseur(fournisseur) {
    setFournisseurs((prev) => [...prev, { id: "f" + Date.now(), ...fournisseur }]);
  }
  function affecterAgent(chantierId, agentId) {
    setChantiers((prev) => prev.map((c) =>
      c.id === chantierId && !c.affectations.includes(agentId)
        ? { ...c, affectations: [...c.affectations, agentId] }
        : c
    ));
  }
  function retirerAgent(chantierId, agentId) {
    setChantiers((prev) => prev.map((c) =>
      c.id === chantierId ? { ...c, affectations: c.affectations.filter((id) => id !== agentId) } : c
    ));
  }

  function saisirHeure({ chantier, chantierId, heures, poste }) {
    setHeuresListe((prev) => [{ id: "h" + Date.now(), chantier, chantierId, ouvrier: utilisateur.nom, heures, poste, statut: "a_valider", commentaire: "", quand: "a l'instant" }, ...prev]);
  }
  function accepterHeure(id) {
    setHeuresListe((prev) => prev.map((h) => (h.id === id ? { ...h, statut: "acceptee" } : h)));
  }
  function refuserHeure(id, commentaire) {
    setHeuresListe((prev) => prev.map((h) => (h.id === id ? { ...h, statut: "refusee", commentaire } : h)));
  }
  function validerHeure(id) {
    setHeuresListe((prev) => prev.map((h) => (h.id === id ? { ...h, statut: "validee" } : h)));
  }

  if (!utilisateur) {
    return <LoginScreen onLogin={(u) => {
      setUtilisateur(u);
      setTab(u.role === "direction" ? "dashboard" : "heures");
    }} />;
  }

  if (!verifie) {
    return <VerificationScreen utilisateur={utilisateur} onRetour={() => setUtilisateur(null)} onVerified={() => {
      setVerifie(true);
      const messagesNonLus = messages.filter((m) => m.versId === utilisateur.id && !m.lu).length;
      setTab(messagesNonLus > 0 ? "messagerie" : (utilisateur.role === "direction" ? "dashboard" : "heures"));
    }} />;
  }

  return (
    <div style={{ fontFamily: "'Inter', 'Segoe UI', ui-sans-serif, system-ui, sans-serif" }} className="min-h-screen flex" >
      <style>{`
        .num { font-variant-numeric: tabular-nums; font-family: 'JetBrains Mono', 'SF Mono', ui-monospace, monospace; }
        @media print {
          .no-print { display: none !important; }
          .print-full { max-width: 100% !important; padding: 0 !important; }
          .print-only { display: block !important; }
          .print-only-inline { display: inline-block !important; }
          body, .min-h-screen { background: #fff !important; }
        }
        .print-only, .print-only-inline { display: none; }
      `}</style>
      <Sidebar tab={tab} setTab={setTab} utilisateur={utilisateur} onLogout={() => { api.logout(); setUtilisateur(null); }} mobileOuvert={menuMobileOuvert} setMobileOuvert={setMenuMobileOuvert} setComptaSousOngletCible={setComptaSousOngletCible} />
      <div className="flex-1 flex flex-col min-h-screen w-full min-w-0" style={{ background: BG }}>
        <MobileTopBar onOuvrirMenu={() => setMenuMobileOuvert(true)} />
        <TopHeader tab={tab} />
        <div className="flex-1 w-full max-w-[1180px] mx-auto px-4 sm:px-8 py-5 sm:py-7 min-w-0">
          {tab === "devis" && (
            <DevisTab
              library={library}
              clients={clients} ajouterClient={ajouterClient}
              devisNom={devisNom} setDevisNom={setDevisNom}
              devisClientId={devisClientId} setDevisClientId={setDevisClientId}
              devisSociete={devisSociete} setDevisSociete={setDevisSociete}
              taux={taux} setTaux={setTaux}
              pct={pct} setPct={setPct}
              remise={remise} setRemise={setRemise}
              tauxDefaut={tauxDefaut} pctDefaut={pctDefaut} remiseDefaut={remiseDefaut}
              lignes={lignes} addLigne={addLigne} updateQte={updateQte} removeLigne={removeLigne}
              sections={sections} sectionActiveId={sectionActiveId} setSectionActiveId={setSectionActiveId}
              ajouterSection={ajouterSection} supprimerSection={supprimerSection}
              montantTotal={montantTotal} montantRemise={montantRemise} montantApresRemise={montantApresRemise}
              montantFourniture={montantFourniture} montantFrais={montantFrais} montantMO={montantMO} heures={heures} jours={jours}
              enregistrerDevis={enregistrerDevis}
            />
          )}
          {tab === "library" && <LibraryTab library={library} setLibrary={setLibrary} />}
          {tab === "repertoires" && (
            <RepertoiresTab clients={clients} setClients={setClients} fournisseurs={fournisseurs} setFournisseurs={setFournisseurs} />
          )}
          {tab === "dashboard" && (
            <DashboardTab
              chantiers={chantiersActifs(utilisateur, chantiers)} clients={clients}
              personnel={utilisateursSysteme.filter((u) => u.role !== "direction")}
              taux={tauxDefaut} pct={pctDefaut} seuilAlerte={seuilAlerte}
              affecterAgent={affecterAgent} retirerAgent={retirerAgent} utilisateur={utilisateur}
              heuresListe={heuresListe} notifications={notifications} setTab={setTab}
              remunerations={remunerations} charges={charges} utilisateursSysteme={utilisateursSysteme}
              tousLesChantiers={chantiers} changerStatutChantier={changerStatutChantier}
            />
          )}
          {tab === "heures" && (
            <HeuresTab
              heuresListe={heuresListe} chantiers={chantiersActifs(utilisateur, chantiers)} utilisateur={utilisateur}
              saisirHeure={saisirHeure} accepterHeure={accepterHeure} refuserHeure={refuserHeure} validerHeure={validerHeure}
            />
          )}
          {tab === "plan" && (
            <PlanTab
              chantiers={chantiersActifs(utilisateur, chantiers)} clients={clients}
              planElements={planElements} changerStatutElement={changerStatutElement} historiqueAvancement={historiqueAvancement}
              utilisateur={utilisateur} personnel={utilisateursSysteme} tauxDefaut={tauxDefaut} pctDefaut={pctDefaut}
              chantierCible={planChantierCible}
              plansPdf={plansPdf} importerPlanPdf={importerPlanPdf} envoyerFicheAffectation={envoyerFicheAffectation}
            />
          )}
          {tab === "messagerie" && (
            <MessagerieTab
              messages={messages} envoyerMessage={envoyerMessage} marquerMessageLu={marquerMessageLu}
              utilisateur={utilisateur} utilisateursSysteme={utilisateursSysteme} chantiers={chantiers}
              correspondants={correspondantsAutorises(utilisateur, utilisateursSysteme, chantiers)}
              ouvrirPartieConcernee={ouvrirPartieConcernee}
            />
          )}
          {tab === "comptabilite" && (
            <ComptabiliteTab
              chantiers={chantiers} clients={clients} factures={factures} paiements={paiements}
              enregistrerPaiement={enregistrerPaiement}
              utilisateursSysteme={utilisateursSysteme} remunerations={remunerations} definirRemuneration={definirRemuneration}
              charges={charges} ajouterCharge={ajouterCharge} supprimerCharge={supprimerCharge}
              heuresListe={heuresListe} pctDefaut={pctDefaut} tauxDefaut={tauxDefaut}
              sousOngletCible={comptaSousOngletCible}
            />
          )}
          {tab === "notifications" && (
            <NotificationsTab notifications={notifications} setNotifications={setNotifications} utilisateur={utilisateur} chantiers={chantiers} ouvrirPartieConcernee={ouvrirPartieConcernee} />
          )}
          {tab === "productivite" && (
            <ProductiviteTab heuresListe={heuresListe} library={library} utilisateur={utilisateur} chantiers={chantiersActifs(utilisateur, chantiers)} tauxDefaut={tauxDefaut} pctDefaut={pctDefaut} />
          )}
          {tab === "parametres" && (
            <ParametresTab
              pctDefaut={pctDefaut} setPctDefaut={setPctDefaut}
              tauxDefaut={tauxDefaut} setTauxDefaut={setTauxDefaut}
              seuilAlerte={seuilAlerte} setSeuilAlerte={setSeuilAlerte}
              coeffConsommables={coeffConsommables} setCoeffConsommables={setCoeffConsommables}
              remiseDefaut={remiseDefaut} setRemiseDefaut={setRemiseDefaut}
              utilisateursSysteme={utilisateursSysteme} setUtilisateursSysteme={setUtilisateursSysteme} utilisateur={utilisateur}
              remunerations={remunerations} definirRemuneration={definirRemuneration}
            />
          )}
        </div>
        <Footer />
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
function Sidebar({ tab, setTab, utilisateur, onLogout, mobileOuvert, setMobileOuvert, setComptaSousOngletCible }) {
  // Structure du menu : items simples + groupes depliables. Chaque enfant de
  // groupe "comptabilite" cible un sous-onglet precis de ComptabiliteTab ;
  // les autres groupes ont des enfants qui sont directement des onglets.
  const NAV_STRUCTURE = [
    { type: "item", id: "dashboard", label: "Tableau de bord", icon: LayoutDashboard, roles: ["direction", "chef_chantier"] },
    { type: "group", id: "commercial", label: "Devis & clients", icon: FileText, roles: ["direction"], children: [
        { id: "devis", label: "Nouveau devis", roles: ["direction"] },
        { id: "library", label: "Bibliotheque de prix", roles: ["direction"] },
        { id: "repertoires", label: "Clients & fournisseurs", roles: ["direction"] },
      ] },
    { type: "group", id: "comptabilite", label: "Comptabilite", icon: Receipt, roles: ["direction"], sousOnglets: true, children: [
        { id: "factures", label: "Factures & paiements", roles: ["direction"] },
        { id: "remuneration", label: "Remuneration du personnel", roles: ["direction"] },
        { id: "charges", label: "Charges de l'entreprise", roles: ["direction"] },
        { id: "analytique", label: "Comptabilite analytique", roles: ["direction"] },
        { id: "bilan", label: "Bilan annuel", roles: ["direction"] },
      ] },
    { type: "group", id: "chantiers", label: "Suivi chantiers", icon: MapPin, roles: ["direction", "chef_chantier", "ouvrier"], children: [
        { id: "plan", label: "Plan & avancement", roles: ["direction", "chef_chantier", "ouvrier"] },
        { id: "heures", label: "Heures a valider", roles: ["direction", "chef_chantier", "ouvrier"] },
        { id: "productivite", label: "Productivite", roles: ["direction", "chef_chantier"] },
      ] },
    { type: "group", id: "communication", label: "Communication", icon: MessageCircle, roles: ["direction", "chef_chantier", "ouvrier"], children: [
        { id: "messagerie", label: "Messagerie interne", roles: ["direction", "chef_chantier", "ouvrier"] },
        { id: "notifications", label: "Notifications", roles: ["direction", "chef_chantier", "ouvrier"] },
      ] },
    { type: "item", id: "parametres", label: "Parametres", icon: Settings, roles: ["direction"] },
  ];

  // Le groupe contenant l'onglet actif reste toujours deplie automatiquement.
  const groupeActif = NAV_STRUCTURE.find((g) => g.type === "group" && (g.sousOnglets ? tab === g.id : g.children.some((c) => c.id === tab)));
  const [groupeOuvert, setGroupeOuvert] = useState(groupeActif ? groupeActif.id : null);
  useEffect(() => {
    if (groupeActif) setGroupeOuvert(groupeActif.id);
  }, [tab]);

  function choisir(id) {
    setTab(id);
    if (setMobileOuvert) setMobileOuvert(false);
  }
  function choisirEnfant(groupe, enfantId) {
    if (groupe.sousOnglets) {
      setTab(groupe.id);
      if (setComptaSousOngletCible) setComptaSousOngletCible(enfantId);
    } else {
      setTab(enfantId);
    }
    if (setMobileOuvert) setMobileOuvert(false);
  }

  const contenu = (fermer) => (
    <>
      <div className="px-5 pt-7 pb-6 flex items-start justify-between" style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
        <div>
          <img src={LOGO_SRC} alt="SLK Clim" className="h-10 w-auto mb-3" style={{ filter: "brightness(0) invert(1)" }} />
          <div className="text-white font-semibold text-[15px] leading-tight tracking-tight">SLK Manager</div>
          <div className="text-[11px] mt-1" style={{ color: STEEL }}>Ventilation &middot; Desenfumage</div>
        </div>
        {fermer && (
          <button onClick={() => setMobileOuvert(false)} className="text-white p-1 shrink-0" aria-label="Fermer le menu">
            <X size={20} />
          </button>
        )}
      </div>
      <nav className="flex-1 px-3 pt-5 overflow-y-auto">
        {NAV_STRUCTURE.filter((it) => it.roles.includes(utilisateur.role)).map((it) => {
          if (it.type === "item") {
            const active = tab === it.id;
            const Icon = it.icon;
            return (
              <button
                key={it.id}
                onClick={() => choisir(it.id)}
                className="w-full flex items-center gap-3 px-3.5 py-2.5 rounded-md text-[13.5px] font-medium mb-1 transition-colors"
                style={{
                  background: active ? "rgba(232,98,44,0.14)" : "transparent",
                  color: active ? "#FFFFFF" : STEEL,
                  borderLeft: active ? "2.5px solid " + ACCENT : "2.5px solid transparent",
                }}
                onMouseEnter={(e) => { if (!active) e.currentTarget.style.background = NAVY_SOFT; }}
                onMouseLeave={(e) => { if (!active) e.currentTarget.style.background = "transparent"; }}
              >
                <Icon size={16} strokeWidth={2} />
                {it.label}
                {active && <ChevronRight size={14} className="ml-auto opacity-60" />}
              </button>
            );
          }
          // Groupe avec sous-menu
          const enfants = it.children.filter((c) => c.roles.includes(utilisateur.role));
          if (!enfants.length) return null;
          const Icon = it.icon;
          const ouvert = groupeOuvert === it.id;
          const groupeActifIci = enfants.some((c) => (it.sousOnglets ? tab === it.id : tab === c.id));
          return (
            <div key={it.id} className="mb-1">
              <button
                onClick={() => setGroupeOuvert(ouvert ? null : it.id)}
                className="w-full flex items-center gap-3 px-3.5 py-2.5 rounded-md text-[13.5px] font-medium transition-colors"
                style={{ background: groupeActifIci ? "rgba(232,98,44,0.10)" : "transparent", color: groupeActifIci ? "#FFFFFF" : STEEL }}
                onMouseEnter={(e) => { if (!groupeActifIci) e.currentTarget.style.background = NAVY_SOFT; }}
                onMouseLeave={(e) => { e.currentTarget.style.background = groupeActifIci ? "rgba(232,98,44,0.10)" : "transparent"; }}
              >
                <Icon size={16} strokeWidth={2} color={groupeActifIci ? ACCENT : undefined} />
                {it.label}
                <ChevronRight size={14} className="ml-auto transition-transform" style={{ transform: ouvert ? "rotate(90deg)" : "rotate(0deg)", opacity: 0.6 }} />
              </button>
              {ouvert && (
                <div className="ml-4 mt-0.5 mb-1 pl-3" style={{ borderLeft: "1px solid " + hexAlpha(SUBMENU_ACCENT, 0.35) }}>
                  {enfants.map((c) => {
                    const active = it.sousOnglets ? (tab === it.id) : (tab === c.id);
                    return (
                      <button
                        key={c.id}
                        onClick={() => choisirEnfant(it, c.id)}
                        className="w-full text-left px-3 py-2 rounded-md text-[12.5px] font-medium mb-0.5 transition-colors"
                        style={{
                          background: active ? hexAlpha(SUBMENU_ACCENT, 0.18) : "transparent",
                          color: active ? SUBMENU_ACCENT : STEEL,
                          borderLeft: active ? "2px solid " + SUBMENU_ACCENT : "2px solid transparent",
                        }}
                        onMouseEnter={(e) => { if (!active) e.currentTarget.style.background = NAVY_SOFT; }}
                        onMouseLeave={(e) => { if (!active) e.currentTarget.style.background = "transparent"; }}
                      >
                        {c.label}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </nav>
      <div className="px-5 py-4 shrink-0" style={{ borderTop: "1px solid rgba(255,255,255,0.08)" }}>
        <div className="flex items-center gap-2.5 mb-3">
          <div className="w-8 h-8 rounded-full flex items-center justify-center text-[12px] font-semibold text-white shrink-0" style={{ background: ACCENT }}>{utilisateur.initiales}</div>
          <div className="leading-tight">
            <div className="text-[12.5px] font-medium text-white">{utilisateur.nom}</div>
            <div className="text-[10.5px]" style={{ color: STEEL }}>{utilisateur.roleLabel}</div>
          </div>
        </div>
        <button onClick={onLogout} className="w-full text-[11.5px] font-medium text-left px-2.5 py-1.5 rounded-md transition-colors" style={{ color: STEEL }}
          onMouseEnter={(e) => { e.currentTarget.style.background = NAVY_SOFT; e.currentTarget.style.color = "#fff"; }}
          onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; e.currentTarget.style.color = STEEL; }}
        >
          Se deconnecter
        </button>
      </div>
    </>
  );

  return (
    <>
      {/* Desktop : barre laterale persistante */}
      <div className="hidden lg:flex w-[248px] shrink-0 flex-col no-print" style={{ background: NAVY }}>
        {contenu(false)}
      </div>

      {/* Mobile/tablette : tiroir plein ecran, ouvert via le bouton menu */}
      {mobileOuvert && (
        <div className="lg:hidden fixed inset-0 z-50 flex no-print">
          <div className="absolute inset-0" style={{ background: "rgba(15,27,46,0.5)" }} onClick={() => setMobileOuvert(false)} />
          <div className="relative w-[82%] max-w-[300px] flex flex-col" style={{ background: NAVY }}>
            {contenu(true)}
          </div>
        </div>
      )}
    </>
  );
}

function MobileTopBar({ onOuvrirMenu }) {
  return (
    <div className="lg:hidden flex items-center gap-3 px-4 py-3 no-print" style={{ background: NAVY }}>
      <button onClick={onOuvrirMenu} className="text-white p-1" aria-label="Ouvrir le menu">
        <Menu size={22} />
      </button>
      <img src={LOGO_SRC} alt="SLK Clim" className="h-6 w-auto" style={{ filter: "brightness(0) invert(1)" }} />
      <span className="text-white text-[13.5px] font-semibold">SLK Manager</span>
    </div>
  );
}

// ---------------------------------------------------------------------------
function genererOTP() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function VerificationScreen({ utilisateur, onVerified, onRetour }) {
  const [otpEnvoye, setOtpEnvoye] = useState(genererOTP());
  const [otpSaisi, setOtpSaisi] = useState("");
  const [erreur, setErreur] = useState("");

  function verifierOTP() {
    if (otpSaisi.trim() !== otpEnvoye) {
      setErreur("Code de verification incorrect.");
      return;
    }
    onVerified();
  }

  function renvoyerCode() {
    setOtpEnvoye(genererOTP());
    setOtpSaisi("");
    setErreur("");
  }

  return (
    <div className="min-h-screen flex items-center justify-center p-8" style={{ background: NAVY }}>
      <div className="w-full max-w-[420px]">
        <div className="flex flex-col items-center mb-7">
          <img src={LOGO_SRC} alt="SLK Clim" className="h-11 w-auto mb-4" style={{ filter: "brightness(0) invert(1)" }} />
          <div className="text-white font-semibold text-[17px] tracking-tight">SLK Manager</div>
          <div className="text-[12px] mt-1" style={{ color: STEEL }}>Verification d'acces</div>
        </div>

        <Card className="p-6">
          <button onClick={onRetour} className="text-[11.5px] font-medium mb-4" style={{ color: ACCENT_DEEP }}>&larr; Ce n'est pas moi</button>
          <h2 className="text-[15px] font-semibold mb-1" style={{ color: INK }}>Code de verification</h2>
          <p className="text-[12.5px] mb-4" style={{ color: MUTE }}>
            Un code a ete envoye a <strong>{utilisateur.email || "votre email personnel (non renseigne — a ajouter dans Parametres)"}</strong>, l'email personnel de {utilisateur.nom}. Saisissez-le ci-dessous pour continuer.
          </p>
          <Field label="Code recu par email">
            <input
              value={otpSaisi}
              onChange={(e) => { setOtpSaisi(e.target.value); setErreur(""); }}
              onKeyDown={(e) => { if (e.key === "Enter") verifierOTP(); }}
              placeholder="6 chiffres"
              maxLength={6}
              autoFocus
              className="num w-full rounded-md px-3 py-2.5 text-[15px] tracking-[0.3em] text-center"
              style={{ border: "1px solid " + BORDER }}
            />
          </Field>
          {erreur && <div className="text-[11.5px] mt-2" style={{ color: BAD }}>{erreur}</div>}
          <button onClick={verifierOTP} className="w-full mt-4 text-white text-[13px] font-semibold py-2.5 rounded-md" style={{ background: ACCENT }}>
            Valider et continuer
          </button>
          <div className="flex items-center justify-end mt-3">
            <button onClick={renvoyerCode} className="text-[11.5px] font-medium" style={{ color: ACCENT_DEEP }}>Renvoyer le code</button>
          </div>
          <div className="text-[10.5px] mt-3 text-center italic" style={{ color: MUTE }}>Code de demonstration (email simule, aucun envoi reel) : {otpEnvoye}</div>
        </Card>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
function LoginScreen({ onLogin }) {
  const [email, setEmail] = useState("");
  const [motDePasse, setMotDePasse] = useState("");
  const [erreur, setErreur] = useState("");
  const [enCours, setEnCours] = useState(false);

  async function handleConnexion() {
    if (!email || !motDePasse) { setErreur("Saisissez votre email et votre mot de passe."); return; }
    setErreur("");
    setEnCours(true);
    try {
      const data = await api.login(email.trim(), motDePasse);
      // L'API renvoie { token, user: { id, nom, prenom, role } }. On complete
      // les champs d'affichage attendus par l'interface (initiales, roleLabel).
      const u = data.user;
      const roleLabel = { direction: "Direction", chef_chantier: "Chef de chantier", ouvrier: "Ouvrier" }[u.role] || u.role;
      const initiales = ((u.prenom ? u.prenom[0] : "") + (u.nom ? u.nom[0] : "")).toUpperCase() || "?";
      onLogin({ ...u, roleLabel, initiales, poste: roleLabel, affectations: [] });
    } catch (e) {
      setErreur(e.message || "Connexion impossible.");
    } finally {
      setEnCours(false);
    }
  }

  return (
    <div className="min-h-screen flex" style={{ fontFamily: "'Inter', 'Segoe UI', ui-sans-serif, system-ui, sans-serif" }}>
      <div className="hidden lg:flex flex-col justify-between w-[380px] shrink-0 p-10" style={{ background: NAVY }}>
        <div>
          <img src={LOGO_SRC} alt="SLK Clim" className="h-12 w-auto mb-6" style={{ filter: "brightness(0) invert(1)" }} />
          <div className="text-white font-semibold text-[22px] leading-tight tracking-tight">SLK Manager</div>
          <div className="text-[13px] mt-2" style={{ color: STEEL }}>Logiciel metier &mdash; ventilation, desenfumage et reseaux aerauliques</div>
        </div>
        <div className="text-[11.5px]" style={{ color: STEEL }}>SAS SLK CLIM &mdash; 8 avenue Roland Moreno, 95740 Frepillon</div>
      </div>

      <div className="flex-1 flex items-center justify-center p-8" style={{ background: BG }}>
        <div className="w-full max-w-[440px]">
          <img src={LOGO_SRC} alt="SLK Clim" className="h-10 w-auto mb-6 lg:hidden" />
          <h1 className="text-[20px] font-semibold mb-1" style={{ color: INK }}>Connexion</h1>
          <p className="text-[13px] mb-6" style={{ color: MUTE }}>Saisissez vos identifiants pour acceder au logiciel.</p>

          <Card className="p-6">
            <Field label="Email">
              <input
                type="email"
                value={email}
                onChange={(e) => { setEmail(e.target.value); setErreur(""); }}
                onKeyDown={(e) => { if (e.key === "Enter") handleConnexion(); }}
                placeholder="vous@slkclim.fr"
                autoFocus
                className="w-full rounded-md px-3 py-2.5 text-[13.5px]" style={{ border: "1px solid " + BORDER }}
              />
            </Field>
            <div className="h-3" />
            <Field label="Mot de passe">
              <input
                type="password"
                value={motDePasse}
                onChange={(e) => { setMotDePasse(e.target.value); setErreur(""); }}
                onKeyDown={(e) => { if (e.key === "Enter") handleConnexion(); }}
                placeholder="********"
                className="w-full rounded-md px-3 py-2.5 text-[13.5px]" style={{ border: "1px solid " + BORDER }}
              />
            </Field>
            {erreur && <div className="text-[11.5px] mt-3" style={{ color: BAD }}>{erreur}</div>}
            <button onClick={handleConnexion} disabled={enCours}
              className="w-full mt-4 text-white text-[13px] font-semibold py-2.5 rounded-md disabled:opacity-50" style={{ background: ACCENT }}>
              {enCours ? "Connexion en cours..." : "Se connecter"}
            </button>
          </Card>

          <div className="text-[11px] mt-6 text-center" style={{ color: MUTE }}>
            Compte cree par la Direction. Mot de passe oublie ? Contactez votre responsable.
          </div>
        </div>
      </div>
    </div>
  );
}

function TopHeader({ tab }) {
  const titles = {
    repertoires: ["Clients & fournisseurs", "A enregistrer avant de pouvoir creer un devis"],
    devis: ["Nouveau devis", "Chiffrage automatique a partir de la bibliotheque de prix"],
    library: ["Bibliotheque de prix", "Base d'articles modifiable, sans valeur codee en dur"],
    dashboard: ["Tableau de bord", "Suivi budgetaire et d'avancement multi-chantiers"],
    heures: ["Heures et circuit de validation", "A valider (ouvrier) -> Acceptee (chef de chantier) -> Validee (Direction)"],
    plan: ["Plan & avancement", "Suivi visuel des elements par chantier, photos a l'appui"],
    messagerie: ["Messagerie interne", "Communication hierarchisee, tracee pour chaque chantier"],
    productivite: ["Productivite", "Heures reelles par ouvrier et temps de pose de reference"],
    comptabilite: ["Comptabilite", "Paiements, remuneration, charges et comptabilite analytique"],
    notifications: ["Notifications", "Alertes et evenements concernant votre role"],
    parametres: ["Parametres", "Valeurs par defaut du logiciel — jamais codees en dur"],
  };
  const [title, sub] = titles[tab];
  return (
    <div className="px-4 sm:px-8 py-4 sm:py-5 flex items-center justify-between no-print" style={{ background: SURFACE, borderBottom: "1px solid " + BORDER }}>
      <div className="min-w-0">
        <h1 className="text-[16px] sm:text-[19px] font-semibold tracking-tight truncate" style={{ color: INK }}>{title}</h1>
        <p className="text-[11.5px] sm:text-[12.5px] mt-0.5 hidden sm:block" style={{ color: MUTE }}>{sub}</p>
      </div>
      <div className="hidden sm:flex items-center gap-2 px-3 py-1.5 rounded-md text-[12px] shrink-0" style={{ background: BG, color: MUTE }}>
        <Building2 size={13} /> SAS SLK CLIM
      </div>
    </div>
  );
}

function Card({ children, className = "", style = {} }) {
  return (
    <div
      className={"rounded-lg " + className}
      style={{ background: SURFACE, border: "1px solid " + BORDER, boxShadow: "0 1px 2px rgba(15,27,46,0.04)", ...style }}
    >
      {children}
    </div>
  );
}

// ---------------------------------------------------------------------------
function DevisTab(props) {
  const {
    library, clients, ajouterClient, devisNom, setDevisNom, devisClientId, setDevisClientId, devisSociete, setDevisSociete, taux, setTaux, pct, setPct,
    remise, setRemise, tauxDefaut, pctDefaut, remiseDefaut,
    lignes, addLigne, updateQte, removeLigne, sections, sectionActiveId, setSectionActiveId, ajouterSection, supprimerSection,
    montantTotal, montantRemise, montantApresRemise,
    montantFourniture, montantFrais, montantMO, heures, jours, enregistrerDevis,
  } = props;
  const [nouveauClient, setNouveauClient] = useState(false);
  const [ncNom, setNcNom] = useState("");
  const [ncAdresse, setNcAdresse] = useState("");

  function handleAjouterClient() {
    if (!ncNom) return;
    const id = ajouterClient({ nom: ncNom, adresse: ncAdresse, telephone: "", contact: "" });
    setDevisClientId(id);
    setNcNom(""); setNcAdresse(""); setNouveauClient(false);
  }

  const [cat, setCat] = useState("Tous");
  const categories = ["Tous", ...Array.from(new Set(library.map((l) => l.categorie)))];
  const filtered = cat === "Tous" ? library : library.filter((l) => l.categorie === cat);

  // Sections (module 12) — au lieu de retaper Niveau/Zone/Plan a chaque ligne
  // comme sur un export EBP, on cree la section une fois, puis on clique
  // simplement sur celle qui est active avant d'ajouter des postes.
  const [formOuvert, setFormOuvert] = useState(sections.length === 0);
  const [niveau, setNiveau] = useState("NIV VS");
  const [zone, setZone] = useState("ZONE 1");
  const [planNumero, setPlanNumero] = useState("100");
  const [planIndice, setPlanIndice] = useState("IND D");
  const [categorieSection, setCategorieSection] = useState("SOUFFLAGE");

  function sectionLabel(s) {
    if (!s) return "Hors structure";
    return s.niveau + (s.zone ? " - " + s.zone : "") + " - PLAN N°" + s.planNumero + (s.planIndice ? " - " + s.planIndice : "") + " / " + s.categorie;
  }
  function creerSection() {
    if (!niveau || !planNumero) return;
    ajouterSection({ niveau, zone, planNumero, planIndice, categorie: categorieSection });
    setFormOuvert(false);
  }

  const sectionActive = sections.find((s) => s.id === sectionActiveId) || null;

  const groupes = [];
  for (const l of lignes) {
    const label = sectionLabel(sections.find((s) => s.id === l.sectionId));
    let g = groupes.find((g) => g.label === label);
    if (!g) { g = { label, lignes: [] }; groupes.push(g); }
    g.lignes.push(l);
  }

  return (
    <div className="grid grid-cols-1 lg:grid-cols-[1fr_380px] gap-5">
      <div className="space-y-5">
        <Card className="p-6">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <Field label="Nom du chantier">
              <input value={devisNom} onChange={(e) => setDevisNom(e.target.value)} placeholder="Ex. Chantier Bureaux Nanterre"
                className="w-full rounded-md px-3 py-2 text-[13.5px] focus:outline-none focus:ring-2" style={{ border: "1px solid " + BORDER, ["--tw-ring-color"]: ACCENT }} />
            </Field>
            <Field label="Client">
              <select value={devisClientId} onChange={(e) => setDevisClientId(e.target.value)}
                className="w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }}>
                <option value="">Selectionner un client...</option>
                {clients.map((c) => <option key={c.id} value={c.id}>{c.nom}</option>)}
              </select>
              {!nouveauClient ? (
                <button onClick={() => setNouveauClient(true)} className="mt-1.5 text-[11.5px] font-medium flex items-center gap-1" style={{ color: ACCENT_DEEP }}>
                  <UserPlus size={12} /> Nouveau client
                </button>
              ) : (
                <div className="mt-2 p-3 rounded-md space-y-2" style={{ background: BG }}>
                  <input value={ncNom} onChange={(e) => setNcNom(e.target.value)} placeholder="Nom du client"
                    className="w-full rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }} />
                  <input value={ncAdresse} onChange={(e) => setNcAdresse(e.target.value)} placeholder="Adresse (optionnel)"
                    className="w-full rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }} />
                  <div className="flex gap-2">
                    <button onClick={handleAjouterClient} className="text-[11.5px] font-semibold text-white px-3 py-1.5 rounded-md" style={{ background: NAVY }}>Enregistrer</button>
                    <button onClick={() => setNouveauClient(false)} className="text-[11.5px]" style={{ color: MUTE }}>Annuler</button>
                  </div>
                </div>
              )}
            </Field>
            <Field label="Societe donneuse d'ordre (texte libre)">
              <input value={devisSociete} onChange={(e) => setDevisSociete(e.target.value)} placeholder="Ex. AXIMA Concept Equans (si different du client facture)"
                className="w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
              <p className="text-[10.5px] mt-1" style={{ color: MUTE }}>Utile quand le chantier passe par un maitre d'oeuvre ou une entreprise generale differente du client facture.</p>
            </Field>
          </div>
        </Card>

        <Card className="p-6">
          <div className="flex items-center justify-between mb-3">
            <div>
              <h3 className="text-[13px] font-semibold" style={{ color: INK }}>Sections du devis</h3>
              <p className="text-[11.5px] mt-0.5" style={{ color: MUTE }}>Creez chaque section une fois, puis cliquez pour choisir celle qui est active.</p>
            </div>
            <button onClick={() => setFormOuvert((v) => !v)} className="flex items-center gap-1 text-[11.5px] font-semibold px-2.5 py-1.5 rounded-md" style={{ color: ACCENT_DEEP, border: "1px dashed " + ACCENT }}>
              <Plus size={13} /> Section
            </button>
          </div>

          {sections.length > 0 && (
            <div className="flex flex-wrap gap-2 mb-3">
              {sections.map((s) => (
                <button key={s.id} onClick={() => setSectionActiveId(s.id)}
                  className="flex items-center gap-1.5 text-[11.5px] px-2.5 py-1.5 rounded-full font-medium"
                  style={sectionActiveId === s.id ? { background: NAVY, color: "#fff" } : { border: "1px solid " + BORDER, color: MUTE }}>
                  {sectionLabel(s)}
                  <XCircle size={12} className="opacity-60 hover:opacity-100" onClick={(e) => { e.stopPropagation(); supprimerSection(s.id); }} />
                </button>
              ))}
            </div>
          )}

          {formOuvert && (
            <div className="p-3 rounded-md space-y-2" style={{ background: BG }}>
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                <input value={niveau} onChange={(e) => setNiveau(e.target.value)} placeholder="Niveau" className="rounded-md px-2 py-1.5 text-[12px]" style={{ border: "1px solid " + BORDER }} />
                <input value={zone} onChange={(e) => setZone(e.target.value)} placeholder="Zone" className="rounded-md px-2 py-1.5 text-[12px]" style={{ border: "1px solid " + BORDER }} />
                <input value={planNumero} onChange={(e) => setPlanNumero(e.target.value)} placeholder="N Plan" className="rounded-md px-2 py-1.5 text-[12px]" style={{ border: "1px solid " + BORDER }} />
                <input value={planIndice} onChange={(e) => setPlanIndice(e.target.value)} placeholder="Indice" className="rounded-md px-2 py-1.5 text-[12px]" style={{ border: "1px solid " + BORDER }} />
                <select value={categorieSection} onChange={(e) => setCategorieSection(e.target.value)} className="rounded-md px-2 py-1.5 text-[12px]" style={{ border: "1px solid " + BORDER }}>
                  {["SOUFFLAGE", "REPRISE", "AIR NEUF", "REJET", "VEX"].map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <button onClick={creerSection} className="text-[11.5px] font-semibold text-white px-3 py-1.5 rounded-md" style={{ background: NAVY }}>Creer cette section</button>
            </div>
          )}

          <div className="text-[11px] mt-3 flex items-center gap-1.5" style={{ color: sectionActive ? GOOD : MUTE }}>
            {sectionActive ? <CheckCircle2 size={13} /> : <AlertTriangle size={13} />}
            Section active pour les prochains ajouts : <strong>{sectionLabel(sectionActive)}</strong>
          </div>
        </Card>

        <Card className="p-6">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-[13px] font-semibold" style={{ color: INK }}>Ajouter des postes depuis la bibliotheque</h3>
          </div>
          <div className="flex gap-2 flex-wrap mb-4">
            {categories.map((c) => (
              <button key={c} onClick={() => setCat(c)}
                className="text-[11.5px] px-3 py-1.5 rounded-full font-medium transition-colors"
                style={cat === c ? { background: NAVY, color: "#fff" } : { border: "1px solid " + BORDER, color: MUTE }}>
                {c}
              </button>
            ))}
          </div>
          <div className="max-h-64 overflow-y-auto divide-y" style={{ borderColor: BORDER }}>
            {filtered.map((item) => (
              <div key={item.id} className="flex items-center justify-between py-2.5" style={{ borderColor: BORDER }}>
                <div className="flex items-center gap-3">
                  <span className="num text-[10.5px] px-1.5 py-0.5 rounded" style={{ background: BG, color: MUTE }}>{item.code}</span>
                  <div>
                    <div className="text-[13px]" style={{ color: INK }}>{item.designation}</div>
                    <div className="num text-[11px]" style={{ color: MUTE }}>{item.prix.toFixed(2)} EUR / {item.unite}</div>
                  </div>
                </div>
                <button onClick={() => addLigne(item)} className="flex items-center gap-1 text-[11.5px] font-semibold px-2.5 py-1.5 rounded-md transition-colors" style={{ color: ACCENT_DEEP }}>
                  <Plus size={13} /> Ajouter
                </button>
              </div>
            ))}
          </div>
        </Card>

        <Card className="p-6">
          <h3 className="text-[13px] font-semibold mb-4" style={{ color: INK }}>Lignes du devis</h3>
          {lignes.length === 0 ? (
            <p className="text-[13px]" style={{ color: MUTE }}>Aucune ligne pour l'instant. Ajoutez des postes depuis la bibliotheque ci-dessus.</p>
          ) : (
            <div className="space-y-5">
              {groupes.map((g) => {
                const sousTotal = g.lignes.reduce((s, l) => s + l.quantite * l.prix, 0);
                return (
                  <div key={g.label}>
                    <div className="text-[11px] font-semibold uppercase tracking-wide mb-2" style={{ color: ACCENT_DEEP }}>{g.label}</div>
                    <div className="space-y-2">
                      {g.lignes.map((l) => (
                        <div key={l.id} className="flex items-center gap-3 text-[13px] rounded-md px-3.5 py-2.5" style={{ border: "1px solid " + BORDER }}>
                          <span className="num text-[10.5px] px-1.5 py-0.5 rounded" style={{ background: BG, color: MUTE }}>{l.code}</span>
                          <div className="flex-1">
                            <div style={{ color: INK }}>{l.designation}</div>
                            <div className="num text-[11px]" style={{ color: MUTE }}>{l.prix.toFixed(2)} EUR / {l.unite}</div>
                          </div>
                          <input type="number" min="0" value={l.quantite} onChange={(e) => updateQte(l.id, parseFloat(e.target.value) || 0)}
                            className="num w-20 rounded-md px-2 py-1.5 text-[13px] text-right" style={{ border: "1px solid " + BORDER }} />
                          <div className="num w-24 text-right font-semibold" style={{ color: INK }}>{fmtEUR(l.quantite * l.prix)}</div>
                          <button onClick={() => removeLigne(l.id)} style={{ color: BAD }}>
                            <Trash2 size={15} />
                          </button>
                        </div>
                      ))}
                    </div>
                    <div className="text-[11.5px] text-right mt-1.5 font-semibold" style={{ color: MUTE }}>Sous-total : {fmtEUR(sousTotal)}</div>
                  </div>
                );
              })}
            </div>
          )}
        </Card>
      </div>

      <div className="space-y-5">
        <Card className="p-6">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-[13px] font-semibold" style={{ color: INK }}>Parametres de calcul</h3>
            <button
              onClick={() => { setTaux(tauxDefaut); setPct(pctDefaut); setRemise(remiseDefaut); }}
              className="text-[11px] font-medium"
              style={{ color: ACCENT_DEEP }}
              title="Reprendre les valeurs definies dans Parametres"
            >
              Reinitialiser
            </button>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Taux horaire moyen (EUR/h)">
              <input type="number" value={taux} onChange={(e) => setTaux(parseFloat(e.target.value) || 0)}
                className="num w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
            </Field>
            <Field label="Remise globale (%)">
              <input type="number" step="0.1" value={remise} onChange={(e) => setRemise(parseFloat(e.target.value) || 0)}
                className="num w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
            </Field>
          </div>
          <div className="grid grid-cols-3 gap-2 mt-3">
            {["fourniture", "frais", "mainOeuvre"].map((k) => (
              <Field key={k} label={k === "fourniture" ? "Fourn. %" : k === "frais" ? "Frais %" : "MO %"}>
                <input type="number" value={pct[k]} onChange={(e) => setPct({ ...pct, [k]: parseFloat(e.target.value) || 0 })}
                  className="num w-full rounded-md px-2 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
              </Field>
            ))}
          </div>
          {pct.fourniture + pct.frais + pct.mainOeuvre !== 100 && (
            <div className="text-[11px] mt-2 flex items-center gap-1" style={{ color: BAD }}>
              <AlertTriangle size={12} /> La somme des pourcentages doit faire 100%.
            </div>
          )}
        </Card>

        <Card style={{ background: NAVY, border: "none" }} className="p-6">
          <h3 className="text-[13px] font-semibold mb-4" style={{ color: "rgba(255,255,255,0.7)" }}>Chiffrage automatique</h3>
          <div className="space-y-3">
            <Row label="Montant total HT" value={fmtEUR(montantTotal)} big />
            {remise > 0 && (
              <>
                <Row label={"Remise globale (" + remise + "%)"} value={"-" + fmtEUR(montantRemise)} />
                <Row label="Total HT apres remise" value={fmtEUR(montantApresRemise)} />
              </>
            )}
            <div className="h-px my-1" style={{ background: "rgba(255,255,255,0.1)" }} />
            <Row label={"Fournitures (" + pct.fourniture + "%)"} value={fmtEUR(montantFourniture)} />
            <Row label={"Frais d'entreprise (" + pct.frais + "%)"} value={fmtEUR(montantFrais)} />
            <Row label={"Main-d'oeuvre (" + pct.mainOeuvre + "%)"} value={fmtEUR(montantMO)} />
            <div className="h-px my-1" style={{ background: "rgba(255,255,255,0.1)" }} />
            <Row label="Heures prevues" value={fmtH(heures)} icon={Clock} accent />
            <Row label="Soit en jours (1j = 16h)" value={fmtJours(jours)} />
          </div>
          <button onClick={enregistrerDevis} disabled={!devisNom || !devisClientId || montantTotal <= 0}
            className="w-full mt-5 text-white text-[13px] font-semibold py-2.5 rounded-md transition-colors disabled:opacity-30"
            style={{ background: ACCENT }}>
            Enregistrer et creer le chantier
          </button>
        </Card>
      </div>
    </div>
  );
}

function Row({ label, value, big, icon: Icon, accent }) {
  return (
    <div className="flex items-center justify-between">
      <span className="text-[12.5px]" style={{ color: big ? "rgba(255,255,255,0.85)" : "rgba(255,255,255,0.55)" }}>{label}</span>
      <span className={"num flex items-center gap-1.5 " + (big ? "text-[19px] font-semibold" : "text-[13px] font-medium")} style={{ color: accent ? "#F0A385" : "#fff" }}>
        {Icon && <Icon size={13} />}
        {value}
      </span>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="block text-[11.5px] font-medium mb-1.5" style={{ color: MUTE }}>{label}</span>
      {children}
    </label>
  );
}

// ---------------------------------------------------------------------------
function LibraryTab({ library, setLibrary }) {
  function updateField(id, field, value) {
    setLibrary((prev) => prev.map((l) => (l.id === id ? { ...l, [field]: value } : l)));
  }
  function addRow() {
    setLibrary((prev) => [...prev, { id: "new-" + Date.now(), code: "--", categorie: "Pose", designation: "Nouvel article", unite: "UTE", prix: 0, tempsPose: null }]);
  }
  function removeRow(id) {
    setLibrary((prev) => prev.filter((l) => l.id !== id));
  }

  return (
    <Card className="overflow-hidden">
      <div className="flex items-center justify-between px-6 py-5" style={{ borderBottom: "1px solid " + BORDER }}>
        <div>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Bibliotheque de prix unitaires</h3>
          <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Base reelle (devis SIG) &mdash; modifiable, a completer avec le temps de pose.</p>
        </div>
        <button onClick={addRow} className="flex items-center gap-1.5 text-[12px] font-semibold text-white px-3.5 py-2 rounded-md" style={{ background: NAVY }}>
          <Plus size={14} /> Ajouter un article
        </button>
      </div>
      <div className="overflow-x-auto">
      <table className="w-full text-[13px]">
        <thead>
          <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
            <th className="px-6 py-2.5 font-semibold">Code</th>
            <th className="px-3 py-2.5 font-semibold">Categorie</th>
            <th className="px-3 py-2.5 font-semibold">Designation</th>
            <th className="px-3 py-2.5 font-semibold">Unite</th>
            <th className="px-3 py-2.5 font-semibold text-right">Prix HT</th>
            <th className="px-3 py-2.5 font-semibold text-right">Temps de pose</th>
            <th className="px-3 py-2.5"></th>
          </tr>
        </thead>
        <tbody>
          {library.map((item, i) => (
            <tr key={item.id} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
              <td className="px-6 py-2">
                <input value={item.code} onChange={(e) => updateField(item.id, "code", e.target.value)}
                  className="num w-14 bg-transparent text-[12px] focus:outline-none" style={{ color: MUTE }} />
              </td>
              <td className="px-3 py-2">
                <input value={item.categorie} onChange={(e) => updateField(item.id, "categorie", e.target.value)}
                  className="w-full bg-transparent text-[13px] focus:outline-none" />
              </td>
              <td className="px-3 py-2">
                <input value={item.designation} onChange={(e) => updateField(item.id, "designation", e.target.value)}
                  className="w-full bg-transparent text-[13px] focus:outline-none" />
              </td>
              <td className="px-3 py-2">
                <input value={item.unite} onChange={(e) => updateField(item.id, "unite", e.target.value)}
                  className="w-16 bg-transparent text-[13px] focus:outline-none" />
              </td>
              <td className="px-3 py-2 text-right">
                <input type="number" value={item.prix} onChange={(e) => updateField(item.id, "prix", parseFloat(e.target.value) || 0)}
                  className="num w-20 bg-transparent text-[13px] text-right focus:outline-none" />
              </td>
              <td className="px-3 py-2 text-right">
                {item.tempsPose === null ? (
                  <span className="text-[10.5px] italic" style={{ color: BAD }}>a definir</span>
                ) : (
                  <input type="number" value={item.tempsPose} onChange={(e) => updateField(item.id, "tempsPose", parseFloat(e.target.value) || 0)}
                    className="num w-16 bg-transparent text-[13px] text-right focus:outline-none" />
                )}
              </td>
              <td className="px-3 py-2 text-right">
                <button onClick={() => removeRow(item.id)} style={{ color: BAD }}>
                  <Trash2 size={14} />
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      </div>
    </Card>
  );
}

// ---------------------------------------------------------------------------
function DashboardTab({ chantiers, clients, personnel, taux, pct, seuilAlerte, affecterAgent, retirerAgent, utilisateur, heuresListe, notifications, setTab, remunerations, charges, utilisateursSysteme, tousLesChantiers, changerStatutChantier }) {
  const [ajoutId, setAjoutId] = useState(null);
  const [choixAgent, setChoixAgent] = useState("");

  const clientNom = (id) => (clients.find((c) => c.id === id) || {}).nom || "-";
  const agent = (id) => personnel.find((p) => p.id === id);

  const totalChargeMensuelle = (charges || []).reduce((s, c) => s + c.montantMensuel, 0);
  const nbChantiersTotal = (tousLesChantiers || chantiers).length;
  const chargeParChantier = nbChantiersTotal > 0 ? totalChargeMensuelle / nbChantiersTotal : 0;

  const rows = chantiers.map((c) => {
    const hPrevues = heuresPrevues(c.montantHT, pct, taux);
    const avancement = hPrevues > 0 ? (c.heuresConsommees / hPrevues) * 100 : 0;
    const marge = c.montantHT - (c.montantHT * pct.fourniture) / 100 - c.heuresConsommees * taux - (c.montantHT * pct.frais) / 100;
    const enAlerte = avancement >= seuilAlerte || c.statut === "en_difficulte";
    // Cout de revient reel (module 20) : fournitures + main-d'oeuvre reellement
    // validee (chargee) + quote-part de charges fixes — pas juste le previsionnel.
    const coutMOReel = remunerations && utilisateursSysteme ? coutReelChantier(c.id, heuresListe || [], remunerations, utilisateursSysteme, taux) : c.heuresConsommees * taux;
    const montantFourniture = (c.montantHT * pct.fourniture) / 100;
    const coutRevientReel = montantFourniture + coutMOReel + chargeParChantier;
    const beneficeReel = c.montantHT - coutRevientReel;
    return { ...c, hPrevues, avancement, marge, enAlerte, coutRevientReel, beneficeReel };
  });

  const totalCA = rows.reduce((s, r) => s + r.montantHT, 0);
  const totalHPrevues = rows.reduce((s, r) => s + r.hPrevues, 0);
  const totalHConsommees = rows.reduce((s, r) => s + r.heuresConsommees, 0);
  const totalCoutRevient = rows.reduce((s, r) => s + r.coutRevientReel, 0);
  const totalBenefice = rows.reduce((s, r) => s + r.beneficeReel, 0);
  const sansChef = rows.filter((r) => !r.affectations.some((id) => agent(id) && agent(id).role === "chef_chantier")).length;
  const chantiersEnAlerte = rows.filter((r) => r.enAlerte);

  // Actions prioritaires — ce qui merite l'attention immediate d'un gerant,
  // pour transformer le tableau de bord en veritable outil de decision et
  // pas seulement en tableau de chiffres.
  const chantierIds = chantiers.map((c) => c.id);
  const heuresAValider = (heuresListe || []).filter((h) => h.statut === "a_valider" && chantierIds.includes(h.chantierId)).length;
  const notifsNonLues = (notifications || []).filter((n) => !n.lu && (utilisateur.role === "direction" || n.pourNom === utilisateur.nom)).length;

  // Export Excel de la liste des devis (module 1) — chaque chantier reflete
  // ici un devis accepte. Fichier .xlsx genere directement dans le navigateur.
  function exporterDevisExcel() {
    const donnees = rows.map((r) => ({
      "Chantier": r.nom,
      "Client": clientNom(r.clientId),
      "Statut devis": r.devisStatut === "accepte" ? "Accepte" : r.devisStatut,
      "Statut chantier": r.statut,
      "Montant devis HT (EUR)": Math.round(r.montantHT * 100) / 100,
      "Heures prevues": Math.round(r.hPrevues * 10) / 10,
      "Heures consommees": r.heuresConsommees,
      "Avancement (%)": Math.round(r.avancement * 10) / 10,
      "Cout de revient reel (EUR)": Math.round(r.coutRevientReel * 100) / 100,
      "Benefice reel (EUR)": Math.round(r.beneficeReel * 100) / 100,
      "Marge previsionnelle (EUR)": Math.round(r.marge * 100) / 100,
    }));
    const feuille = XLSX.utils.json_to_sheet(donnees);
    feuille["!cols"] = Object.keys(donnees[0] || {}).map((k) => ({ wch: Math.max(k.length, 16) }));
    const classeur = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(classeur, feuille, "Devis");
    XLSX.writeFile(classeur, "SLK_Manager_liste_devis_" + new Date().toISOString().slice(0, 10) + ".xlsx");
  }

  function confirmerAjout(chantierId) {
    if (!choixAgent) return;
    affecterAgent(chantierId, choixAgent);
    setAjoutId(null);
    setChoixAgent("");
  }

  return (
    <div className="space-y-5">
      {(chantiersEnAlerte.length > 0 || sansChef > 0 || heuresAValider > 0 || notifsNonLues > 0) && (
        <Card className="p-5 no-print" style={{ borderColor: "#F0C4B4", background: "#FDF8F5" }}>
          <h3 className="text-[12.5px] font-semibold mb-3 flex items-center gap-1.5" style={{ color: INK }}>
            <AlertTriangle size={15} style={{ color: ACCENT_DEEP }} /> Actions prioritaires
          </h3>
          <div className="flex flex-wrap gap-2">
            {chantiersEnAlerte.length > 0 && (
              <span className="text-[12px] px-3 py-1.5 rounded-md" style={{ background: "#FBEBE5", color: BAD }}>
                {chantiersEnAlerte.length} chantier{chantiersEnAlerte.length > 1 ? "s" : ""} en alerte budget/planning
              </span>
            )}
            {sansChef > 0 && (
              <span className="text-[12px] px-3 py-1.5 rounded-md" style={{ background: "#FDF3E2", color: "#B5710A" }}>
                {sansChef} chantier{sansChef > 1 ? "s" : ""} sans chef affecte
              </span>
            )}
            {heuresAValider > 0 && setTab && (
              <button onClick={() => setTab("heures")} className="text-[12px] px-3 py-1.5 rounded-md font-medium" style={{ background: "#EAF1FB", color: "#1F5088" }}>
                {heuresAValider} saisie{heuresAValider > 1 ? "s" : ""} d'heures a valider &rarr;
              </button>
            )}
            {notifsNonLues > 0 && setTab && (
              <button onClick={() => setTab("notifications")} className="text-[12px] px-3 py-1.5 rounded-md font-medium" style={{ background: "#F1F2F4", color: "#5B6472" }}>
                {notifsNonLues} notification{notifsNonLues > 1 ? "s" : ""} non lue{notifsNonLues > 1 ? "s" : ""} &rarr;
              </button>
            )}
          </div>
        </Card>
      )}

      <div className="grid grid-cols-2 lg:grid-cols-5 gap-4">
        <KPI icon={LayoutDashboard} label="Chantiers en cours" value={rows.length} color="#2563EB" />
        <KPI icon={Wallet} label="Chiffre d'affaires" value={fmtEUR(totalCA)} color="#1F7A4D" />
        <KPI icon={Receipt} label="Cout de revient reel" value={fmtEUR(totalCoutRevient)} color="#B5710A" />
        <KPI icon={TrendingUp} label="Benefice reel" value={fmtEUR(totalBenefice)} alert={totalBenefice < 0} color="#7C3AED" />
        <KPI icon={Clock} label="Heures prevues / consommees" value={fmtH(totalHPrevues) + " / " + fmtH(totalHConsommees)} small color="#0F9488" />
      </div>

      <Card className="overflow-hidden">
        <div className="px-6 py-5 flex items-start justify-between gap-3" style={{ borderBottom: "1px solid " + BORDER }}>
          <div>
            <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Suivi multi-chantiers</h3>
            <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Devis accepte -&gt; personnel a affecter (plusieurs agents possibles, selon leur role) -&gt; alerte des {seuilAlerte}% des heures prevues consommees.</p>
            <p className="text-[11px] mt-1" style={{ color: MUTE }}>Perimetre : chantiers de {ANNEE_EN_COURS} + chantiers anterieurs non termines. Les chantiers termines d'une annee passee restent consultables dans le Bilan annuel (Comptabilite).</p>
          </div>
          {utilisateur.role === "direction" && rows.length > 0 && (
            <button onClick={exporterDevisExcel} className="flex items-center gap-1.5 text-[12px] font-semibold px-3 py-1.5 rounded-md shrink-0 no-print" style={{ color: "#1F7A4D", border: "1px solid #BFE3CE" }}>
              <FileText size={13} /> Exporter vers Excel
            </button>
          )}
        </div>
        <div className="overflow-x-auto">
        <table className="w-full text-[13px]">
          <thead>
            <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
              <th className="px-6 py-2.5 font-semibold">Chantier</th>
              <th className="px-3 py-2.5 font-semibold w-[220px]">Personnel affecte</th>
              <th className="px-3 py-2.5 font-semibold">Budget prevu</th>
              <th className="px-3 py-2.5 font-semibold">Heures (prev./conso.)</th>
              <th className="px-3 py-2.5 font-semibold">Avancement</th>
              <th className="px-3 py-2.5 font-semibold">Statut</th>
              <th className="px-3 py-2.5 font-semibold text-right">Marge restante</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => {
              const disponibles = personnel.filter((p) => !r.affectations.includes(p.id));
              return (
              <tr key={r.id} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                <td className="px-6 py-3.5 align-top">
                  <div className="font-medium flex items-center gap-2" style={{ color: INK }}>
                    {r.nom}
                    {r.enAlerte && <AlertTriangle size={13} style={{ color: BAD }} />}
                  </div>
                  <div className="text-[11px]" style={{ color: MUTE }}>{clientNom(r.clientId)}</div>
                </td>
                <td className="px-3 py-3.5 align-top">
                  <div className="flex flex-col gap-1 mb-1.5">
                    {r.affectations.map((id) => {
                      const a = agent(id);
                      if (!a) return null;
                      const isChef = a.role === "chef_chantier";
                      return (
                        <span key={id} className="inline-flex items-center gap-1.5 text-[11.5px] px-2 py-0.5 rounded-full w-fit"
                          style={{ background: isChef ? "#E9F2FB" : "#F1F2F4", color: isChef ? "#1F5088" : "#5B6472" }}>
                          {a.nom}
                          <span className="text-[9.5px] opacity-70">({a.roleLabel})</span>
                          {utilisateur.role === "direction" && (
                            <button onClick={() => retirerAgent(r.id, id)} className="opacity-60 hover:opacity-100"><XCircle size={11} /></button>
                          )}
                        </span>
                      );
                    })}
                    {r.affectations.length === 0 && <span className="text-[11.5px] italic" style={{ color: MUTE }}>Aucun agent affecte</span>}
                  </div>
                  {utilisateur.role === "direction" && (ajoutId === r.id ? (
                    <div className="flex items-center gap-1.5">
                      <select value={choixAgent} onChange={(e) => setChoixAgent(e.target.value)}
                        className="rounded-md px-2 py-1 text-[11.5px]" style={{ border: "1px solid " + BORDER }}>
                        <option value="">Choisir...</option>
                        {disponibles.map((p) => <option key={p.id} value={p.id}>{p.nom} ({p.roleLabel})</option>)}
                      </select>
                      <button onClick={() => confirmerAjout(r.id)} className="text-[11px] font-semibold px-2 py-1 rounded-md text-white" style={{ background: NAVY }}>OK</button>
                    </div>
                  ) : (
                    <button onClick={() => setAjoutId(r.id)} className="flex items-center gap-1 text-[11px] font-semibold px-2 py-1 rounded-md" style={{ color: ACCENT_DEEP, border: "1px dashed " + ACCENT }}>
                      <UserPlus size={11} /> Ajouter un agent
                    </button>
                  ))}
                </td>
                <td className="num px-3 py-3.5 text-[13px]">{fmtEUR(r.montantHT)}</td>
                <td className="num px-3 py-3.5 text-[13px]">{fmtH(r.hPrevues)} / {fmtH(r.heuresConsommees)}</td>
                <td className="px-3 py-3.5">
                  <div className="flex items-center gap-2">
                    <div className="w-24 h-1.5 rounded-full overflow-hidden" style={{ background: BORDER }}>
                      <div className="h-full rounded-full" style={{ width: Math.min(100, r.avancement) + "%", background: r.enAlerte ? BAD : GOOD }} />
                    </div>
                    <span className="num text-[11.5px]" style={{ color: MUTE }}>{r.avancement.toFixed(0)}%</span>
                  </div>
                </td>
                <td className="px-3 py-3.5">
                  {utilisateur.role === "direction" ? (
                    <select value={r.statut} onChange={(e) => changerStatutChantier(r.id, e.target.value)}
                      className="text-[11.5px] rounded-md px-2 py-1" style={{ border: "1px solid " + BORDER }}>
                      <option value="en_cours">En cours</option>
                      <option value="en_pause">En pause</option>
                      <option value="en_difficulte">En difficulte</option>
                      <option value="termine">Termine</option>
                    </select>
                  ) : (
                    <span className="text-[11.5px]" style={{ color: MUTE }}>{r.statut.replace("_", " ")}</span>
                  )}
                </td>
                <td className="num px-3 py-3.5 text-right font-semibold" style={{ color: r.marge < 0 ? BAD : INK }}>{fmtEUR(r.marge)}</td>
              </tr>
            );})}
          </tbody>
        </table>
        </div>
      </Card>
    </div>
  );
}

function KPI({ icon: Icon, label, value, alert, small, color }) {
  const teinte = alert ? BAD : (color || NAVY);
  return (
    <Card className="p-4" style={alert ? { borderColor: "#F0C4B4", background: "#FDF4F1" } : { borderTop: "2.5px solid " + teinte }}>
      <div className="w-8 h-8 rounded-md flex items-center justify-center mb-3" style={{ background: hexAlpha(teinte, 0.12), color: teinte }}>
        <Icon size={16} />
      </div>
      <div className={"num font-semibold " + (small ? "text-[14px]" : "text-[20px]")} style={{ color: INK }}>{value}</div>
      <div className="text-[11px] mt-0.5" style={{ color: MUTE }}>{label}</div>
    </Card>
  );
}

// ---------------------------------------------------------------------------
function RepertoiresTab({ clients, setClients, fournisseurs, setFournisseurs }) {
  const [subTab, setSubTab] = useState("clients");
  const [form, setForm] = useState({ nom: "", adresse: "", telephone: "", contact: "", specialite: "" });

  function addClient() {
    if (!form.nom) return;
    setClients((prev) => [...prev, { id: "cl" + Date.now(), nom: form.nom, adresse: form.adresse, telephone: form.telephone, contact: form.contact }]);
    setForm({ nom: "", adresse: "", telephone: "", contact: "", specialite: "" });
  }
  function addFournisseur() {
    if (!form.nom) return;
    setFournisseurs((prev) => [...prev, { id: "f" + Date.now(), nom: form.nom, specialite: form.specialite, telephone: form.telephone }]);
    setForm({ nom: "", adresse: "", telephone: "", contact: "", specialite: "" });
  }
  function removeClient(id) { setClients((prev) => prev.filter((c) => c.id !== id)); }
  function removeFournisseur(id) { setFournisseurs((prev) => prev.filter((f) => f.id !== id)); }

  return (
    <div className="space-y-5">
      <div className="flex gap-2">
        <button onClick={() => setSubTab("clients")} className="flex items-center gap-1.5 text-[12.5px] font-semibold px-3.5 py-2 rounded-md"
          style={subTab === "clients" ? { background: NAVY, color: "#fff" } : { border: "1px solid " + BORDER, color: MUTE }}>
          <Users size={14} /> Clients ({clients.length})
        </button>
        <button onClick={() => setSubTab("fournisseurs")} className="flex items-center gap-1.5 text-[12.5px] font-semibold px-3.5 py-2 rounded-md"
          style={subTab === "fournisseurs" ? { background: NAVY, color: "#fff" } : { border: "1px solid " + BORDER, color: MUTE }}>
          <Truck size={14} /> Fournisseurs ({fournisseurs.length})
        </button>
      </div>

      {subTab === "clients" && (
        <>
          <Card className="p-6">
            <h3 className="text-[13px] font-semibold mb-4" style={{ color: INK }}>Enregistrer un client</h3>
            <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
              <input value={form.nom} onChange={(e) => setForm({ ...form, nom: e.target.value })} placeholder="Nom du client"
                className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              <input value={form.adresse} onChange={(e) => setForm({ ...form, adresse: e.target.value })} placeholder="Adresse"
                className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              <input value={form.telephone} onChange={(e) => setForm({ ...form, telephone: e.target.value })} placeholder="Telephone"
                className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              <input value={form.contact} onChange={(e) => setForm({ ...form, contact: e.target.value })} placeholder="Contact (nom)"
                className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
            </div>
            <button onClick={addClient} className="mt-3 flex items-center gap-1.5 text-[12px] font-semibold text-white px-3.5 py-2 rounded-md" style={{ background: ACCENT }}>
              <Plus size={14} /> Ajouter le client
            </button>
          </Card>
          <Card className="overflow-hidden">
            <div className="overflow-x-auto">
            <table className="w-full text-[13px]">
              <thead>
                <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
                  <th className="px-6 py-2.5 font-semibold">Client</th>
                  <th className="px-3 py-2.5 font-semibold">Adresse</th>
                  <th className="px-3 py-2.5 font-semibold">Contact</th>
                  <th className="px-3 py-2.5"></th>
                </tr>
              </thead>
              <tbody>
                {clients.map((c, i) => (
                  <tr key={c.id} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                    <td className="px-6 py-2.5" style={{ color: INK }}>{c.nom}</td>
                    <td className="px-3 py-2.5" style={{ color: MUTE }}>{c.adresse}</td>
                    <td className="px-3 py-2.5" style={{ color: MUTE }}>{c.telephone} {c.contact ? " · " + c.contact : ""}</td>
                    <td className="px-3 py-2.5 text-right"><button onClick={() => removeClient(c.id)} style={{ color: BAD }}><Trash2 size={14} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
            </div>
          </Card>
        </>
      )}

      {subTab === "fournisseurs" && (
        <>
          <Card className="p-6">
            <h3 className="text-[13px] font-semibold mb-4" style={{ color: INK }}>Enregistrer un fournisseur</h3>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <input value={form.nom} onChange={(e) => setForm({ ...form, nom: e.target.value })} placeholder="Nom du fournisseur"
                className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              <input value={form.specialite} onChange={(e) => setForm({ ...form, specialite: e.target.value })} placeholder="Specialite"
                className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              <input value={form.telephone} onChange={(e) => setForm({ ...form, telephone: e.target.value })} placeholder="Telephone"
                className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
            </div>
            <button onClick={addFournisseur} className="mt-3 flex items-center gap-1.5 text-[12px] font-semibold text-white px-3.5 py-2 rounded-md" style={{ background: ACCENT }}>
              <Plus size={14} /> Ajouter le fournisseur
            </button>
          </Card>
          <Card className="overflow-hidden">
            <div className="overflow-x-auto">
            <table className="w-full text-[13px]">
              <thead>
                <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
                  <th className="px-6 py-2.5 font-semibold">Fournisseur</th>
                  <th className="px-3 py-2.5 font-semibold">Specialite</th>
                  <th className="px-3 py-2.5 font-semibold">Telephone</th>
                  <th className="px-3 py-2.5"></th>
                </tr>
              </thead>
              <tbody>
                {fournisseurs.map((f, i) => (
                  <tr key={f.id} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                    <td className="px-6 py-2.5" style={{ color: INK }}>{f.nom}</td>
                    <td className="px-3 py-2.5" style={{ color: MUTE }}>{f.specialite}</td>
                    <td className="px-3 py-2.5" style={{ color: MUTE }}>{f.telephone}</td>
                    <td className="px-3 py-2.5 text-right"><button onClick={() => removeFournisseur(f.id)} style={{ color: BAD }}><Trash2 size={14} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
function BoutonImprimer({ label = "Imprimer / PDF" }) {
  return (
    <button onClick={() => window.print()} className="no-print flex items-center gap-1.5 text-[12px] font-semibold px-3 py-1.5 rounded-md" style={{ border: "1px solid " + BORDER, color: INK }}>
      <FileText size={13} /> {label}
    </button>
  );
}

function StatutBadge({ statut }) {
  const info = STATUT_INFO[statut];
  return (
    <span className="text-[11px] font-semibold px-2 py-0.5 rounded-full" style={{ color: info.color, background: info.bg }}>
      {info.label}
    </span>
  );
}

function HeuresTab({ heuresListe, chantiers, utilisateur, saisirHeure, accepterHeure, refuserHeure, validerHeure }) {
  const [refusId, setRefusId] = useState(null);
  const [commentaire, setCommentaire] = useState("");
  const [nouvChantier, setNouvChantier] = useState(chantiers[0] ? chantiers[0].nom : "");
  const [nouvPoste, setNouvPoste] = useState("");
  const [nouvHeures, setNouvHeures] = useState(7.5);

  const chantierIds = chantiers.map((c) => c.id);
  const mesSaisies = heuresListe.filter((h) => h.ouvrier === utilisateur.nom);
  const aValider = heuresListe.filter((h) => h.statut === "a_valider" && chantierIds.includes(h.chantierId));
  const acceptees = heuresListe.filter((h) => h.statut === "acceptee" && chantierIds.includes(h.chantierId));
  const historique = heuresListe.filter((h) => (h.statut === "validee" || h.statut === "refusee") && chantierIds.includes(h.chantierId));

  function handleSaisir() {
    if (!nouvPoste || nouvHeures <= 0) return;
    const chantierChoisi = chantiers.find((c) => c.nom === nouvChantier);
    saisirHeure({ chantier: nouvChantier, chantierId: chantierChoisi ? chantierChoisi.id : null, heures: nouvHeures, poste: nouvPoste });
    setNouvPoste("");
    setNouvHeures(7.5);
  }

  return (
    <div className="space-y-5">
      <div className="flex items-center gap-2 px-4 py-2.5 rounded-md text-[12.5px]" style={{ background: "#EAF1FB", color: "#1F5088" }}>
        <UserCog size={15} />
        Connecte en tant que <strong>{utilisateur.roleLabel}</strong> ({utilisateur.nom}) &mdash; les actions disponibles ci-dessous dependent de ce profil.
      </div>

      {utilisateur.role === "ouvrier" && (
        <>
          <Card className="p-6">
            <h3 className="text-[13.5px] font-semibold mb-4" style={{ color: INK }}>Saisir mes heures du jour</h3>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <Field label="Chantier">
                <select value={nouvChantier} onChange={(e) => setNouvChantier(e.target.value)} className="w-full rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }}>
                  {chantiers.map((c) => <option key={c.id} value={c.nom}>{c.nom}</option>)}
                </select>
              </Field>
              <Field label="Poste de travail">
                <input value={nouvPoste} onChange={(e) => setNouvPoste(e.target.value)} placeholder="Ex. Gaine D250 - zone AS-Bureau"
                  className="w-full rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              </Field>
              <Field label="Heures travaillees">
                <input type="number" step="0.5" value={nouvHeures} onChange={(e) => setNouvHeures(parseFloat(e.target.value) || 0)}
                  className="num w-full rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              </Field>
            </div>
            <button onClick={handleSaisir} className="mt-4 text-white text-[12.5px] font-semibold px-4 py-2.5 rounded-md" style={{ background: ACCENT }}>
              Envoyer ma saisie
            </button>
          </Card>

          <Card className="overflow-hidden">
            <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
              <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Mes saisies recentes</h3>
            </div>
            <div className="divide-y" style={{ borderColor: BORDER }}>
              {mesSaisies.map((h) => (
                <div key={h.id} className="px-6 py-3.5 flex items-center justify-between">
                  <div>
                    <div className="text-[13px]" style={{ color: INK }}>{h.poste}</div>
                    <div className="text-[11.5px]" style={{ color: MUTE }}>{h.chantier} &middot; <span className="num">{h.heures} h</span></div>
                    {h.commentaire && <div className="text-[11px] mt-1" style={{ color: BAD }}>{h.commentaire}</div>}
                  </div>
                  <StatutBadge statut={h.statut} />
                </div>
              ))}
            </div>
          </Card>
        </>
      )}

      {utilisateur.role === "chef_chantier" && (
        <Card className="overflow-hidden">
          <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
            <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>A valider ({aValider.length})</h3>
            <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Controle de 1er niveau avant transmission a la Direction.</p>
          </div>
          {aValider.length === 0 ? (
            <div className="px-6 py-8 text-center text-[13px]" style={{ color: MUTE }}>Aucune saisie en attente.</div>
          ) : (
            <div className="divide-y" style={{ borderColor: BORDER }}>
              {aValider.map((h) => (
                <div key={h.id} className="px-6 py-4">
                  <div className="flex items-center justify-between">
                    <div>
                      <div className="text-[13px] font-medium" style={{ color: INK }}>{h.ouvrier} &middot; <span className="num">{h.heures} h</span></div>
                      <div className="text-[11.5px]" style={{ color: MUTE }}>{h.poste} &middot; {h.chantier}</div>
                    </div>
                    <div className="flex items-center gap-2">
                      <button onClick={() => accepterHeure(h.id)} className="flex items-center gap-1.5 text-[12px] font-semibold px-3 py-1.5 rounded-md text-white" style={{ background: GOOD }}>
                        <CheckCircle2 size={14} /> Accepter
                      </button>
                      <button onClick={() => setRefusId(refusId === h.id ? null : h.id)} className="flex items-center gap-1.5 text-[12px] font-semibold px-3 py-1.5 rounded-md" style={{ border: "1px solid " + BORDER, color: BAD }}>
                        <XCircle size={14} /> Refuser
                      </button>
                    </div>
                  </div>
                  {refusId === h.id && (
                    <div className="mt-3 flex items-center gap-2">
                      <input value={commentaire} onChange={(e) => setCommentaire(e.target.value)} placeholder="Motif du refus (obligatoire)"
                        className="flex-1 rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
                      <button
                        onClick={() => { if (commentaire) { refuserHeure(h.id, commentaire); setCommentaire(""); setRefusId(null); } }}
                        className="text-[12px] font-semibold px-3 py-2 rounded-md text-white" style={{ background: BAD }}
                      >
                        Confirmer le refus
                      </button>
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
        </Card>
      )}

      {utilisateur.role === "direction" && (
        <Card className="overflow-hidden">
          <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
            <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Validation finale ({acceptees.length})</h3>
            <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Deja acceptees par le chef de chantier &mdash; verrouillage definitif apres validation.</p>
          </div>
          {acceptees.length === 0 ? (
            <div className="px-6 py-8 text-center text-[13px]" style={{ color: MUTE }}>Aucune saisie en attente de validation finale.</div>
          ) : (
            <div className="divide-y" style={{ borderColor: BORDER }}>
              {acceptees.map((h) => (
                <div key={h.id} className="px-6 py-4 flex items-center justify-between">
                  <div>
                    <div className="text-[13px] font-medium" style={{ color: INK }}>{h.ouvrier} &middot; <span className="num">{h.heures} h</span></div>
                    <div className="text-[11.5px]" style={{ color: MUTE }}>{h.poste} &middot; {h.chantier}</div>
                  </div>
                  <button onClick={() => validerHeure(h.id)} className="flex items-center gap-1.5 text-[12px] font-semibold px-3.5 py-1.5 rounded-md text-white" style={{ background: NAVY }}>
                    <ShieldCheck size={14} /> Valider definitivement
                  </button>
                </div>
              ))}
            </div>
          )}
        </Card>
      )}

      <Card className="overflow-hidden">
        <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Historique</h3>
        </div>
        <div className="divide-y" style={{ borderColor: BORDER }}>
          {historique.map((h) => (
            <div key={h.id} className="px-6 py-3 flex items-center justify-between">
              <div>
                <div className="text-[13px]" style={{ color: INK }}>{h.ouvrier} &middot; {h.poste} &middot; <span className="num">{h.heures} h</span></div>
                <div className="text-[11px]" style={{ color: MUTE }}>{h.chantier}</div>
                {h.commentaire && <div className="text-[11px] mt-0.5" style={{ color: BAD }}>{h.commentaire}</div>}
              </div>
              <StatutBadge statut={h.statut} />
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Module 16 — Notifications. Un chef de chantier voit ses propres notifications
// ET celles de ses ouvriers (supervision) ; un ouvrier ne voit que les siennes ;
// la Direction voit tout.
function NotificationsTab({ notifications, setNotifications, utilisateur, chantiers, ouvrirPartieConcernee }) {
  const mesChantierIds = chantiersActifs(utilisateur, chantiers).map((c) => c.id);
  const visibles = notifications.filter((n) => {
    if (utilisateur.role === "direction") return true;
    if (utilisateur.role === "chef_chantier") {
      if (n.pourNom === utilisateur.nom) return true;
      return n.pourRole === "ouvrier" && mesChantierIds.includes(n.chantierId);
    }
    return n.pourRole === "ouvrier" && n.pourNom === utilisateur.nom;
  });
  const nonLues = visibles.filter((n) => !n.lu).length;

  function marquerLu(id) {
    setNotifications((prev) => prev.map((n) => (n.id === id ? { ...n, lu: true } : n)));
  }
  function handleClic(n) {
    marquerLu(n.id);
    if (n.chantierId && ouvrirPartieConcernee) ouvrirPartieConcernee(n.chantierId);
  }

  return (
    <div className="space-y-5">
      {utilisateur.role === "chef_chantier" && (
        <div className="flex items-center gap-2 px-4 py-2.5 rounded-md text-[12.5px]" style={{ background: "#EAF1FB", color: "#1F5088" }}>
          <UserCog size={15} /> Vous voyez vos notifications, ainsi que celles concernant vos ouvriers.
        </div>
      )}
      <Card className="overflow-hidden">
        <div className="px-6 py-4 flex items-center justify-between" style={{ borderBottom: "1px solid " + BORDER }}>
          <div>
            <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Notifications</h3>
            <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>{nonLues} non lue{nonLues > 1 ? "s" : ""} &middot; cliquez pour ouvrir la partie concernee</p>
          </div>
          <Bell size={18} style={{ color: MUTE }} />
        </div>
        {visibles.length === 0 ? (
          <div className="px-6 py-8 text-center text-[13px]" style={{ color: MUTE }}>Aucune notification.</div>
        ) : (
          <div className="divide-y" style={{ borderColor: BORDER }}>
            {visibles.map((n) => (
              <div key={n.id} onClick={() => handleClic(n)} className="px-6 py-3.5 flex items-start gap-3 cursor-pointer" style={{ background: n.lu ? "transparent" : "#FDF8F5" }}>
                <div className="w-2 h-2 rounded-full mt-1.5 shrink-0" style={{ background: n.lu ? "transparent" : ACCENT }} />
                <div className="flex-1">
                  <div className="text-[13px]" style={{ color: INK }}>{n.message}</div>
                  <div className="text-[11px] mt-0.5" style={{ color: MUTE }}>
                    {n.pourNom && utilisateur.role !== "ouvrier" ? n.pourNom + " · " : ""}{n.quand}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Modules 7/12/13/14 — Plan de chantier, avancement graphique et photos.
// Vue simplifiee en liste (l'annotation visuelle sur le PDF lui-meme est un
// developpement frontend a part entiere, non couvert par ce prototype).
function PlanTab({ chantiers, clients, planElements, changerStatutElement, historiqueAvancement, utilisateur, personnel, tauxDefaut, pctDefaut, chantierCible, plansPdf, importerPlanPdf, envoyerFicheAffectation }) {
  const [chantierId, setChantierId] = useState(chantierCible || (chantiers[0] ? chantiers[0].id : ""));
  const [photos, setPhotos] = useState({});
  const [formPlanOuvert, setFormPlanOuvert] = useState(false);
  const [nvPlanNumero, setNvPlanNumero] = useState("");
  const [nvPlanIndice, setNvPlanIndice] = useState("");
  const [nvPlanFichier, setNvPlanFichier] = useState(null);
  const [historiquePlanOuvert, setHistoriquePlanOuvert] = useState(null);

  // Arrivee depuis une notification (ex. message lie a ce chantier) : bascule
  // automatiquement sur le chantier concerne, sans action supplementaire.
  useEffect(() => {
    if (chantierCible) setChantierId(chantierCible);
  }, [chantierCible]);

  const peutModifier = utilisateur.role === "direction" || utilisateur.role === "chef_chantier";
  const chantier = chantiers.find((c) => c.id === chantierId);
  const client = chantier ? clients.find((cl) => cl.id === chantier.clientId) : null;
  const elements = planElements.filter((e) => e.chantierId === chantierId);
  const total = elements.length;
  const poses = elements.filter((e) => e.statut === "pose").length;
  const pctAvancement = total > 0 ? Math.round((poses / total) * 100) : 0;
  const historiqueChantier = historiqueAvancement.filter((h) => h.chantierId === chantierId);

  const agentsAffectes = chantier ? chantier.affectations.map((id) => personnel.find((p) => p.id === id)).filter(Boolean) : [];

  const hPrevues = chantier ? heuresPrevues(chantier.montantHT, pctDefaut, tauxDefaut) : 0;
  const hRestantes = chantier ? Math.max(0, hPrevues - chantier.heuresConsommees) : 0;
  const joursRestants = hRestantes / HEURES_PAR_JOUR;

  function ajouterPhoto(id, file) {
    if (!file) return;
    const url = URL.createObjectURL(file);
    setPhotos((prev) => ({ ...prev, [id]: [...(prev[id] || []), { url, par: utilisateur.nom, quand: "a l'instant" }] }));
  }

  const plansChantier = plansPdf.filter((p) => p.chantierId === chantierId);
  const numerosDePlan = Array.from(new Set(plansChantier.map((p) => p.planNumero)));
  const peutGererPlans = utilisateur.role === "direction" || utilisateur.role === "chef_chantier";

  function soumettreNouveauPlan() {
    if (!nvPlanNumero || !nvPlanIndice) return;
    importerPlanPdf(chantierId, nvPlanNumero, nvPlanIndice, nvPlanFichier);
    setNvPlanNumero("");
    setNvPlanIndice("");
    setNvPlanFichier(null);
    setFormPlanOuvert(false);
  }

  return (
    <div className="space-y-5">
      <Card className="p-4 no-print">
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <Field label="Chantier">
            <select value={chantierId} onChange={(e) => setChantierId(e.target.value)} className="w-full max-w-sm rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }}>
              {chantiers.map((c) => <option key={c.id} value={c.id}>{c.nom}</option>)}
            </select>
          </Field>
          <div className="flex items-center gap-2">
            {peutGererPlans && envoyerFicheAffectation && (
              <button onClick={() => envoyerFicheAffectation(chantierId)} className="flex items-center gap-1.5 text-[12px] font-semibold px-3 py-1.5 rounded-md" style={{ color: "#1F5088", border: "1px solid #C7DCF2" }}>
                <Send size={13} /> Envoyer la fiche aux agents affectes
              </button>
            )}
            <BoutonImprimer label="Imprimer la fiche" />
          </div>
        </div>
      </Card>

      {/* En-tete visible seulement a l'impression : fiche d'affectation complete
          (chef de chantier, equipe, adresse) + situation d'avancement. */}
      <div className="print-only">
        <h1 className="text-[18px] font-bold" style={{ color: INK }}>Fiche de chantier &mdash; {chantier ? chantier.nom : ""}</h1>
        <p className="text-[12px] mt-1" style={{ color: MUTE }}>Client : {client ? client.nom : "-"}{chantier && chantier.societe ? " (via " + chantier.societe + ")" : ""}</p>
        <p className="text-[12px]" style={{ color: MUTE }}>Adresse du projet : {client && client.adresse ? client.adresse : "Non renseignee"}</p>
        <p className="text-[12px]" style={{ color: MUTE }}>
          Chef de chantier : {agentsAffectes.filter((a) => a.role === "chef_chantier").map((a) => a.nom).join(", ") || "Non affecte"}
        </p>
        <p className="text-[12px]" style={{ color: MUTE }}>
          Ouvriers affectes : {agentsAffectes.filter((a) => a.role === "ouvrier").map((a) => a.nom).join(", ") || "Aucun"}
        </p>
        <p className="text-[12px] mt-1" style={{ color: MUTE }}>Edite le {new Date().toLocaleDateString("fr-FR")} par {utilisateur.nom}</p>
      </div>

      <Card className="overflow-hidden no-print">
        <div className="px-6 py-4 flex items-center justify-between" style={{ borderBottom: "1px solid " + BORDER }}>
          <div>
            <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Plans PDF du chantier</h3>
            <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Une seule version active par numero de plan ; les indices precedents restent consultables.</p>
          </div>
          {peutGererPlans && (
            <button onClick={() => setFormPlanOuvert((v) => !v)} className="flex items-center gap-1.5 text-[12px] font-semibold px-3 py-1.5 rounded-md" style={{ color: ACCENT_DEEP, border: "1px dashed " + ACCENT }}>
              <Plus size={13} /> Importer un plan
            </button>
          )}
        </div>

        {formPlanOuvert && (
          <div className="px-6 py-4" style={{ background: BG, borderBottom: "1px solid " + BORDER }}>
            <div className="grid grid-cols-1 sm:grid-cols-4 gap-2 items-end">
              <Field label="Numero de plan">
                <input value={nvPlanNumero} onChange={(e) => setNvPlanNumero(e.target.value)} placeholder="Ex. 103"
                  className="w-full rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }} />
              </Field>
              <Field label="Indice de revision">
                <input value={nvPlanIndice} onChange={(e) => setNvPlanIndice(e.target.value)} placeholder="Ex. IND A"
                  className="w-full rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }} />
              </Field>
              <Field label="Fichier PDF">
                <input type="file" accept="application/pdf" onChange={(e) => setNvPlanFichier(e.target.files[0])}
                  className="w-full text-[12px]" />
              </Field>
              <button onClick={soumettreNouveauPlan} className="text-[12px] font-semibold text-white px-3 py-2 rounded-md" style={{ background: NAVY }}>Importer</button>
            </div>
          </div>
        )}

        {numerosDePlan.length === 0 ? (
          <div className="px-6 py-8 text-center text-[13px]" style={{ color: MUTE }}>Aucun plan importe pour ce chantier.</div>
        ) : (
          <div className="divide-y" style={{ borderColor: BORDER }}>
            {numerosDePlan.map((numero) => {
              const versions = plansChantier.filter((p) => p.planNumero === numero).sort((a, b) => (a.dateImport < b.dateImport ? 1 : -1));
              const actif = versions.find((v) => v.actif) || versions[0];
              const historique = versions.filter((v) => v.id !== actif.id);
              return (
                <div key={numero} className="px-6 py-3.5">
                  <div className="flex items-center justify-between gap-3">
                    <div className="flex items-center gap-2.5">
                      <FileText size={16} style={{ color: MUTE }} />
                      <div>
                        <div className="text-[13px]" style={{ color: INK }}>Plan n&deg;{numero} <span className="text-[11.5px] font-semibold px-1.5 py-0.5 rounded-full ml-1" style={{ color: "#1F5088", background: "#E9F2FB" }}>{actif.indice}</span></div>
                        <div className="text-[11px]" style={{ color: MUTE }}>Importe le {actif.dateImport} par {actif.importePar}</div>
                      </div>
                    </div>
                    <div className="flex items-center gap-3">
                      {actif.fichierUrl ? (
                        <a href={actif.fichierUrl} target="_blank" rel="noreferrer" className="text-[11.5px] font-semibold" style={{ color: ACCENT_DEEP }}>Voir le PDF</a>
                      ) : (
                        <span className="text-[11px] italic" style={{ color: MUTE }}>Fichier non joint (demo)</span>
                      )}
                      {historique.length > 0 && (
                        <button onClick={() => setHistoriquePlanOuvert(historiquePlanOuvert === numero ? null : numero)} className="text-[11.5px] font-medium" style={{ color: MUTE }}>
                          {historiquePlanOuvert === numero ? "Masquer" : "Historique"} ({historique.length})
                        </button>
                      )}
                    </div>
                  </div>
                  {historiquePlanOuvert === numero && (
                    <div className="mt-2.5 ml-7 space-y-1.5">
                      {historique.map((v) => (
                        <div key={v.id} className="text-[11.5px] flex items-center gap-2" style={{ color: MUTE }}>
                          <span className="px-1.5 py-0.5 rounded-full" style={{ background: "#F1F2F4" }}>{v.indice}</span>
                          importe le {v.dateImport} par {v.importePar}
                          {v.fichierUrl && <a href={v.fichierUrl} target="_blank" rel="noreferrer" style={{ color: ACCENT_DEEP }}>Voir</a>}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </Card>

      <Card className="p-6">
        <h3 className="text-[13.5px] font-semibold mb-3" style={{ color: INK }}>Personnel affecte a ce chantier</h3>
        {agentsAffectes.length === 0 ? (
          <p className="text-[13px] italic" style={{ color: MUTE }}>Aucun agent affecte pour le moment.</p>
        ) : (
          <div className="flex flex-wrap gap-2">
            {agentsAffectes.map((a) => (
              <span key={a.id} className="inline-flex items-center gap-1.5 text-[12px] px-2.5 py-1 rounded-full"
                style={{ background: a.role === "chef_chantier" ? "#E9F2FB" : "#F1F2F4", color: a.role === "chef_chantier" ? "#1F5088" : "#5B6472" }}>
                {a.nom} <span className="opacity-70">({a.roleLabel})</span>
              </span>
            ))}
          </div>
        )}
      </Card>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <Card className="p-5">
          <div className="text-[11px] font-semibold uppercase tracking-wide mb-1" style={{ color: MUTE }}>Avancement global</div>
          <div className="num text-[22px] font-semibold" style={{ color: INK }}>{pctAvancement}%</div>
          <div className="h-2 rounded-full overflow-hidden mt-2" style={{ background: BORDER }}>
            <div className="h-full rounded-full" style={{ width: pctAvancement + "%", background: GOOD }} />
          </div>
        </Card>
        <Card className="p-5">
          <div className="text-[11px] font-semibold uppercase tracking-wide mb-1" style={{ color: MUTE }}>Heures consommees / prevues</div>
          <div className="num text-[16px] font-semibold" style={{ color: INK }}>{chantier ? fmtH(chantier.heuresConsommees) : "-"} / {fmtH(hPrevues)}</div>
        </Card>
        <Card className="p-5">
          <div className="text-[11px] font-semibold uppercase tracking-wide mb-1" style={{ color: MUTE }}>Temps restant estime</div>
          <div className="num text-[16px] font-semibold" style={{ color: hRestantes === 0 ? BAD : INK }}>{fmtH(hRestantes)} ({fmtJours(joursRestants)})</div>
        </Card>
      </div>

      <Card className="overflow-hidden">
        <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Elements du plan</h3>
          <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>{peutModifier ? "Statut modifiable directement." : "Lecture seule."}</p>
        </div>
        <div className="divide-y" style={{ borderColor: BORDER }}>
          {elements.map((e) => {
            const info = COULEUR_STATUT[e.statut];
            const elemPhotos = photos[e.id] || [];
            return (
              <div key={e.id} className="px-6 py-3.5">
                <div className="flex items-center justify-between gap-3">
                  <div className="flex items-center gap-2.5">
                    <span className="w-2.5 h-2.5 rounded-full" style={{ background: info.dot }} />
                    <div>
                      <div className="text-[13px]" style={{ color: INK }}>{e.designation}</div>
                      <div className="num text-[11px]" style={{ color: MUTE }}>{e.quantitePosee} / {e.quantitePrevue}</div>
                    </div>
                  </div>
                  <div className="flex items-center gap-2">
                    {peutModifier ? (
                      <select value={e.statut} onChange={(ev) => changerStatutElement(e.id, ev.target.value)}
                        className="no-print text-[11.5px] font-semibold px-2 py-1 rounded-full border-none" style={{ color: info.color, background: info.bg }}>
                        {Object.entries(COULEUR_STATUT).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                      </select>
                    ) : null}
                    <span className="print-only-inline text-[11.5px] font-semibold px-2 py-0.5 rounded-full" style={{ color: info.color, background: info.bg }}>{info.label}</span>
                    {peutModifier && (
                      <label className="no-print cursor-pointer" title="Ajouter une photo">
                        <Camera size={15} style={{ color: MUTE }} />
                        <input type="file" accept="image/*" className="hidden" onChange={(ev) => ajouterPhoto(e.id, ev.target.files[0])} />
                      </label>
                    )}
                  </div>
                </div>
                {elemPhotos.length > 0 && (
                  <div className="flex gap-2 mt-2.5 ml-5">
                    {elemPhotos.map((p, i) => (
                      <div key={i} className="relative">
                        <img src={p.url} alt="" className="w-14 h-14 object-cover rounded-md" style={{ border: "1px solid " + BORDER }} />
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
          {elements.length === 0 && <div className="px-6 py-8 text-center text-[13px]" style={{ color: MUTE }}>Aucun element de plan pour ce chantier.</div>}
        </div>
      </Card>

      <Card className="overflow-hidden">
        <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Historique de l'avancement</h3>
          <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Chaque changement de statut est trace avec l'auteur et la date.</p>
        </div>
        {historiqueChantier.length === 0 ? (
          <div className="px-6 py-6 text-center text-[13px]" style={{ color: MUTE }}>Aucun changement enregistre pour ce chantier.</div>
        ) : (
          <div className="divide-y" style={{ borderColor: BORDER }}>
            {historiqueChantier.map((h) => (
              <div key={h.id} className="px-6 py-2.5 text-[12.5px] flex items-center justify-between">
                <div>
                  <span style={{ color: INK }}>{h.designation}</span>
                  <span style={{ color: MUTE }}> &mdash; {COULEUR_STATUT[h.avant].label} &rarr; {COULEUR_STATUT[h.apres].label}</span>
                </div>
                <div className="text-[11px]" style={{ color: MUTE }}>{h.par} &middot; {h.quand}</div>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Module 15 — Suivi de productivite : compare le temps de pose theorique
// (bibliotheque de prix) au temps reellement passe (heures saisies).
function ProductiviteTab({ heuresListe, library, utilisateur, chantiers, tauxDefaut, pctDefaut }) {
  const chantierIds = chantiers.map((c) => c.id);
  const parOuvrier = {};
  heuresListe.filter((h) => (h.statut === "validee" || h.statut === "acceptee") && chantierIds.includes(h.chantierId)).forEach((h) => {
    parOuvrier[h.ouvrier] = parOuvrier[h.ouvrier] || { heures: 0, saisies: 0 };
    parOuvrier[h.ouvrier].heures += h.heures;
    parOuvrier[h.ouvrier].saisies += 1;
  });
  const rows = Object.entries(parOuvrier).map(([nom, d]) => ({ nom, ...d, moyenne: d.heures / d.saisies }));

  const avecTempsPose = library.filter((l) => l.tempsPose !== null);

  // Situation par rapport au projet : heures prevues (calcul module 1) vs
  // heures reellement consommees, chantier par chantier.
  const situationProjet = chantiers.map((c) => {
    const hPrevues = heuresPrevues(c.montantHT, pctDefaut, tauxDefaut);
    const ecart = hPrevues - c.heuresConsommees;
    const pct = hPrevues > 0 ? Math.round((c.heuresConsommees / hPrevues) * 100) : 0;
    return { ...c, hPrevues, ecart, pct };
  });

  const historique = heuresListe
    .filter((h) => (h.statut === "validee" || h.statut === "refusee") && chantierIds.includes(h.chantierId))
    .slice(0, 20);

  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between no-print">
        {utilisateur.role === "chef_chantier" ? (
          <div className="flex items-center gap-2 px-4 py-2.5 rounded-md text-[12.5px]" style={{ background: "#EAF1FB", color: "#1F5088" }}>
            <UserCog size={15} /> Vue limitee a vos ouvriers.
          </div>
        ) : <div />}
        <BoutonImprimer label="Imprimer le rapport" />
      </div>

      <Card className="overflow-hidden">
        <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Situation par rapport au projet</h3>
          <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Heures prevues (module 1) comparees aux heures reellement consommees, par chantier.</p>
        </div>
        <div className="overflow-x-auto">
        <table className="w-full text-[13px]">
          <thead>
            <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
              <th className="px-6 py-2.5 font-semibold">Chantier</th>
              <th className="px-3 py-2.5 font-semibold">Prevues</th>
              <th className="px-3 py-2.5 font-semibold">Consommees</th>
              <th className="px-3 py-2.5 font-semibold">Ecart</th>
              <th className="px-3 py-2.5 font-semibold text-right">Avancement</th>
            </tr>
          </thead>
          <tbody>
            {situationProjet.map((c, i) => (
              <tr key={c.id} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                <td className="px-6 py-2.5" style={{ color: INK }}>{c.nom}</td>
                <td className="num px-3 py-2.5" style={{ color: MUTE }}>{fmtH(c.hPrevues)}</td>
                <td className="num px-3 py-2.5" style={{ color: MUTE }}>{fmtH(c.heuresConsommees)}</td>
                <td className="num px-3 py-2.5" style={{ color: c.ecart < 0 ? BAD : GOOD }}>{c.ecart >= 0 ? "+" : ""}{fmtH(c.ecart)}</td>
                <td className="num px-3 py-2.5 text-right font-semibold" style={{ color: c.pct >= 90 ? BAD : INK }}>{c.pct}%</td>
              </tr>
            ))}
          </tbody>
        </table>
        </div>
      </Card>

      <Card className="overflow-hidden">
        <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Heures par ouvrier</h3>
          <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Base sur les saisies acceptees ou validees.</p>
        </div>
        <div className="overflow-x-auto">
        <table className="w-full text-[13px]">
          <thead>
            <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
              <th className="px-6 py-2.5 font-semibold">Ouvrier</th>
              <th className="px-3 py-2.5 font-semibold">Saisies</th>
              <th className="px-3 py-2.5 font-semibold">Total heures</th>
              <th className="px-3 py-2.5 font-semibold text-right">Moyenne / saisie</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={r.nom} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                <td className="px-6 py-2.5" style={{ color: INK }}>{r.nom}</td>
                <td className="num px-3 py-2.5" style={{ color: MUTE }}>{r.saisies}</td>
                <td className="num px-3 py-2.5" style={{ color: INK }}>{fmtH(r.heures)}</td>
                <td className="num px-3 py-2.5 text-right" style={{ color: MUTE }}>{fmtH(r.moyenne)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        </div>
      </Card>

      <Card className="overflow-hidden">
        <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Historique des saisies traitees</h3>
        </div>
        {historique.length === 0 ? (
          <div className="px-6 py-6 text-center text-[13px]" style={{ color: MUTE }}>Aucune saisie traitee pour le moment.</div>
        ) : (
          <div className="divide-y" style={{ borderColor: BORDER }}>
            {historique.map((h) => (
              <div key={h.id} className="px-6 py-2.5 flex items-center justify-between text-[12.5px]">
                <div>
                  <span style={{ color: INK }}>{h.ouvrier}</span>
                  <span style={{ color: MUTE }}> &mdash; {h.poste} ({h.chantier}) &middot; <span className="num">{h.heures} h</span></span>
                </div>
                <div className="flex items-center gap-2">
                  <StatutBadge statut={h.statut} />
                  <span className="text-[11px]" style={{ color: MUTE }}>{h.quand}</span>
                </div>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card className="overflow-hidden">
        <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Temps de pose de reference (module 18)</h3>
          <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Articles pour lesquels un temps de pose theorique est deja renseigne.</p>
        </div>
        <div className="divide-y" style={{ borderColor: BORDER }}>
          {avecTempsPose.map((l) => (
            <div key={l.id} className="px-6 py-2.5 flex items-center justify-between text-[13px]">
              <span style={{ color: INK }}>{l.designation}</span>
              <span className="num" style={{ color: MUTE }}>{l.tempsPose} h / {l.unite}</span>
            </div>
          ))}
          {avecTempsPose.length === 0 && <div className="px-6 py-6 text-center text-[13px]" style={{ color: MUTE }}>Aucun temps de pose renseigne pour le moment.</div>}
        </div>
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Communication interne — hierarchisee (voir correspondantsAutorises) et
// tracee : chaque message reste visible dans le tableau recapitulatif pour
// la Direction, meme apres lecture.
function MessagerieTab({ messages, envoyerMessage, marquerMessageLu, utilisateur, utilisateursSysteme, chantiers, correspondants }) {
  const [versId, setVersId] = useState(correspondants[0] ? correspondants[0].id : "");
  const [chantierId, setChantierId] = useState("");
  const [sujet, setSujet] = useState("");
  const [texte, setTexte] = useState("");
  const [repondreA, setRepondreA] = useState(null);
  const [texteReponse, setTexteReponse] = useState("");

  const nomDe = (id) => (utilisateursSysteme.find((u) => u.id === id) || {}).nom || "Inconnu";
  const roleLabelDe = (id) => (utilisateursSysteme.find((u) => u.id === id) || {}).roleLabel || "";
  const chantierNom = (id) => (chantiers.find((c) => c.id === id) || {}).nom || "-";

  const mesChantiers = chantiersActifs(utilisateur, chantiers);

  const recus = messages.filter((m) => m.versId === utilisateur.id);
  const boiteReception = [...recus].sort((a, b) => {
    const urgenceA = a.necessiteReponse && !a.repondu ? 0 : a.lu ? 2 : 1;
    const urgenceB = b.necessiteReponse && !b.repondu ? 0 : b.lu ? 2 : 1;
    return urgenceA - urgenceB;
  });
  const enAttenteDeReponse = recus.filter((m) => m.necessiteReponse && !m.repondu).length;

  // Tableau recapitulatif pour la tracabilite : la Direction voit tout ;
  // un chef de chantier voit les echanges de ses chantiers ; un ouvrier
  // ne voit que ses propres echanges.
  const idsMesChantiers = mesChantiers.map((c) => c.id);
  const tableauTracabilite =
    utilisateur.role === "direction"
      ? messages
      : messages.filter((m) => idsMesChantiers.includes(m.chantierId) || m.deId === utilisateur.id || m.versId === utilisateur.id);

  function handleEnvoyer() {
    if (!versId || !sujet || !texte) return;
    envoyerMessage({ versId, chantierId: chantierId || null, sujet, message: texte, necessiteReponse: true });
    setSujet("");
    setTexte("");
  }
  function handleRepondre(m) {
    if (!texteReponse) return;
    envoyerMessage({ versId: m.deId, chantierId: m.chantierId, sujet: "RE: " + m.sujet, message: texteReponse, necessiteReponse: false, repondA: m.id });
    setTexteReponse("");
    setRepondreA(null);
  }

  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between no-print">
        <div className="flex items-center gap-2 px-4 py-2.5 rounded-md text-[12.5px]" style={{ background: "#EAF1FB", color: "#1F5088" }}>
          <UserCog size={15} /> Hierarchie stricte : {utilisateur.role === "direction" ? "vous ecrivez uniquement aux chefs de chantier" : utilisateur.role === "chef_chantier" ? "vous ecrivez uniquement a votre equipe affectee ; vous repondez a la Direction sans initier" : "vous ne pouvez que repondre a votre chef de chantier"}.
        </div>
        <BoutonImprimer label="Imprimer le tableau" />
      </div>

      <div className="print-only">
        <h1 className="text-[18px] font-bold" style={{ color: INK }}>Tableau recapitulatif de la communication interne</h1>
        <p className="text-[12px]" style={{ color: MUTE }}>Edite le {new Date().toLocaleDateString("fr-FR")} par {utilisateur.nom}</p>
      </div>

      {enAttenteDeReponse > 0 && (
        <Card className="p-4 no-print" style={{ borderColor: "#F0C4B4", background: "#FDF8F5" }}>
          <div className="text-[12.5px] font-semibold flex items-center gap-1.5" style={{ color: BAD }}>
            <AlertTriangle size={14} /> {enAttenteDeReponse} message{enAttenteDeReponse > 1 ? "s" : ""} en attente d'une reponse obligatoire.
          </div>
        </Card>
      )}

      {correspondants.length > 0 && (
        <Card className="p-6 no-print">
          <h3 className="text-[13px] font-semibold mb-4" style={{ color: INK }}>Nouveau message</h3>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
            <Field label="Destinataire">
              <select value={versId} onChange={(e) => setVersId(e.target.value)} className="w-full rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }}>
                {correspondants.map((c) => <option key={c.id} value={c.id}>{c.nom} ({c.roleLabel})</option>)}
              </select>
            </Field>
            <Field label="Chantier concerne (optionnel)">
              <select value={chantierId} onChange={(e) => setChantierId(e.target.value)} className="w-full rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }}>
                <option value="">Aucun chantier precis</option>
                {mesChantiers.map((c) => <option key={c.id} value={c.id}>{c.nom}</option>)}
              </select>
            </Field>
          </div>
          <Field label="Sujet">
            <input value={sujet} onChange={(e) => setSujet(e.target.value)} placeholder="Ex. Explication sur le retard zone AR-Bureau"
              className="w-full rounded-md px-3 py-2 text-[13px] mb-3" style={{ border: "1px solid " + BORDER }} />
          </Field>
          <Field label="Message">
            <textarea value={texte} onChange={(e) => setTexte(e.target.value)} rows={3} placeholder="Votre message..."
              className="w-full rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
          </Field>
          <button onClick={handleEnvoyer} disabled={!versId || !sujet || !texte}
            className="mt-3 flex items-center gap-1.5 text-white text-[12.5px] font-semibold px-4 py-2.5 rounded-md disabled:opacity-30" style={{ background: ACCENT }}>
            <Send size={14} /> Envoyer (reponse obligatoire pour le destinataire)
          </button>
        </Card>
      )}
      {correspondants.length === 0 && (
        <Card className="p-4 no-print">
          <p className="text-[12.5px]" style={{ color: MUTE }}>Vous ne pouvez pas initier de nouveau message : repondez aux messages recus ci-dessous.</p>
        </Card>
      )}

      <Card className="overflow-hidden no-print">
        <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Boite de reception</h3>
          <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Les messages en attente de reponse apparaissent en premier.</p>
        </div>
        {boiteReception.length === 0 ? (
          <div className="px-6 py-8 text-center text-[13px]" style={{ color: MUTE }}>Aucun message recu.</div>
        ) : (
          <div className="divide-y" style={{ borderColor: BORDER }}>
            {boiteReception.map((m) => {
              const urgent = m.necessiteReponse && !m.repondu;
              return (
                <div key={m.id} className="px-6 py-3.5" style={{ background: urgent ? "#FDF3E2" : m.lu ? "transparent" : "#FDF8F5" }}>
                  <div onClick={() => marquerMessageLu(m.id)} className="flex items-start justify-between gap-3 cursor-pointer">
                    <div className="flex items-start gap-2.5">
                      <div className="w-2 h-2 rounded-full mt-1.5 shrink-0" style={{ background: m.lu ? "transparent" : ACCENT }} />
                      <div>
                        <div className="text-[13px] font-medium flex items-center gap-2" style={{ color: INK }}>
                          {m.sujet}
                          {urgent && <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded-full" style={{ color: BAD, background: "#FBEBE5" }}>Reponse requise</span>}
                          {m.necessiteReponse && m.repondu && <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded-full" style={{ color: GOOD, background: "#E9F7EF" }}>Repondu</span>}
                        </div>
                        <div className="text-[12px] mt-0.5" style={{ color: INK }}>{m.message}</div>
                        <div className="text-[11px] mt-1" style={{ color: MUTE }}>
                          De {nomDe(m.deId)} ({roleLabelDe(m.deId)}){m.chantierId ? " · " + chantierNom(m.chantierId) : ""} &middot; {m.quand}
                        </div>
                      </div>
                    </div>
                    <button onClick={(e) => { e.stopPropagation(); setRepondreA(repondreA === m.id ? null : m.id); }}
                      className="text-[11.5px] font-semibold px-2.5 py-1 rounded-md shrink-0" style={{ color: ACCENT_DEEP, border: "1px dashed " + ACCENT }}>
                      Repondre
                    </button>
                  </div>
                  {repondreA === m.id && (
                    <div className="mt-3 ml-5 flex items-center gap-2">
                      <input value={texteReponse} onChange={(e) => setTexteReponse(e.target.value)} placeholder="Votre reponse..."
                        className="flex-1 rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
                      <button onClick={() => handleRepondre(m)} className="text-[12px] font-semibold text-white px-3 py-2 rounded-md" style={{ background: NAVY }}>Envoyer</button>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </Card>

      <Card className="overflow-hidden">
        <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Tableau recapitulatif &mdash; tracabilite</h3>
          <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>{utilisateur.role === "direction" ? "Tous les echanges de l'entreprise." : "Les echanges vous concernant ou concernant vos chantiers."}</p>
        </div>
        <div className="overflow-x-auto">
        <table className="w-full text-[12.5px]">
          <thead>
            <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
              <th className="px-6 py-2.5 font-semibold">Date</th>
              <th className="px-3 py-2.5 font-semibold">De</th>
              <th className="px-3 py-2.5 font-semibold">Vers</th>
              <th className="px-3 py-2.5 font-semibold">Chantier</th>
              <th className="px-3 py-2.5 font-semibold">Sujet</th>
              <th className="px-3 py-2.5 font-semibold text-right">Statut</th>
            </tr>
          </thead>
          <tbody>
            {tableauTracabilite.map((m, i) => {
              const statut = m.necessiteReponse ? (m.repondu ? { l: "Repondu", c: GOOD, b: "#E9F7EF" } : { l: "Reponse requise", c: BAD, b: "#FBEBE5" }) : (m.lu ? { l: "Lu", c: GOOD, b: "#E9F7EF" } : { l: "Non lu", c: "#5B6472", b: "#F1F2F4" });
              return (
                <tr key={m.id} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                  <td className="px-6 py-2" style={{ color: MUTE }}>{m.quand}</td>
                  <td className="px-3 py-2" style={{ color: INK }}>{nomDe(m.deId)}</td>
                  <td className="px-3 py-2" style={{ color: INK }}>{nomDe(m.versId)}</td>
                  <td className="px-3 py-2" style={{ color: MUTE }}>{chantierNom(m.chantierId)}</td>
                  <td className="px-3 py-2" style={{ color: INK }}>{m.sujet}</td>
                  <td className="px-3 py-2 text-right">
                    <span className="text-[10.5px] font-semibold px-2 py-0.5 rounded-full" style={{ background: statut.b, color: statut.c }}>{statut.l}</span>
                  </td>
                </tr>
              );
            })}
            {tableauTracabilite.length === 0 && (
              <tr><td colSpan={6} className="px-6 py-6 text-center" style={{ color: MUTE }}>Aucun echange enregistre.</td></tr>
            )}
          </tbody>
        </table>
        </div>
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Module 20 — Comptabilite : paiements, remuneration, charges, analytique.
// RAPPEL : les taux/coefficients affiches ici sont des valeurs de depart
// modifiables, pas des regles legales — la paie reelle (bulletins, DSN,
// charges sociales exactes) reste du ressort de l'expert-comptable de SLK Clim.
function coutReelChantier(chantierId, heuresListe, remunerations, utilisateursSysteme, tauxDefaut) {
  const heuresValidees = heuresListe.filter((h) => h.chantierId === chantierId && h.statut === "validee");
  let total = 0;
  heuresValidees.forEach((h) => {
    const user = utilisateursSysteme.find((u) => u.nom === h.ouvrier);
    const remu = user ? remunerations.find((r) => r.userId === user.id) : null;
    const taux = remu && remu.tauxHoraireBrut ? remu.tauxHoraireBrut : tauxDefaut;
    const coeff = remu ? remu.coeffCharges : 42;
    total += h.heures * taux * (1 + coeff / 100);
  });
  return total;
}

function ComptabiliteTab({ chantiers, clients, factures, paiements, enregistrerPaiement, utilisateursSysteme, remunerations, definirRemuneration, charges, ajouterCharge, supprimerCharge, heuresListe, pctDefaut, tauxDefaut, sousOngletCible }) {
  const [subTab, setSubTab] = useState(sousOngletCible || "factures");
  useEffect(() => {
    if (sousOngletCible) setSubTab(sousOngletCible);
  }, [sousOngletCible]);
  const chantierNom = (id) => (chantiers.find((c) => c.id === id) || {}).nom || "-";

  const [paiementPour, setPaiementPour] = useState(null);
  const [montantPaiement, setMontantPaiement] = useState("");
  const [modePaiement, setModePaiement] = useState("virement");

  const [formCharge, setFormCharge] = useState({ libelle: "", categorie: "", montantMensuel: "" });

  function confirmerPaiement(factureId) {
    const m = parseFloat(montantPaiement);
    if (!m || m <= 0) return;
    enregistrerPaiement(factureId, m, modePaiement);
    setPaiementPour(null);
    setMontantPaiement("");
  }
  function handleAjouterCharge() {
    if (!formCharge.libelle || !formCharge.montantMensuel) return;
    ajouterCharge({ libelle: formCharge.libelle, categorie: formCharge.categorie, montantMensuel: parseFloat(formCharge.montantMensuel) || 0 });
    setFormCharge({ libelle: "", categorie: "", montantMensuel: "" });
  }

  const totalChargeMensuelle = charges.reduce((s, c) => s + c.montantMensuel, 0);
  const chantiersActuels = chantiers.filter((c) => chantierAppartientAnneeEnCours(c));
  const chargeParChantier = chantiersActuels.length > 0 ? totalChargeMensuelle / chantiersActuels.length : 0;

  const analytique = chantiersActuels.map((c) => {
    const coutMO = coutReelChantier(c.id, heuresListe, remunerations, utilisateursSysteme, tauxDefaut);
    const montantFourniture = (c.montantHT * pctDefaut.fourniture) / 100;
    const margeReelle = c.montantHT - montantFourniture - coutMO - chargeParChantier;
    return { ...c, coutMO, montantFourniture, margeReelle };
  });

  const subTabs = [
    { id: "factures", label: "Factures & paiements", icon: Receipt },
    { id: "remuneration", label: "Remuneration du personnel", icon: Wallet },
    { id: "charges", label: "Charges de l'entreprise", icon: Building2 },
    { id: "analytique", label: "Comptabilite analytique", icon: TrendingUp },
    { id: "bilan", label: "Bilan annuel", icon: LayoutDashboard },
  ];

  const anneesDisponibles = Object.keys(seedBilanMensuel).map(Number).sort((a, b) => b - a);
  const [anneeChoisie, setAnneeChoisie] = useState(anneesDisponibles[0]);
  const anneeN = seedBilanMensuel[anneeChoisie] || [];
  const anneeN1 = seedBilanMensuel[anneeChoisie - 1] || [];
  const sommeAnnee = (arr, cle) => arr.reduce((s, m) => s + (m[cle] || 0), 0);
  const caN = sommeAnnee(anneeN, "ca");
  const coutN = sommeAnnee(anneeN, "cout");
  const beneficeN = caN - coutN;
  const caN1 = sommeAnnee(anneeN1, "ca");
  const beneficeN1 = caN1 - sommeAnnee(anneeN1, "cout");
  const evolutionCA = caN1 > 0 ? ((caN - caN1) / caN1) * 100 : null;
  const evolutionBenefice = beneficeN1 > 0 ? ((beneficeN - beneficeN1) / beneficeN1) * 100 : null;

  return (
    <div className="space-y-5">
      <div className="text-[11.5px] px-4 py-2.5 rounded-md no-print" style={{ background: "#FDF3E2", color: "#8A5A00" }}>
        Valeurs indicatives (coefficient de charges, taux) — a confirmer avec votre expert-comptable avant tout usage reel. Ce module ne remplace pas un logiciel de paie.
      </div>

      <div className="flex gap-2 flex-wrap no-print">
        {subTabs.map(({ id, label, icon: Icon }) => (
          <button key={id} onClick={() => setSubTab(id)} className="flex items-center gap-1.5 text-[12.5px] font-semibold px-3.5 py-2 rounded-md"
            style={subTab === id ? { background: NAVY, color: "#fff" } : { border: "1px solid " + BORDER, color: MUTE }}>
            <Icon size={14} /> {label}
          </button>
        ))}
      </div>

      {subTab === "factures" && (
        <Card className="overflow-hidden">
          <div className="overflow-x-auto">
          <table className="w-full text-[13px]">
            <thead>
              <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
                <th className="px-6 py-2.5 font-semibold">Numero</th>
                <th className="px-3 py-2.5 font-semibold">Chantier</th>
                <th className="px-3 py-2.5 font-semibold">Type</th>
                <th className="px-3 py-2.5 font-semibold">Montant TTC</th>
                <th className="px-3 py-2.5 font-semibold">Paye</th>
                <th className="px-3 py-2.5 font-semibold">Echeance</th>
                <th className="px-3 py-2.5 font-semibold">Statut</th>
                <th className="px-3 py-2.5"></th>
              </tr>
            </thead>
            <tbody>
              {factures.map((f, i) => {
                const totalDu = f.montantHT + f.montantTVA;
                const dejaPaye = paiements.filter((p) => p.factureId === f.id).reduce((s, p) => s + p.montant, 0);
                const statutInfo = { emise: { l: "Emise", c: "#5B6472", b: "#F1F2F4" }, payee_partiellement: { l: "Partiel", c: "#B5710A", b: "#FDF3E2" }, payee: { l: "Payee", c: GOOD, b: "#E9F7EF" }, en_retard: { l: "En retard", c: BAD, b: "#FBEBE5" } }[f.statutPaiement];
                return (
                  <React.Fragment key={f.id}>
                    <tr style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                      <td className="px-6 py-2.5" style={{ color: INK }}>{f.numero}</td>
                      <td className="px-3 py-2.5" style={{ color: MUTE }}>{chantierNom(f.chantierId)}</td>
                      <td className="px-3 py-2.5 capitalize" style={{ color: MUTE }}>{f.type}</td>
                      <td className="num px-3 py-2.5" style={{ color: INK }}>{fmtEUR(totalDu)}</td>
                      <td className="num px-3 py-2.5" style={{ color: MUTE }}>{fmtEUR(dejaPaye)}</td>
                      <td className="px-3 py-2.5" style={{ color: MUTE }}>{f.echeanceLe}</td>
                      <td className="px-3 py-2.5"><span className="text-[10.5px] font-semibold px-2 py-0.5 rounded-full" style={{ color: statutInfo.c, background: statutInfo.b }}>{statutInfo.l}</span></td>
                      <td className="px-3 py-2.5 text-right no-print">
                        {f.statutPaiement !== "payee" && (
                          <button onClick={() => setPaiementPour(paiementPour === f.id ? null : f.id)} className="text-[11.5px] font-semibold" style={{ color: ACCENT_DEEP }}>
                            + Paiement
                          </button>
                        )}
                      </td>
                    </tr>
                    {paiementPour === f.id && (
                      <tr>
                        <td colSpan={8} className="px-6 py-3 no-print" style={{ background: BG }}>
                          <div className="flex items-center gap-2">
                            <input type="number" value={montantPaiement} onChange={(e) => setMontantPaiement(e.target.value)} placeholder="Montant EUR"
                              className="num rounded-md px-2.5 py-1.5 text-[12.5px] w-32" style={{ border: "1px solid " + BORDER }} />
                            <select value={modePaiement} onChange={(e) => setModePaiement(e.target.value)} className="rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }}>
                              <option value="virement">Virement</option>
                              <option value="cheque">Cheque</option>
                              <option value="especes">Especes</option>
                              <option value="prelevement">Prelevement</option>
                            </select>
                            <button onClick={() => confirmerPaiement(f.id)} className="text-[12px] font-semibold text-white px-3 py-1.5 rounded-md" style={{ background: NAVY }}>Enregistrer</button>
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
          </div>
        </Card>
      )}

      {subTab === "remuneration" && (
        <Card className="overflow-hidden">
          <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
            <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Mode de remuneration par agent</h3>
            <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Salaire fixe, taux horaire, ou forfait par chantier (defini directement sur le chantier concerne).</p>
          </div>
          <div className="divide-y" style={{ borderColor: BORDER }}>
            {utilisateursSysteme.filter((u) => u.role !== "direction").map((u) => {
              const remu = remunerations.find((r) => r.userId === u.id) || { mode: "taux_horaire", salaireMensuelBrut: "", tauxHoraireBrut: "", coeffCharges: 42 };
              return (
                <div key={u.id} className="px-6 py-3.5">
                  <div className="flex items-center justify-between gap-3 flex-wrap">
                    <div className="flex items-center gap-2.5">
                      <div className="w-8 h-8 rounded-full flex items-center justify-center text-[11.5px] font-semibold text-white shrink-0" style={{ background: NAVY }}>{u.initiales}</div>
                      <div>
                        <div className="text-[13px]" style={{ color: INK }}>{u.nom}</div>
                        <div className="text-[11px]" style={{ color: MUTE }}>{u.poste}</div>
                      </div>
                    </div>
                    <div className="flex items-center gap-2 flex-wrap no-print">
                      <select value={remu.mode} onChange={(e) => definirRemuneration(u.id, { ...remu, mode: e.target.value })}
                        className="text-[12px] rounded-md px-2 py-1.5" style={{ border: "1px solid " + BORDER }}>
                        <option value="salaire_fixe">Salaire fixe</option>
                        <option value="taux_horaire">Taux horaire</option>
                        <option value="forfait_projet">Forfait par chantier</option>
                      </select>
                      {remu.mode === "salaire_fixe" && (
                        <input type="number" value={remu.salaireMensuelBrut || ""} onChange={(e) => definirRemuneration(u.id, { ...remu, salaireMensuelBrut: parseFloat(e.target.value) || 0 })}
                          placeholder="EUR brut / mois" className="num w-32 rounded-md px-2 py-1.5 text-[12px]" style={{ border: "1px solid " + BORDER }} />
                      )}
                      {remu.mode === "taux_horaire" && (
                        <input type="number" value={remu.tauxHoraireBrut || ""} onChange={(e) => definirRemuneration(u.id, { ...remu, tauxHoraireBrut: parseFloat(e.target.value) || 0 })}
                          placeholder="EUR brut / h" className="num w-28 rounded-md px-2 py-1.5 text-[12px]" style={{ border: "1px solid " + BORDER }} />
                      )}
                      <input type="number" value={remu.coeffCharges} onChange={(e) => definirRemuneration(u.id, { ...remu, coeffCharges: parseFloat(e.target.value) || 0 })}
                        title="Coefficient de charges patronales (%)" className="num w-20 rounded-md px-2 py-1.5 text-[12px]" style={{ border: "1px solid " + BORDER }} />
                      <span className="text-[11px]" style={{ color: MUTE }}>% charges</span>
                    </div>
                    <div className="print-only text-[12.5px]" style={{ color: INK }}>
                      {MODE_REMUNERATION_LABEL[remu.mode]} &middot; {remu.mode === "salaire_fixe" ? fmtEUR(remu.salaireMensuelBrut || 0) + "/mois" : remu.mode === "taux_horaire" ? (remu.tauxHoraireBrut || 0) + " EUR/h" : "forfait par chantier"} &middot; {remu.coeffCharges}% charges
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
      )}

      {subTab === "charges" && (
        <>
          <Card className="p-6 no-print">
            <h3 className="text-[13px] font-semibold mb-4" style={{ color: INK }}>Ajouter une charge fixe</h3>
            <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
              <input value={formCharge.libelle} onChange={(e) => setFormCharge({ ...formCharge, libelle: e.target.value })} placeholder="Libelle"
                className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              <input value={formCharge.categorie} onChange={(e) => setFormCharge({ ...formCharge, categorie: e.target.value })} placeholder="Categorie"
                className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              <input type="number" value={formCharge.montantMensuel} onChange={(e) => setFormCharge({ ...formCharge, montantMensuel: e.target.value })} placeholder="EUR / mois"
                className="num rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }} />
              <button onClick={handleAjouterCharge} className="flex items-center justify-center gap-1.5 text-[12px] font-semibold text-white px-3 py-2 rounded-md" style={{ background: ACCENT }}>
                <Plus size={14} /> Ajouter
              </button>
            </div>
          </Card>
          <Card className="overflow-hidden">
            <div className="overflow-x-auto">
            <table className="w-full text-[13px]">
              <thead>
                <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
                  <th className="px-6 py-2.5 font-semibold">Libelle</th>
                  <th className="px-3 py-2.5 font-semibold">Categorie</th>
                  <th className="px-3 py-2.5 font-semibold text-right">EUR / mois</th>
                  <th className="px-3 py-2.5"></th>
                </tr>
              </thead>
              <tbody>
                {charges.map((c, i) => (
                  <tr key={c.id} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                    <td className="px-6 py-2.5" style={{ color: INK }}>{c.libelle}</td>
                    <td className="px-3 py-2.5" style={{ color: MUTE }}>{c.categorie}</td>
                    <td className="num px-3 py-2.5 text-right" style={{ color: INK }}>{fmtEUR(c.montantMensuel)}</td>
                    <td className="px-3 py-2.5 text-right no-print"><button onClick={() => supprimerCharge(c.id)} style={{ color: BAD }}><Trash2 size={14} /></button></td>
                  </tr>
                ))}
                <tr>
                  <td className="px-6 py-2.5 font-semibold" style={{ color: INK }} colSpan={2}>Total charges fixes</td>
                  <td className="num px-3 py-2.5 text-right font-semibold" style={{ color: INK }}>{fmtEUR(totalChargeMensuelle)}</td>
                  <td></td>
                </tr>
              </tbody>
            </table>
            </div>
          </Card>
        </>
      )}

      {subTab === "analytique" && (
        <Card className="overflow-hidden">
          <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
            <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Marge reelle par chantier</h3>
            <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Montant devis - fournitures - main-d'oeuvre reelle chargee - quote-part de charges fixes ({fmtEUR(chargeParChantier)}/chantier).</p>
          </div>
          <div className="overflow-x-auto">
          <table className="w-full text-[13px]">
            <thead>
              <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
                <th className="px-6 py-2.5 font-semibold">Chantier</th>
                <th className="px-3 py-2.5 font-semibold">Devis HT</th>
                <th className="px-3 py-2.5 font-semibold">MO reelle chargee</th>
                <th className="px-3 py-2.5 font-semibold">Charges allouees</th>
                <th className="px-3 py-2.5 font-semibold text-right">Marge reelle</th>
              </tr>
            </thead>
            <tbody>
              {analytique.map((c, i) => (
                <tr key={c.id} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                  <td className="px-6 py-2.5" style={{ color: INK }}>{c.nom}</td>
                  <td className="num px-3 py-2.5" style={{ color: MUTE }}>{fmtEUR(c.montantHT)}</td>
                  <td className="num px-3 py-2.5" style={{ color: MUTE }}>{fmtEUR(c.coutMO)}</td>
                  <td className="num px-3 py-2.5" style={{ color: MUTE }}>{fmtEUR(chargeParChantier)}</td>
                  <td className="num px-3 py-2.5 text-right font-semibold" style={{ color: c.margeReelle < 0 ? BAD : GOOD }}>{fmtEUR(c.margeReelle)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          </div>
        </Card>
      )}

      {subTab === "bilan" && (
        <>
          <div className="flex items-center justify-between no-print">
            <Field label="Exercice">
              <select value={anneeChoisie} onChange={(e) => setAnneeChoisie(Number(e.target.value))} className="rounded-md px-3 py-2 text-[13px]" style={{ border: "1px solid " + BORDER }}>
                {anneesDisponibles.map((a) => <option key={a} value={a}>{a}</option>)}
              </select>
            </Field>
            <BoutonImprimer label="Imprimer le bilan" />
          </div>

          <div className="print-only">
            <h1 className="text-[18px] font-bold" style={{ color: INK }}>Bilan de l'exercice {anneeChoisie}</h1>
            <p className="text-[12px]" style={{ color: MUTE }}>Edite le {new Date().toLocaleDateString("fr-FR")}</p>
          </div>

          <div className="text-[11px] px-4 py-2 rounded-md no-print" style={{ background: "#F1F2F4", color: MUTE }}>
            Repartition mensuelle illustrative, a brancher sur les vraies dates de facturation une fois le logiciel connecte a la base reelle.
          </div>

          <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
            <KPI icon={Wallet} label={"CA exercice " + anneeChoisie} value={fmtEUR(caN)} color="#1F7A4D" />
            <KPI icon={Receipt} label="Cout de revient exercice" value={fmtEUR(coutN)} color="#B5710A" />
            <KPI icon={TrendingUp} label="Benefice exercice" value={fmtEUR(beneficeN)} alert={beneficeN < 0} color="#7C3AED" />
            <KPI icon={TrendingUp}
              label={"Evolution CA vs " + (anneeChoisie - 1)}
              value={evolutionCA === null ? "-" : (evolutionCA >= 0 ? "+" : "") + evolutionCA.toFixed(1) + "%"}
              alert={evolutionCA !== null && evolutionCA < 0}
              color="#2563EB"
            />
          </div>

          <Card className="overflow-hidden">
            <div className="px-6 py-4" style={{ borderBottom: "1px solid " + BORDER }}>
              <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Comparatif mensuel {anneeChoisie} vs {anneeChoisie - 1}</h3>
            </div>
            <div className="overflow-x-auto">
            <table className="w-full text-[13px]">
              <thead>
                <tr className="text-left text-[10.5px] uppercase tracking-wide" style={{ color: MUTE, borderBottom: "1px solid " + BORDER }}>
                  <th className="px-6 py-2.5 font-semibold">Mois</th>
                  <th className="px-3 py-2.5 font-semibold">CA {anneeChoisie}</th>
                  <th className="px-3 py-2.5 font-semibold">CA {anneeChoisie - 1}</th>
                  <th className="px-3 py-2.5 font-semibold">Evolution</th>
                  <th className="px-3 py-2.5 font-semibold text-right">Benefice {anneeChoisie}</th>
                </tr>
              </thead>
              <tbody>
                {MOIS_LABELS.map((label, i) => {
                  const mN = anneeN[i] || {};
                  const mN1 = anneeN1[i] || {};
                  const evo = mN1.ca ? ((mN.ca - mN1.ca) / mN1.ca) * 100 : null;
                  const beneficeMois = mN.ca != null && mN.cout != null ? mN.ca - mN.cout : null;
                  return (
                    <tr key={label} style={{ background: i % 2 ? "#FAFBFC" : "transparent" }}>
                      <td className="px-6 py-2" style={{ color: INK }}>{label}</td>
                      <td className="num px-3 py-2" style={{ color: mN.ca == null ? MUTE : INK }}>{mN.ca == null ? "-" : fmtEUR(mN.ca)}</td>
                      <td className="num px-3 py-2" style={{ color: MUTE }}>{mN1.ca == null ? "-" : fmtEUR(mN1.ca)}</td>
                      <td className="num px-3 py-2" style={{ color: evo == null ? MUTE : evo >= 0 ? GOOD : BAD }}>
                        {evo == null ? "-" : (evo >= 0 ? "+" : "") + evo.toFixed(1) + "%"}
                      </td>
                      <td className="num px-3 py-2 text-right font-semibold" style={{ color: beneficeMois == null ? MUTE : beneficeMois < 0 ? BAD : INK }}>
                        {beneficeMois == null ? "-" : fmtEUR(beneficeMois)}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
function ParametresTab({ pctDefaut, setPctDefaut, tauxDefaut, setTauxDefaut, seuilAlerte, setSeuilAlerte, coeffConsommables, setCoeffConsommables, remiseDefaut, setRemiseDefaut, utilisateursSysteme, setUtilisateursSysteme, utilisateur, remunerations, definirRemuneration }) {
  const [nouvUser, setNouvUser] = useState({ nom: "", poste: "", email: "", role: "ouvrier", modeRemuneration: "taux_horaire", montantRemuneration: "" });

  function initialesDe(nom) {
    return nom.split(" ").filter(Boolean).slice(0, 2).map((m) => m[0].toUpperCase()).join("");
  }
  function ajouterUtilisateur() {
    if (!nouvUser.nom || !nouvUser.poste || !nouvUser.email) return;
    const roleLabel = { direction: "Direction", chef_chantier: "Chef de chantier", ouvrier: "Ouvrier" }[nouvUser.role];
    const id = "u" + Date.now();
    setUtilisateursSysteme((prev) => [...prev, { id, nom: nouvUser.nom, poste: nouvUser.poste, email: nouvUser.email, role: nouvUser.role, roleLabel, initiales: initialesDe(nouvUser.nom) }]);
    // Mode de remuneration choisi directement a la creation (salaire fixe,
    // taux horaire, ou paiement par chantier/forfait) — modifiable ensuite
    // dans Comptabilite si besoin.
    if (nouvUser.role !== "direction") {
      const montant = parseFloat(nouvUser.montantRemuneration) || 0;
      definirRemuneration(id, {
        mode: nouvUser.modeRemuneration,
        salaireMensuelBrut: nouvUser.modeRemuneration === "salaire_fixe" ? montant : null,
        tauxHoraireBrut: nouvUser.modeRemuneration === "taux_horaire" ? montant : null,
        coeffCharges: 42,
      });
    }
    setNouvUser({ nom: "", poste: "", email: "", role: "ouvrier", modeRemuneration: "taux_horaire", montantRemuneration: "" });
  }
  function changerRole(id, role) {
    const roleLabel = { direction: "Direction", chef_chantier: "Chef de chantier", ouvrier: "Ouvrier" }[role];
    setUtilisateursSysteme((prev) => prev.map((u) => (u.id === id ? { ...u, role, roleLabel } : u)));
  }
  function supprimerUtilisateur(id) {
    if (id === utilisateur.id) return; // on ne se supprime pas soi-meme
    setUtilisateursSysteme((prev) => prev.filter((u) => u.id !== id));
  }
  const sommePct = pctDefaut.fourniture + pctDefaut.frais + pctDefaut.mainOeuvre;
  return (
    <div className="space-y-5 max-w-2xl">
      <Card className="overflow-hidden">
        <div className="px-6 py-5" style={{ borderBottom: "1px solid " + BORDER }}>
          <h3 className="text-[13.5px] font-semibold" style={{ color: INK }}>Utilisateurs et droits d'acces</h3>
          <p className="text-[12px] mt-0.5" style={{ color: MUTE }}>Modules 3/5/20 &mdash; seule la Direction cree les comptes ; le mode de remuneration se choisit des la creation.</p>
        </div>
        <div className="divide-y" style={{ borderColor: BORDER }}>
          {utilisateursSysteme.map((u) => {
            const remu = remunerations.find((r) => r.userId === u.id);
            return (
            <div key={u.id} className="px-6 py-2.5 flex items-center justify-between gap-3 flex-wrap">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-full flex items-center justify-center text-[11.5px] font-semibold text-white shrink-0" style={{ background: NAVY }}>{u.initiales}</div>
                <div>
                  <div className="text-[13px]" style={{ color: INK }}>{u.nom} {u.id === utilisateur.id && <span className="text-[10.5px]" style={{ color: MUTE }}>(vous)</span>}</div>
                  <div className="text-[11px]" style={{ color: MUTE }}>
                    {u.poste} &middot; {u.email || "email non renseigne"}
                    {remu && <span> &middot; {MODE_REMUNERATION_LABEL[remu.mode]}{remu.mode === "salaire_fixe" && remu.salaireMensuelBrut ? " (" + fmtEUR(remu.salaireMensuelBrut) + "/mois)" : ""}{remu.mode === "taux_horaire" && remu.tauxHoraireBrut ? " (" + remu.tauxHoraireBrut + " EUR/h)" : ""}</span>}
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <select value={u.role} onChange={(e) => changerRole(u.id, e.target.value)}
                  className="text-[11.5px] rounded-md px-2 py-1" style={{ border: "1px solid " + BORDER }}>
                  <option value="direction">Direction</option>
                  <option value="chef_chantier">Chef de chantier</option>
                  <option value="ouvrier">Ouvrier</option>
                </select>
                <button onClick={() => supprimerUtilisateur(u.id)} disabled={u.id === utilisateur.id} style={{ color: u.id === utilisateur.id ? "#CBD0D8" : BAD }}>
                  <Trash2 size={14} />
                </button>
              </div>
            </div>
          );})}
        </div>
        <div className="px-6 py-4" style={{ background: BG, borderTop: "1px solid " + BORDER }}>
          <div className="grid grid-cols-1 sm:grid-cols-4 gap-2 mb-2">
            <input value={nouvUser.nom} onChange={(e) => setNouvUser({ ...nouvUser, nom: e.target.value })} placeholder="Nom complet"
              className="rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }} />
            <input value={nouvUser.poste} onChange={(e) => setNouvUser({ ...nouvUser, poste: e.target.value })} placeholder="Poste"
              className="rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }} />
            <input type="email" value={nouvUser.email} onChange={(e) => setNouvUser({ ...nouvUser, email: e.target.value })} placeholder="Email personnel (code de connexion)"
              className="rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }} />
            <select value={nouvUser.role} onChange={(e) => setNouvUser({ ...nouvUser, role: e.target.value })}
              className="rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }}>
              <option value="direction">Direction</option>
              <option value="chef_chantier">Chef de chantier</option>
              <option value="ouvrier">Ouvrier</option>
            </select>
          </div>
          {nouvUser.role !== "direction" && (
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 mb-2">
              <select value={nouvUser.modeRemuneration} onChange={(e) => setNouvUser({ ...nouvUser, modeRemuneration: e.target.value, montantRemuneration: "" })}
                className="rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }}>
                <option value="salaire_fixe">Salaire fixe</option>
                <option value="taux_horaire">Taux horaire</option>
                <option value="forfait_projet">Paye par chantier (forfait)</option>
              </select>
              {nouvUser.modeRemuneration === "salaire_fixe" && (
                <input type="number" value={nouvUser.montantRemuneration} onChange={(e) => setNouvUser({ ...nouvUser, montantRemuneration: e.target.value })}
                  placeholder="EUR brut / mois" className="num rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }} />
              )}
              {nouvUser.modeRemuneration === "taux_horaire" && (
                <input type="number" value={nouvUser.montantRemuneration} onChange={(e) => setNouvUser({ ...nouvUser, montantRemuneration: e.target.value })}
                  placeholder="EUR brut / heure" className="num rounded-md px-2.5 py-1.5 text-[12.5px]" style={{ border: "1px solid " + BORDER }} />
              )}
              {nouvUser.modeRemuneration === "forfait_projet" && (
                <div className="text-[11px] flex items-center" style={{ color: MUTE }}>Le montant du forfait se definit chantier par chantier, dans Comptabilite.</div>
              )}
            </div>
          )}
          <button onClick={ajouterUtilisateur} className="flex items-center justify-center gap-1.5 text-[12px] font-semibold text-white px-3 py-1.5 rounded-md" style={{ background: ACCENT }}>
            <UserPlus size={13} /> Creer le compte
          </button>
          <p className="text-[10.5px] mt-2" style={{ color: MUTE }}>Le mot de passe initial est communique separement par la Direction (non gere dans ce prototype visuel). Le mode de remuneration reste modifiable ensuite dans Comptabilite.</p>
        </div>
      </Card>

      <Card className="p-6">
        <h3 className="text-[13.5px] font-semibold mb-1" style={{ color: INK }}>Repartition par defaut d'un devis</h3>
        <p className="text-[12px] mb-4" style={{ color: MUTE }}>Module 1 &mdash; utilisee pour chaque nouveau devis, modifiable ensuite au cas par cas.</p>
        <div className="grid grid-cols-3 gap-3">
          <Field label="Fournitures %">
            <input type="number" value={pctDefaut.fourniture} onChange={(e) => setPctDefaut({ ...pctDefaut, fourniture: parseFloat(e.target.value) || 0 })}
              className="num w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
          </Field>
          <Field label="Frais entreprise %">
            <input type="number" value={pctDefaut.frais} onChange={(e) => setPctDefaut({ ...pctDefaut, frais: parseFloat(e.target.value) || 0 })}
              className="num w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
          </Field>
          <Field label="Main-d'oeuvre %">
            <input type="number" value={pctDefaut.mainOeuvre} onChange={(e) => setPctDefaut({ ...pctDefaut, mainOeuvre: parseFloat(e.target.value) || 0 })}
              className="num w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
          </Field>
        </div>
        {sommePct !== 100 && (
          <div className="text-[11px] mt-2.5 flex items-center gap-1" style={{ color: BAD }}>
            <AlertTriangle size={12} /> La somme doit faire 100% (actuellement {sommePct}%).
          </div>
        )}
      </Card>

      <Card className="p-6">
        <h3 className="text-[13.5px] font-semibold mb-1" style={{ color: INK }}>Taux horaire moyen par defaut</h3>
        <p className="text-[12px] mb-4" style={{ color: MUTE }}>Module 3 &mdash; utilise pour convertir la main-d'oeuvre en heures prevues quand aucun taux specifique n'est saisi.</p>
        <div className="max-w-[200px]">
          <Field label="EUR / heure">
            <input type="number" value={tauxDefaut} onChange={(e) => setTauxDefaut(parseFloat(e.target.value) || 0)}
              className="num w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
          </Field>
        </div>
      </Card>

      <Card className="p-6">
        <h3 className="text-[13.5px] font-semibold mb-1" style={{ color: INK }}>Remise globale par defaut</h3>
        <p className="text-[12px] mb-4" style={{ color: MUTE }}>Module 1 &mdash; remise nego appliquee au total HT d'un devis (ex. 4,10% observe sur le chiffrage CPSSD Bat QF). Modifiable au cas par cas dans chaque devis.</p>
        <div className="max-w-[200px]">
          <Field label="% de remise">
            <input type="number" step="0.1" value={remiseDefaut} onChange={(e) => setRemiseDefaut(parseFloat(e.target.value) || 0)}
              className="num w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
          </Field>
        </div>
      </Card>

      <Card className="p-6">
        <h3 className="text-[13.5px] font-semibold mb-1" style={{ color: INK }}>Seuil d'alerte tableau de bord</h3>
        <p className="text-[12px] mb-4" style={{ color: MUTE }}>Module 8 &mdash; pourcentage d'heures consommees a partir duquel un chantier est signale en alerte.</p>
        <div className="max-w-[200px]">
          <Field label="% des heures prevues">
            <input type="number" value={seuilAlerte} onChange={(e) => setSeuilAlerte(parseFloat(e.target.value) || 0)}
              className="num w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
          </Field>
        </div>
      </Card>

      <Card className="p-6">
        <h3 className="text-[13.5px] font-semibold mb-1" style={{ color: INK }}>Coefficient de consommables</h3>
        <p className="text-[12px] mb-4" style={{ color: MUTE }}>Section 6.6 du CDC &mdash; methode alternative pour les petites fournitures sur petits chantiers, en % du montant de la pose.</p>
        <div className="max-w-[200px]">
          <Field label="% du montant de la pose">
            <input type="number" step="0.1" value={coeffConsommables} onChange={(e) => setCoeffConsommables(parseFloat(e.target.value) || 0)}
              className="num w-full rounded-md px-3 py-2 text-[13.5px]" style={{ border: "1px solid " + BORDER }} />
          </Field>
        </div>
      </Card>

      <div className="text-[11.5px] px-1" style={{ color: MUTE }}>
        Conformement au cahier des charges (section 6.1bis), aucune de ces valeurs n'est ecrite en dur dans le code du logiciel : elles vivent toutes ici, modifiables par la Direction a tout moment.
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
function Footer() {
  return (
    <div className="px-8 py-4 text-center text-[11.5px] no-print" style={{ borderTop: "1px solid " + BORDER, color: MUTE, background: SURFACE }}>
      Developpe par <span style={{ color: INK, fontWeight: 600 }}>www.informaint.com</span>
      {" "}&middot;{" "}
      <span style={{ color: INK, fontWeight: 600 }}>contact@informaint.com</span>
      <br className="sm:hidden" />
      <span className="sm:ml-1">Copyright 2026. INFORMAINT &mdash; www.informaint.com</span>
    </div>
  );
}
