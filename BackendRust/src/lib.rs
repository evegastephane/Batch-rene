//! Conversion des classeurs Excel de MMC Batch et historique des conversions.
//!
//! Utilise directement par l'appli Tauri (commandes) et par le serveur HTTP
//! optionnel (`src/main.rs`, pour travailler sur l'interface dans un navigateur).

use std::{
    collections::HashMap,
    fmt,
    io::{Cursor, Write},
    path::{Path, PathBuf},
};

use calamine::{open_workbook_auto_from_rs, Data, Reader};
use rusqlite::{params, Connection};
use rust_xlsxwriter::{Workbook, XlsxError};
use serde::{Deserialize, Serialize};
use uuid::Uuid;
use zip::{write::SimpleFileOptions, CompressionMethod, ZipWriter};

// Elements par defaut des lignes AddAlias (mode CRPLMT), modifiables dans le formulaire :
// HDR,"<action>","<titre>","<ext_id>","<nb>","<description>","1"
// <identifiant>,"<numero>","<prefixe><separateur><numero><symbole><montant>"
pub const ADD_ALIAS_ACTION_DEFAUT: &str = "AddAlias";
pub const ADD_ALIAS_TITRE_DEFAUT: &str = "Add Alias Title";
pub const ADD_ALIAS_EXT_ID_DEFAUT: &str = "ExtId12340";
pub const ADD_ALIAS_DESCRIPTION_DEFAUT: &str = "Agent  advance pilote";
pub const CRPLMT_IDENTIFIANT_DEFAUT: &str = "MSISDN";
pub const CRPLMT_PREFIXE_DEFAUT: &str = "CRPLMT";
pub const CRPLMT_SEPARATEUR_DEFAUT: &str = "_";
pub const CRPLMT_SYMBOLE_DEFAUT: &str = "@";
pub const CRPLMT_MONTANT_DEFAUT: &str = "100000";

// ---------------------------------------------------------------------------
// Erreurs
// ---------------------------------------------------------------------------

#[derive(Debug)]
pub enum Erreur {
    /// Donnees envoyees invalides (fichier, montant...) : a corriger par l'utilisateur.
    Invalide(String),
    /// Entree d'historique ou fichier introuvable.
    Introuvable(String),
    /// Probleme inattendu (disque, base de donnees...).
    Interne(String),
}

impl fmt::Display for Erreur {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Erreur::Invalide(m) | Erreur::Introuvable(m) | Erreur::Interne(m) => f.write_str(m),
        }
    }
}

impl std::error::Error for Erreur {}

impl From<rusqlite::Error> for Erreur {
    fn from(e: rusqlite::Error) -> Self {
        Erreur::Interne(format!("Base de donnees : {e}"))
    }
}

impl From<std::io::Error> for Erreur {
    fn from(e: std::io::Error) -> Self {
        Erreur::Interne(format!("Fichier : {e}"))
    }
}

impl From<zip::result::ZipError> for Erreur {
    fn from(e: zip::result::ZipError) -> Self {
        Erreur::Interne(format!("Creation du zip : {e}"))
    }
}

impl From<XlsxError> for Erreur {
    fn from(e: XlsxError) -> Self {
        Erreur::Interne(format!("Creation du classeur Excel : {e}"))
    }
}

pub type Resultat<T> = std::result::Result<T, Erreur>;

// ---------------------------------------------------------------------------
// Demande et resultat de conversion
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Mode {
    /// Classeur deja au format `MSISDN,...,Parametre` : ajoute l'entete
    /// `UpdateIndividualRatingParameter` (Init ou Set), un CSV par feuille.
    #[default]
    Standard,
    /// Colonne de numeros : produit un classeur .xlsx dont la colonne A contient
    /// l'entete `HDR,"AddAlias",...` puis une ligne
    /// `MSISDN,"<numero>","CRPLMT_<numero>@<montant>"` par numero.
    Crplmt,
}

#[derive(Debug, Deserialize)]
pub struct Demande {
    pub nom_fichier: String,
    pub contenu: Vec<u8>,
    #[serde(default)]
    pub mode: Mode,
    /// Mode standard : "Init" (defaut) ou "Set".
    #[serde(default)]
    pub style_entete: Option<String>,
    /// Mode CRPLMT.
    #[serde(flatten)]
    pub elements: ElementsAddAlias,
}

/// Elements des lignes AddAlias saisis dans le formulaire (mode CRPLMT).
/// Un element absent ou vide prend sa valeur par defaut.
#[derive(Debug, Default, Deserialize)]
#[serde(default)]
pub struct ElementsAddAlias {
    // Entete HDR
    pub action: Option<String>,
    pub titre: Option<String>,
    pub ext_id: Option<String>,
    pub description: Option<String>,
    // Ligne de chaque numero
    pub identifiant: Option<String>,
    pub prefixe: Option<String>,
    pub separateur: Option<String>,
    pub symbole: Option<String>,
    pub montant: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct Conversion {
    /// Fichier produit : zip de CSV (Standard) ou classeur .xlsx (CRPLMT).
    pub nom_zip: String,
    pub nb_feuilles: usize,
    /// Lignes de donnees ecrites, entetes HDR non comprises.
    pub nb_lignes: usize,
    pub historique_id: i64,
    pub zip: Vec<u8>,
}

/// Convertit le classeur, enregistre le fichier produit dans l'historique et le renvoie.
pub fn convertir(historique: &Historique, demande: Demande) -> Resultat<Conversion> {
    if !a_extension_excel(&demande.nom_fichier) {
        return Err(Erreur::Invalide(
            "Le fichier doit etre un .xlsx ou .xls".into(),
        ));
    }
    let base = sans_extension(&demande.nom_fichier);

    let (sortie, nom_sortie, nb_feuilles, nb_lignes, type_entete) = match demande.mode {
        Mode::Standard => {
            let style = if demande.style_entete.as_deref() == Some("Set") {
                "Set"
            } else {
                "Init"
            };
            let (fichiers, nb_lignes) = fichiers_standard(&demande.contenu, style)?;
            (
                construire_zip(&fichiers)?,
                format!("{base}_csv.zip"),
                fichiers.len(),
                nb_lignes,
                style.to_string(),
            )
        }
        Mode::Crplmt => {
            let elements = demande.elements.resoudre()?;
            let feuilles = feuilles_de_numeros(&demande.contenu)?;
            let nb_lignes = feuilles.iter().map(|(_, numeros)| numeros.len()).sum();
            (
                classeur_add_alias(&feuilles, &elements)?,
                format!("{base}_AddAlias.xlsx"),
                feuilles.len(),
                nb_lignes,
                "AddAlias".to_string(),
            )
        }
    };

    let historique_id = historique.enregistrer(
        &demande.nom_fichier,
        &nom_sortie,
        &sortie,
        nb_feuilles,
        &type_entete,
    )?;

    Ok(Conversion {
        nom_zip: nom_sortie,
        nb_feuilles,
        nb_lignes,
        historique_id,
        zip: sortie,
    })
}

/// Valeur d'un champ texte, ou la valeur par defaut s'il est absent ou vide.
/// Les espaces internes sont conserves (ex. "Agent  advance pilote").
fn valeur_ou_defaut(valeur: &Option<String>, defaut: &str) -> String {
    match valeur.as_deref() {
        Some(v) if !v.trim().is_empty() => v.trim_matches(['\r', '\n']).to_string(),
        _ => defaut.to_string(),
    }
}

fn a_extension_excel(nom: &str) -> bool {
    let nom = nom.to_ascii_lowercase();
    nom.ends_with(".xlsx") || nom.ends_with(".xls")
}

fn sans_extension(nom: &str) -> &str {
    nom.rsplit_once('.').map(|(base, _)| base).unwrap_or(nom)
}

fn extension(nom: &str) -> &str {
    nom.rsplit_once('.').map(|(_, ext)| ext).unwrap_or("zip")
}

// ---------------------------------------------------------------------------
// Lecture Excel
// ---------------------------------------------------------------------------

type Feuille = (String, Vec<Vec<String>>);

/// Lit toutes les feuilles ; les lignes entierement vides sont ignorees.
fn lire_feuilles(contenu: &[u8]) -> Resultat<Vec<Feuille>> {
    let illisible = |e: &dyn fmt::Display| Erreur::Invalide(format!("Fichier Excel illisible : {e}"));

    let mut classeur =
        open_workbook_auto_from_rs(Cursor::new(contenu.to_vec())).map_err(|e| illisible(&e))?;
    let noms = classeur.sheet_names().to_owned();
    if noms.is_empty() {
        return Err(Erreur::Invalide("Aucune feuille trouvee dans le fichier.".into()));
    }

    let mut feuilles = Vec::new();
    for nom in noms {
        let plage = classeur.worksheet_range(&nom).map_err(|e| illisible(&e))?;
        let lignes = plage
            .rows()
            .map(|ligne| ligne.iter().map(nettoyer_cellule).collect::<Vec<_>>())
            .filter(|ligne| ligne.iter().any(|v| !v.trim().is_empty()))
            .collect();
        feuilles.push((nom, lignes));
    }
    Ok(feuilles)
}

fn nettoyer_cellule(cellule: &Data) -> String {
    let valeur = match cellule {
        Data::Empty => String::new(),
        Data::String(v) => v.clone(),
        Data::Float(v) if v.fract() == 0.0 => format!("{v:.0}"),
        Data::Float(v) => v.to_string(),
        Data::Int(v) => v.to_string(),
        Data::Bool(v) => v.to_string(),
        Data::DateTime(v) => v.to_string(),
        Data::DateTimeIso(v) | Data::DurationIso(v) => v.clone(),
        Data::Error(v) => v.to_string(),
    };
    valeur.replace(['\r', '\n'], " ").trim().to_string()
}

fn nom_de_fichier(nom: &str) -> String {
    nom.replace(['/', '\\', ':', '*', '?', '"', '<', '>', '|'], "-")
}

// ---------------------------------------------------------------------------
// Mode standard : entete UpdateIndividualRatingParameter
// ---------------------------------------------------------------------------

/// Fichiers produits et nombre total de lignes de donnees.
type Fichiers = (Vec<(String, Vec<u8>)>, usize);

fn fichiers_standard(contenu: &[u8], style: &str) -> Resultat<Fichiers> {
    let mut fichiers = Vec::new();
    let mut nb_lignes = 0;
    for (feuille, lignes) in lire_feuilles(contenu)? {
        nb_lignes += lignes.len();
        let parametre = parametre_le_plus_frequent(&lignes).unwrap_or_else(|| feuille.clone());

        let mut csv = entete_standard(lignes.len(), style, &parametre);
        csv.push('\n');
        for ligne in &lignes {
            csv.push_str(&ligne.join(","));
            csv.push('\n');
        }

        // BOM UTF-8, comme les fichiers produits jusqu'ici
        let mut octets = vec![0xEF, 0xBB, 0xBF];
        octets.extend_from_slice(csv.as_bytes());
        fichiers.push((format!("{}.csv", nom_de_fichier(&feuille)), octets));
    }
    Ok((fichiers, nb_lignes))
}

fn entete_standard(nb_lignes: usize, style: &str, parametre: &str) -> String {
    if style == "Set" {
        format!(
            "HDR,UpdateIndividualRatingParameter,Set {parametre},ExtId12341,{nb_lignes},{parametre} enable,1"
        )
    } else {
        format!(
            "HDR,\"UpdateIndividualRatingParameter\",\"Init {parametre} TRUE\",\"ExtId12340\",\"{nb_lignes}\",\" {parametre} init\",\"1\""
        )
    }
}

/// Le parametre est la 3e valeur de chaque ligne (`MSISDN,<numero>,<parametre>,...`),
/// y compris quand toute la ligne est dans une seule cellule.
fn parametre_le_plus_frequent(lignes: &[Vec<String>]) -> Option<String> {
    let mut compte: HashMap<String, usize> = HashMap::new();
    for ligne in lignes {
        let complete = ligne.join(",");
        let parametre = complete
            .split(',')
            .map(|s| s.trim().trim_matches('"').trim())
            .nth(2)
            .filter(|p| !p.is_empty());
        if let Some(p) = parametre {
            *compte.entry(p.to_string()).or_insert(0) += 1;
        }
    }
    // A egalite, ordre alphabetique pour un resultat stable
    compte
        .into_iter()
        .max_by(|a, b| a.1.cmp(&b.1).then_with(|| b.0.cmp(&a.0)))
        .map(|(p, _)| p)
}

// ---------------------------------------------------------------------------
// Mode CRPLMT : classeur AddAlias
// ---------------------------------------------------------------------------

/// Elements AddAlias avec les valeurs par defaut appliquees et controlees.
struct LigneAddAlias {
    action: String,
    titre: String,
    ext_id: String,
    description: String,
    identifiant: String,
    prefixe: String,
    separateur: String,
    symbole: String,
    montant: String,
}

impl ElementsAddAlias {
    fn resoudre(&self) -> Resultat<LigneAddAlias> {
        let e = LigneAddAlias {
            action: valeur_ou_defaut(&self.action, ADD_ALIAS_ACTION_DEFAUT),
            titre: valeur_ou_defaut(&self.titre, ADD_ALIAS_TITRE_DEFAUT),
            ext_id: valeur_ou_defaut(&self.ext_id, ADD_ALIAS_EXT_ID_DEFAUT),
            description: valeur_ou_defaut(&self.description, ADD_ALIAS_DESCRIPTION_DEFAUT),
            identifiant: jeton_ou_defaut(&self.identifiant, CRPLMT_IDENTIFIANT_DEFAUT),
            prefixe: jeton_ou_defaut(&self.prefixe, CRPLMT_PREFIXE_DEFAUT),
            separateur: jeton_ou_defaut(&self.separateur, CRPLMT_SEPARATEUR_DEFAUT),
            symbole: jeton_ou_defaut(&self.symbole, CRPLMT_SYMBOLE_DEFAUT),
            montant: jeton_ou_defaut(&self.montant, CRPLMT_MONTANT_DEFAUT),
        };

        if !e.montant.chars().all(|c| c.is_ascii_digit()) {
            return Err(Erreur::Invalide(
                "Le montant doit etre un nombre entier (ex. 100000).".into(),
            ));
        }
        // Un guillemet fermerait la valeur entre guillemets avant la fin
        for (nom, valeur) in [
            ("Action", &e.action),
            ("Titre", &e.titre),
            ("ExtId", &e.ext_id),
            ("Description", &e.description),
            ("Identifiant", &e.identifiant),
            ("Prefixe", &e.prefixe),
            ("Separateur", &e.separateur),
            ("Symbole", &e.symbole),
        ] {
            if valeur.contains('"') {
                return Err(Erreur::Invalide(format!(
                    "{nom} : les guillemets ne sont pas acceptes."
                )));
            }
        }
        // Seul element hors guillemets : une virgule decalerait les colonnes
        if e.identifiant.contains(',') {
            return Err(Erreur::Invalide(
                "Identifiant : la virgule n'est pas acceptee.".into(),
            ));
        }
        Ok(e)
    }
}

/// Element court de la ligne (MSISDN, CRPLMT, _, @, montant) : sans les espaces
/// autour, ou la valeur par defaut s'il est absent ou vide.
fn jeton_ou_defaut(valeur: &Option<String>, defaut: &str) -> String {
    match valeur.as_deref().map(str::trim) {
        Some(v) if !v.is_empty() => v.to_string(),
        _ => defaut.to_string(),
    }
}

/// Feuilles contenant des numeros, avec leurs numeros.
fn feuilles_de_numeros(contenu: &[u8]) -> Resultat<Vec<(String, Vec<String>)>> {
    let mut feuilles = Vec::new();
    for (feuille, lignes) in lire_feuilles(contenu)? {
        let numeros = extraire_numeros(&lignes)
            .map_err(|e| Erreur::Invalide(format!("Feuille \"{feuille}\" : {e}")))?;
        if !numeros.is_empty() {
            feuilles.push((feuille, numeros));
        }
    }

    if feuilles.is_empty() {
        return Err(Erreur::Invalide("Aucun numero trouve dans le fichier.".into()));
    }
    Ok(feuilles)
}

/// Premiere cellule non vide de chaque ligne. Une premiere ligne non numerique
/// est le titre de la colonne ; toute autre valeur non numerique est refusee.
fn extraire_numeros(lignes: &[Vec<String>]) -> std::result::Result<Vec<String>, String> {
    let mut numeros = Vec::new();
    for (index, ligne) in lignes.iter().enumerate() {
        let Some(valeur) = ligne.iter().map(|v| v.trim()).find(|v| !v.is_empty()) else {
            continue;
        };
        let valeur: String = valeur.chars().filter(|c| !c.is_whitespace()).collect();

        if !valeur.is_empty() && valeur.chars().all(|c| c.is_ascii_digit()) {
            numeros.push(valeur);
        } else if index > 0 {
            return Err(format!(
                "valeur \"{valeur}\" invalide a la ligne {} (un numero est attendu)",
                index + 1
            ));
        }
    }
    Ok(numeros)
}

/// L'entete HDR puis une ligne `MSISDN,"<numero>","CRPLMT_<numero>@<montant>"` par numero.
fn add_alias(numeros: &[String], e: &LigneAddAlias) -> Vec<String> {
    let mut lignes = Vec::with_capacity(numeros.len() + 1);
    lignes.push(format!(
        "HDR,\"{}\",\"{}\",\"{}\",\"{}\",\"{}\",\"1\"",
        e.action,
        e.titre,
        e.ext_id,
        numeros.len(),
        e.description
    ));
    lignes.extend(numeros.iter().map(|numero| {
        format!(
            "{},\"{numero}\",\"{}{}{numero}{}{}\"",
            e.identifiant, e.prefixe, e.separateur, e.symbole, e.montant
        )
    }));
    lignes
}

/// Classeur de sortie : une feuille par feuille source, une ligne AddAlias par
/// cellule de la colonne A (comme les classeurs prepares a la main).
fn classeur_add_alias(
    feuilles: &[(String, Vec<String>)],
    elements: &LigneAddAlias,
) -> Resultat<Vec<u8>> {
    let mut classeur = Workbook::new();
    for (nom, numeros) in feuilles {
        let lignes = add_alias(numeros, elements);
        let onglet = classeur.add_worksheet();
        onglet.set_name(nom)?;
        // Colonne assez large pour lire chaque ligne en entier
        let largeur = lignes.iter().map(|l| l.chars().count()).max().unwrap_or(0) + 2;
        onglet.set_column_width(0, largeur.min(255) as f64)?;
        for (index, ligne) in lignes.iter().enumerate() {
            onglet.write_string(index as u32, 0, ligne)?;
        }
    }
    Ok(classeur.save_to_buffer()?)
}

// ---------------------------------------------------------------------------
// Zip
// ---------------------------------------------------------------------------

fn construire_zip(fichiers: &[(String, Vec<u8>)]) -> Resultat<Vec<u8>> {
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (nom, octets) in fichiers {
        zip.start_file(nom.as_str(), options)?;
        zip.write_all(octets)?;
    }
    Ok(zip.finish()?.into_inner())
}

// ---------------------------------------------------------------------------
// Historique (SQLite + zips sur disque)
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Serialize)]
pub struct EntreeHistorique {
    pub id: i64,
    pub nom_original: String,
    /// Fichier produit : .zip (Standard) ou .xlsx (CRPLMT). Le nom de colonne
    /// date d'avant la sortie .xlsx et reste pour les bases existantes.
    pub nom_zip: String,
    pub chemin_zip: String,
    pub nb_feuilles: i64,
    pub taille_octets: i64,
    pub style_entete: String,
    pub date_creation: String,
}

#[derive(Debug, Clone)]
pub struct Historique {
    dossier_zips: PathBuf,
    base: PathBuf,
}

impl Historique {
    /// Historique dans `dossier` : `dossier/stockage/*.zip` et `dossier/historique.db`.
    pub fn ouvrir(dossier: &Path) -> Resultat<Self> {
        Self::avec_chemins(dossier.join("stockage"), dossier.join("historique.db"))
    }

    pub fn avec_chemins(dossier_zips: PathBuf, base: PathBuf) -> Resultat<Self> {
        std::fs::create_dir_all(&dossier_zips).map_err(|e| {
            Erreur::Interne(format!(
                "Impossible de creer le dossier {} : {e}",
                dossier_zips.display()
            ))
        })?;
        if let Some(parent) = base.parent() {
            std::fs::create_dir_all(parent)?;
        }

        let historique = Historique { dossier_zips, base };
        historique.connexion()?.execute(
            "CREATE TABLE IF NOT EXISTS historique (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nom_original TEXT NOT NULL,
                nom_zip TEXT NOT NULL,
                chemin_zip TEXT NOT NULL,
                nb_feuilles INTEGER NOT NULL,
                taille_octets INTEGER NOT NULL,
                style_entete TEXT NOT NULL,
                date_creation TEXT NOT NULL DEFAULT (datetime('now'))
            )",
            [],
        )?;
        Ok(historique)
    }

    pub fn dossier_zips(&self) -> &Path {
        &self.dossier_zips
    }

    pub fn base(&self) -> &Path {
        &self.base
    }

    fn connexion(&self) -> Resultat<Connection> {
        Connection::open(&self.base).map_err(|e| {
            Erreur::Interne(format!(
                "Impossible d'ouvrir la base {} : {e}",
                self.base.display()
            ))
        })
    }

    fn enregistrer(
        &self,
        nom_original: &str,
        nom_zip: &str,
        zip: &[u8],
        nb_feuilles: usize,
        style_entete: &str,
    ) -> Resultat<i64> {
        let chemin = self
            .dossier_zips
            .join(format!("{}.{}", Uuid::new_v4().simple(), extension(nom_zip)));
        std::fs::write(&chemin, zip)?;

        let connexion = self.connexion()?;
        connexion.execute(
            "INSERT INTO historique
             (nom_original, nom_zip, chemin_zip, nb_feuilles, taille_octets, style_entete)
             VALUES (?1, ?2, ?3, ?4, ?5, ?6)",
            params![
                nom_original,
                nom_zip,
                chemin.to_string_lossy(),
                nb_feuilles as i64,
                zip.len() as i64,
                style_entete
            ],
        )?;
        Ok(connexion.last_insert_rowid())
    }

    pub fn lister(&self, limite: i64) -> Resultat<Vec<EntreeHistorique>> {
        let connexion = self.connexion()?;
        let mut requete = connexion
            .prepare("SELECT * FROM historique ORDER BY date_creation DESC, id DESC LIMIT ?1")?;
        let entrees = requete
            .query_map(params![limite.clamp(1, 500)], entree_depuis_ligne)?
            .collect::<std::result::Result<Vec<_>, _>>()?;
        Ok(entrees)
    }

    fn entree(&self, id: i64) -> Resultat<EntreeHistorique> {
        let connexion = self.connexion()?;
        let mut requete = connexion.prepare("SELECT * FROM historique WHERE id = ?1")?;
        match requete.query_row(params![id], entree_depuis_ligne) {
            Ok(entree) => Ok(entree),
            Err(rusqlite::Error::QueryReturnedNoRows) => {
                Err(Erreur::Introuvable("Entree introuvable.".into()))
            }
            Err(e) => Err(e.into()),
        }
    }

    /// Nom et contenu du fichier produit par une conversion passee.
    pub fn lire_zip(&self, id: i64) -> Resultat<(String, Vec<u8>)> {
        let entree = self.entree(id)?;
        let octets = std::fs::read(&entree.chemin_zip).map_err(|_| {
            Erreur::Introuvable("Le fichier n'existe plus sur le disque.".into())
        })?;
        Ok((entree.nom_zip, octets))
    }

    pub fn supprimer(&self, id: i64) -> Resultat<()> {
        let entree = self.entree(id)?;
        let _ = std::fs::remove_file(&entree.chemin_zip);
        self.connexion()?
            .execute("DELETE FROM historique WHERE id = ?1", params![id])?;
        Ok(())
    }
}

fn entree_depuis_ligne(ligne: &rusqlite::Row<'_>) -> rusqlite::Result<EntreeHistorique> {
    Ok(EntreeHistorique {
        id: ligne.get("id")?,
        nom_original: ligne.get("nom_original")?,
        nom_zip: ligne.get("nom_zip")?,
        chemin_zip: ligne.get("chemin_zip")?,
        nb_feuilles: ligne.get("nb_feuilles")?,
        taille_octets: ligne.get("taille_octets")?,
        style_entete: ligne.get("style_entete")?,
        date_creation: ligne.get("date_creation")?,
    })
}

// ---------------------------------------------------------------------------
// Tests unitaires
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn lignes(valeurs: &[&[&str]]) -> Vec<Vec<String>> {
        valeurs
            .iter()
            .map(|l| l.iter().map(|v| v.to_string()).collect())
            .collect()
    }

    #[test]
    fn parametre_en_troisieme_colonne() {
        let l = lignes(&[
            &["ID", "14824152", "EleCashback"],
            &["ID", "14824153", "EleCashback"],
        ]);
        assert_eq!(parametre_le_plus_frequent(&l).as_deref(), Some("EleCashback"));
        assert_eq!(
            entete_standard(l.len(), "Set", "EleCashback"),
            "HDR,UpdateIndividualRatingParameter,Set EleCashback,ExtId12341,2,EleCashback enable,1"
        );
    }

    #[test]
    fn parametre_ligne_dans_une_seule_cellule() {
        let l = lignes(&[
            &["MSISDN,237674497307,AgentFloatRegion,SouthWest"],
            &["MSISDN,237674931273,AgentFloatRegion,West"],
        ]);
        assert_eq!(parametre_le_plus_frequent(&l).as_deref(), Some("AgentFloatRegion"));
        assert_eq!(
            entete_standard(l.len(), "Init", "AgentFloatRegion"),
            "HDR,\"UpdateIndividualRatingParameter\",\"Init AgentFloatRegion TRUE\",\"ExtId12340\",\"2\",\" AgentFloatRegion init\",\"1\""
        );
    }

    #[test]
    fn numeros_titre_ignore_et_espaces_retires() {
        let l = lignes(&[&["msisdn"], &["237653282055"], &["237 681 178 408"]]);
        assert_eq!(
            extraire_numeros(&l).unwrap(),
            vec!["237653282055", "237681178408"]
        );
    }

    #[test]
    fn numeros_valeur_invalide_refusee() {
        let l = lignes(&[&["msisdn"], &["237653282055"], &["abc"]]);
        assert!(extraire_numeros(&l).unwrap_err().contains("ligne 3"));
    }

    #[test]
    fn add_alias_identique_au_template() {
        let elements = ElementsAddAlias::default().resoudre().unwrap();
        let numeros = vec!["237653282055".to_string(), "237679447586".to_string()];
        assert_eq!(
            add_alias(&numeros, &elements),
            vec![
                "HDR,\"AddAlias\",\"Add Alias Title\",\"ExtId12340\",\"2\",\"Agent  advance pilote\",\"1\"",
                "MSISDN,\"237653282055\",\"CRPLMT_237653282055@100000\"",
                "MSISDN,\"237679447586\",\"CRPLMT_237679447586@100000\"",
            ]
        );
    }

    #[test]
    fn add_alias_elements_du_formulaire() {
        let elements = ElementsAddAlias {
            identifiant: Some(" MSISDN ".into()),
            prefixe: Some("PROMO".into()),
            separateur: Some("-".into()),
            symbole: Some("#".into()),
            montant: Some("50000".into()),
            ext_id: Some("ExtId999".into()),
            ..Default::default()
        }
        .resoudre()
        .unwrap();
        let lignes = add_alias(&["237671598153".to_string()], &elements);
        assert_eq!(
            lignes[0],
            "HDR,\"AddAlias\",\"Add Alias Title\",\"ExtId999\",\"1\",\"Agent  advance pilote\",\"1\""
        );
        assert_eq!(lignes[1], "MSISDN,\"237671598153\",\"PROMO-237671598153#50000\"");
    }

    #[test]
    fn add_alias_elements_invalides_refuses() {
        let montant = ElementsAddAlias { montant: Some("10a".into()), ..Default::default() };
        assert!(montant.resoudre().is_err());
        let guillemet = ElementsAddAlias { prefixe: Some("CR\"PL".into()), ..Default::default() };
        assert!(guillemet.resoudre().is_err());
        let virgule = ElementsAddAlias { identifiant: Some("MS,ISDN".into()), ..Default::default() };
        assert!(virgule.resoudre().is_err());
    }
}
