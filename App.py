import io
import json
import os
import re
import time
import unicodedata
from typing import Any, Optional

import gspread
import pandas as pd
import requests
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# --- CONFIGURATION STREAMLIT ---
st.set_page_config(
    page_title="Extracteur de Prospection",
    page_icon="⚡",
    layout="wide",
)

# --- RÉFÉRENTIELS INSEE ---
EFFECTIFS_LIBELLES: dict[str, str] = {
    "NN": "Non renseigné",
    "00": "0 salarié",
    "01": "1 ou 2 salariés",
    "02": "3 à 5 salariés",
    "03": "6 à 9 salariés",
    "11": "10 à 19 salariés",
    "12": "20 à 49 salariés",
    "21": "50 à 99 salariés",
    "22": "100 à 199 salariés",
    "31": "200 à 249 salariés",
    "32": "250 à 499 salariés",
    "41": "500 à 999 salariés",
    "42": "1 000 à 1 999 salariés",
    "51": "2 000 à 4 999 salariés",
    "52": "5 000 à 9 999 salariés",
    "53": "10 000 salariés et plus",
}

FORMES_JURIDIQUES_DE_BASE: dict[str, str] = {
    "1000": "Entrepreneur individuel",
    "5498": "EURL",
    "5499": "SARL / Société à responsabilité limitée",
    "5710": "SAS (Société par actions simplifiée)",
    "5720": "SASU (Société par actions simplifiée unipersonnelle)",
    "6540": "SCI (Société civile immobilière)",
    "9220": "Association déclarée",
}


# --- CHARGEMENT DES RÉFÉRENTIELS JSON & UTILITAIRES ---
@st.cache_data
def charger_dictionnaire_json(nom_fichier: str, repli: Optional[dict[str, str]] = None) -> dict[str, str]:
    dossier_script = os.path.dirname(os.path.abspath(__file__))
    chemin_fichier = os.path.join(dossier_script, nom_fichier)
    if os.path.exists(chemin_fichier):
        with open(chemin_fichier, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return {str(k).strip().upper(): str(v).strip() for k, v in data.items()}
    return dict(repli) if repli else {}


def normaliser_texte_commune(texte: Optional[str]) -> str:
    if not texte:
        return ""
    s = unicodedata.normalize("NFKD", str(texte)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-zA-Z0-9]+", " ", s).strip().upper()


@st.cache_data
def charger_calendrier_cuivre(nom_fichier: str = "calendrier_cuivre_orange.json") -> dict[str, Any]:
    dossier_script = os.path.dirname(os.path.abspath(__file__))
    chemin_fichier = os.path.join(dossier_script, nom_fichier)
    if os.path.exists(chemin_fichier):
        with open(chemin_fichier, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return data
    return {"communes": {}, "par_cp": {}, "par_commune": {}}


def trouver_info_cuivre(
    calendrier: dict[str, Any],
    code_insee: Optional[str],
    nom_commune: Optional[str],
    code_postal: Optional[str],
) -> tuple[str, str]:
    """Retourne (date_fermeture_cuivre, lot_cuivre)."""
    communes = calendrier.get("communes") or {}
    par_cp = calendrier.get("par_cp") or {}
    par_commune = calendrier.get("par_commune") or {}

    insee_clean = str(code_insee or "").strip()
    if insee_clean in communes:
        _, _, date_fin, lot = communes[insee_clean]
        return date_fin or "Non renseignée", lot or ""

    cle_commune = normaliser_texte_commune(nom_commune)
    if cle_commune in par_commune and par_commune[cle_commune]:
        insee_match = par_commune[cle_commune][0]
        if insee_match in communes:
            _, _, date_fin, lot = communes[insee_match]
            return date_fin or "Non renseignée", lot or ""

    cp_clean = str(code_postal or "").strip()
    if cp_clean in par_cp and par_cp[cp_clean]:
        dates = sorted({communes[i][2] for i in par_cp[cp_clean] if i in communes and communes[i][2]})
        lots = sorted({communes[i][3] for i in par_cp[cp_clean] if i in communes and communes[i][3]})
        return " / ".join(dates) if dates else "Non renseignée", " / ".join(lots)

    return "Non renseignée", ""


def normaliser_code_ape(code_brut: Optional[str]) -> str:
    code = (code_brut or "").strip().upper()
    if len(code) == 5 and "." not in code:
        return f"{code[:2]}.{code[2:]}"
    return code


def nettoyer_nom_fichier(nom_brut: str, code_postal: str) -> str:
    nom = (nom_brut or "").strip()
    if not nom:
        return f"entreprises_{code_postal}"
    if nom.lower().endswith(".xlsx"):
        nom = nom[:-5]
    return re.sub(r'[\\/*?:"<>|]', "_", nom)


def a_procedure_collective(ent: dict[str, Any]) -> bool:
    """Détecte un indicateur de procédure collective si présent dans les champs retournés."""
    for conteneur_nom in ("complements", "finances"):
        conteneur = ent.get(conteneur_nom)
        if isinstance(conteneur, dict):
            for cle, val in conteneur.items():
                if "procedure" in str(cle).lower() and val:
                    return True
                if isinstance(val, dict):
                    for sous_cle, sous_val in val.items():
                        if "procedure" in str(sous_cle).lower() and sous_val:
                            return True
    return False


# --- GOOGLE SHEETS & DRIVE (OPTIONNEL) ---
def initialiser_connexion_google() -> tuple[Optional[gspread.Client], Any]:
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        if "gcp_service_account" in st.secrets:
            secret_brut = st.secrets["gcp_service_account"]
            info = json.loads(secret_brut) if isinstance(secret_brut, str) else dict(secret_brut)
            creds = Credentials.from_service_account_info(info, scopes=scopes)
            return gspread.authorize(creds), build("drive", "v3", credentials=creds)

        chemin_creds = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")
        if os.path.exists(chemin_creds):
            creds = Credentials.from_service_account_file(chemin_creds, scopes=scopes)
            return gspread.authorize(creds), build("drive", "v3", credentials=creds)
    except Exception as e:
        st.error(f"Erreur d'authentification Google : {e}")
    return None, None


def obtenir_ou_creer_dossier(drive_service: Any, nom_dossier: str) -> str:
    requete = (
        f"name = '{nom_dossier}' "
        f"and mimeType = 'application/vnd.google-apps.folder' "
        f"and trashed = false"
    )
    resultats = drive_service.files().list(q=requete, spaces="drive", fields="files(id, name)").execute()
    fichiers = resultats.get("files", [])
    if fichiers:
        return fichiers[0]["id"]

    metadonnees = {"name": nom_dossier, "mimeType": "application/vnd.google-apps.folder"}
    dossier = drive_service.files().create(body=metadonnees, fields="id").execute()
    return dossier.get("id")


def enregistrer_dans_drive(df: pd.DataFrame, nom_dossier: str, nom_sheet: str) -> None:
    gc, drive_service = initialiser_connexion_google()
    if not gc or not drive_service:
        st.warning("⚠️ Connexion Google Drive indisponible (vérifiez vos secrets ou `credentials.json`).")
        return

    with st.spinner("💾 Enregistrement dans Google Drive..."):
        id_dossier_cible = obtenir_ou_creer_dossier(drive_service, nom_dossier)
        requete_fichier = (
            f"name = '{nom_sheet}' "
            f"and '{id_dossier_cible}' in parents "
            f"and mimeType = 'application/vnd.google-apps.spreadsheet' "
            f"and trashed = false"
        )
        fichiers_existants = drive_service.files().list(q=requete_fichier, fields="files(id)").execute().get("files", [])

        if fichiers_existants:
            sh = gc.open_by_key(fichiers_existants[0]["id"])
        else:
            sh = gc.create(nom_sheet)
            id_fichier = sh.id
            fichier = drive_service.files().get(fileId=id_fichier, fields="parents").execute()
            parents_precedents = ",".join(fichier.get("parents", []))
            drive_service.files().update(
                fileId=id_fichier,
                addParents=id_dossier_cible,
                removeParents=parents_precedents,
                fields="id, parents",
            ).execute()

        df_drive = df.fillna("")
        worksheet = sh.get_worksheet(0)
        worksheet.clear()
        worksheet.update([df_drive.columns.values.tolist()] + df_drive.values.tolist())
        st.success(f"☁️ Fichier synchronisé dans Google Drive (`{nom_dossier}/{nom_sheet}`).")


# --- APPELS API ---
def obtenir_code_postal_par_ville(nom_ville: str) -> tuple[Optional[str], Optional[str]]:
    url_geo = "https://geo.api.gouv.fr/communes"
    params = {"nom": nom_ville, "fields": "codesPostaux", "boost": "population", "limit": 1}
    try:
        response = requests.get(url_geo, params=params, timeout=10)
        if response.status_code == 200:
            resultats = response.json()
            if resultats and resultats[0].get("codesPostaux"):
                return resultats[0]["codesPostaux"][0], resultats[0].get("nom")
    except requests.exceptions.RequestException as e:
        st.error(f"Erreur recherche ville : {e}")
    return None, None


def requeter_page_entreprises(
    code_postal: str,
    page: int,
    per_page: int,
    placeholder_progression: Any,
) -> tuple[Optional[dict[str, Any]], bool]:
    url = "https://recherche-entreprises.api.gouv.fr/search"
    params = {
        "code_postal": code_postal,
        "etat_administratif": "A",
        "page": page,
        "per_page": per_page,
    }

    for tentative in range(4):
        try:
            response = requests.get(url, params=params, timeout=15)
            if response.status_code == 429:
                attente = 2 * (tentative + 1)
                placeholder_progression.text(
                    f"⏳ Limite API atteinte (429), pause de {attente}s avant nouvelle tentative..."
                )
                time.sleep(attente)
                continue

            if response.status_code != 200:
                st.error(f"Erreur API (Code {response.status_code})")
                return None, False

            return response.json(), True
        except requests.exceptions.RequestException:
            if tentative == 3:
                st.error("❌ Connexion interrompue avec l'API après plusieurs tentatives.")
                return None, False
            time.sleep(2)

    st.error("Erreur API (Code 429) : trop de requêtes envoyées. Patiente quelques instants puis relance.")
    return None, False


def formater_entreprise(
    ent: dict[str, Any],
    code_postal_cible: str,
    libelles_naf: dict[str, str],
    formes_juridiques: dict[str, str],
    calendrier_cuivre: dict[str, Any],
    exclure_procedures: bool,
) -> Optional[dict[str, Any]]:
    etat_ent = str(ent.get("etat_administratif") or "A").strip().upper()
    siege = ent.get("siege") or {}
    etat_siege = str(siege.get("etat_administratif") or "A").strip().upper()

    if etat_ent != "A" or etat_siege != "A":
        return None

    if exclure_procedures and a_procedure_collective(ent):
        return None

    if str(siege.get("code_postal") or "").strip() != str(code_postal_cible).strip():
        return None

    finances = ent.get("finances") or {}
    ca: Any = "Non publié"
    resultat_net: Any = "Non publié"
    annee_bilan = ""

    if isinstance(finances, dict) and finances:
        derniere_annee = max(finances.keys())
        bilan = finances.get(derniere_annee) or {}
        annee_bilan = str(derniere_annee)
        ca_val = bilan.get("ca")
        res_val = bilan.get("resultat_net")
        ca = "Non publié" if ca_val is None else ca_val
        resultat_net = "Non publié" if res_val is None else res_val

    code_eff = str(ent.get("tranche_effectif_salarie") or "NN").strip().upper()
    score_tri = int(code_eff) if code_eff.isdigit() else -1

    raison_sociale = ent.get("nom_raison_sociale") or ent.get("nom_complet") or ""
    enseignes = siege.get("liste_enseignes") or []
    enseigne_str = ", ".join(str(x) for x in enseignes) if isinstance(enseignes, list) else str(enseignes)

    code_ape = normaliser_code_ape(ent.get("activite_principale") or siege.get("activite_principale"))
    libelle_activite = (
        ent.get("libelle_activite_principale")
        or siege.get("libelle_activite_principale")
        or libelles_naf.get(code_ape)
        or (f"Code APE {code_ape}" if code_ape else "Non renseignée")
    )

    code_nj = str(ent.get("nature_juridique") or "").strip().upper()
    forme_juridique = formes_juridiques.get(code_nj, f"Code {code_nj}" if code_nj else "Non renseignée")

    code_insee = siege.get("code_commune") or ent.get("code_commune")
    date_fin_cuivre, lot_cuivre = trouver_info_cuivre(
        calendrier=calendrier_cuivre,
        code_insee=code_insee,
        nom_commune=siege.get("libelle_commune", ""),
        code_postal=siege.get("code_postal"),
    )

    return {
        "SIREN": ent.get("siren"),
        "Raison Sociale": raison_sociale,
        "Enseigne": enseigne_str,
        "Code APE": code_ape,
        "Activité": libelle_activite,
        "Adresse": siege.get("adresse"),
        "Code Postal": siege.get("code_postal"),
        "Ville": siege.get("libelle_commune", ""),
        "Fin Réseau Cuivre": date_fin_cuivre,
        "Lot Cuivre": lot_cuivre,
        "Date de création": ent.get("date_creation"),
        "Catégorie": ent.get("categorie_entreprise") or "EI / Non classé",
        "Tranche Effectif": EFFECTIFS_LIBELLES.get(code_eff, "Inconnu"),
        "Score Tri": score_tri,
        "Année Bilan": annee_bilan,
        "Chiffre d'Affaires (€)": ca,
        "Résultat Net (€)": resultat_net,
        "Statut": "Actif",
        "Forme Juridique": forme_juridique,
    }


def trier_et_preparer_dataframe(lignes: list[dict[str, Any]], ville_prioritaire: Optional[str]) -> pd.DataFrame:
    df = pd.DataFrame(lignes)
    if df.empty:
        return df

    if ville_prioritaire:
        cible_lower = ville_prioritaire.strip().lower()
        df["Priorité Ville"] = df["Ville"].apply(lambda x: 1 if str(x).strip().lower() == cible_lower else 2)
        df = df.sort_values(
            by=["Priorité Ville", "Score Tri", "Raison Sociale"],
            ascending=[True, False, True],
        )
        return df.drop(columns=["Priorité Ville", "Score Tri"])

    df = df.sort_values(by=["Score Tri", "Raison Sociale"], ascending=[False, True])
    return df.drop(columns=["Score Tri"])


# --- INTERFACE UTILISATEUR ---
libelles_naf = charger_dictionnaire_json("codes_naf_rev2.json")
formes_juridiques = charger_dictionnaire_json("formes_juridiques.json", FORMES_JURIDIQUES_DE_BASE)
calendrier_cuivre = charger_calendrier_cuivre("calendrier_cuivre_orange.json")

st.title("⚡ Extracteur de Prospection Automatique")
st.markdown("Saisissez une ville ou un code postal pour extraire les entreprises actives triées par effectif.")

if not libelles_naf:
    st.warning("⚠️ Fichier `codes_naf_rev2.json` introuvable à côté de `App.py`.")
if len(formes_juridiques) <= len(FORMES_JURIDIQUES_DE_BASE):
    st.info("ℹ️ Fichier `formes_juridiques.json` non détecté : un dictionnaire de base est utilisé.")
if not calendrier_cuivre.get("communes"):
    st.warning("⚠️ Fichier `calendrier_cuivre_orange.json` introuvable ou vide à côté de `App.py`.")

with st.form("form_extraction"):
    col1, col2 = st.columns(2)
    with col1:
        ville_ou_cp = st.text_input("Ville ou Code Postal", value="37350", placeholder="Ex: 37000 ou Tours")
    with col2:
        nom_fichier_perso = st.text_input(
            "Nom du fichier de sortie (sans .xlsx)",
            placeholder="Ex: Prospection_Tours_2026 (laisser vide pour auto)",
        )

    exclure_procedures = st.checkbox(
        "Exclure les entreprises avec procédure collective détectée",
        value=True,
    )

    with st.expander("⚙️ Options Google Drive (Facultatif)"):
        sauvegarder_drive = st.checkbox(
            "Enregistrer aussi dans Google Drive (nécessite une clé Google Cloud)",
            value=False,
        )
        nom_dossier_drive = st.text_input("Nom du dossier Google Drive", value="Liste Prospection")

    bouton_lancer = st.form_submit_button("🚀 Lancer l'extraction")

if bouton_lancer:
    saisie = ville_ou_cp.strip()
    dossier_nom = nom_dossier_drive.strip() or "Liste Prospection"

    if not saisie:
        st.warning("⚠️ Veuillez remplir le champ 'Ville ou Code Postal'.")
    else:
        with st.spinner("🔍 Analyse de la localisation..."):
            if saisie.isdigit():
                code_postal_final: Optional[str] = saisie
                ville_prioritaire: Optional[str] = None
            else:
                code_postal_final, ville_prioritaire = obtenir_code_postal_par_ville(saisie)

        if not code_postal_final:
            st.error(f"❌ Impossible de trouver un code postal pour '{saisie}'.")
        else:
            nom_base_fichier = nettoyer_nom_fichier(nom_fichier_perso, code_postal_final)
            cible_texte = f" ({ville_prioritaire})" if ville_prioritaire else ""
            st.info(f"📍 Cible : Code Postal {code_postal_final}{cible_texte} | 📄 Fichier : `{nom_base_fichier}.xlsx`")

            with st.spinner("⏳ Extraction en cours depuis l'API nationale des entreprises..."):
                page = 1
                per_page = 25
                resultats_finaux: list[dict[str, Any]] = []
                placeholder_progression = st.empty()

                while True:
                    data, ok = requeter_page_entreprises(
                        code_postal=code_postal_final,
                        page=page,
                        per_page=per_page,
                        placeholder_progression=placeholder_progression,
                    )
                    if not ok or not data:
                        break

                    entreprises = data.get("results") or []
                    total_pages = int(data.get("total_pages") or 1)

                    if not entreprises:
                        break

                    for ent in entreprises:
                        ligne = formater_entreprise(
                            ent=ent,
                            code_postal_cible=code_postal_final,
                            libelles_naf=libelles_naf,
                            formes_juridiques=formes_juridiques,
                            calendrier_cuivre=calendrier_cuivre,
                            exclure_procedures=exclure_procedures,
                        )
                        if ligne:
                            resultats_finaux.append(ligne)

                    placeholder_progression.text(
                        f"Extraction en cours... Page {page}/{total_pages} ({len(resultats_finaux)} entreprises trouvées)"
                    )

                    if page >= total_pages:
                        break
                    page += 1
                    time.sleep(0.35)

                placeholder_progression.empty()

            df_final = trier_et_preparer_dataframe(resultats_finaux, ville_prioritaire)

            if df_final.empty:
                st.info("Aucune entreprise active trouvée pour ce critère.")
            else:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_final.to_excel(writer, index=False, sheet_name="Entreprises")

                st.success(f"✅ {len(df_final)} entreprises actives trouvées !")
                st.dataframe(df_final.head(50), use_container_width=True)

                st.download_button(
                    label=f"📥 Télécharger {nom_base_fichier}.xlsx",
                    data=buffer.getvalue(),
                    file_name=f"{nom_base_fichier}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                if sauvegarder_drive:
                    enregistrer_dans_drive(df_final, dossier_nom, nom_base_fichier)
