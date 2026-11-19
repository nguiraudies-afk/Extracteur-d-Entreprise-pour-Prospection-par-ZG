import io
import json
import os
import re
import time
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

FORMES_JURIDIQUES: dict[str, str] = {
    "1000": "Entrepreneur individuel (EI)",
    "5410": "SARL nationale",
    "5485": "Autre société à responsabilité limitée",
    "5498": "EURL",
    "5499": "SARL / Société à responsabilité limitée",
    "5505": "SA à conseil d'administration",
    "5510": "SA à directoire",
    "5599": "SA (Société anonyme)",
    "5699": "Autre personne morale de droit commercial",
    "5710": "SAS (Société par actions simplifiée)",
    "5720": "SASU (Société par actions simplifiée unipersonnelle)",
    "6220": "EARL",
    "6540": "SCI (Société civile immobilière)",
    "6599": "Autre société civile",
    "7111": "Autorité constitutionnelle",
    "7120": "Service central de l'État",
    "7210": "Commune et nouvelle commune",
    "7220": "Département",
    "7225": "Région",
    "7344": "Établissement public local à caractère administratif",
    "9220": "Association déclarée",
    "9221": "Association déclarée d'insertion par l'économique",
    "9300": "Fondation",
}


# --- CHARGEMENT NAF & UTILITAIRES ---
@st.cache_data
def charger_libelles_naf(nom_fichier: str = "codes_naf_rev2.json") -> dict[str, str]:
    dossier_script = os.path.dirname(os.path.abspath(__file__))
    chemin_fichier = os.path.join(dossier_script, nom_fichier)
    if os.path.exists(chemin_fichier):
        with open(chemin_fichier, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return {str(k).strip().upper(): str(v).strip() for k, v in data.items()}
    return {}


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


def formater_entreprise(ent: dict[str, Any], code_postal_cible: str, libelles_naf: dict[str, str]) -> Optional[dict[str, Any]]:
    siege = ent.get("siege") or {}
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

    code_nj = str(ent.get("nature_juridique") or "").strip()
    forme_juridique = FORMES_JURIDIQUES.get(code_nj, f"Code {code_nj}" if code_nj else "Non renseignée")

    return {
        "SIREN
