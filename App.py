import io
import json
import os
import re
import time
import gspread
import pandas as pd
import requests
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Extracteur de Prospection",
    page_icon="⚡",
    layout="wide"
)

# Correspondance des codes effectifs INSEE
EFFECTIFS_LIBELLES = {
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
    "53": "10 000 salariés et plus"
}

# Principaux codes de nature / forme juridique INSEE
FORMES_JURIDIQUES = {
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
    "9300": "Fondation"
}


# Chargement de l'ensemble des codes NAF rév. 2 depuis le fichier JSON local
@st.cache_data
def charger_libelles_naf(chemin_fichier="codes_naf_rev2.json"):
    if os.path.exists(chemin_fichier):
        with open(chemin_fichier, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


LIBELLES_NAF = charger_libelles_naf()


def normaliser_code_ape(code_brut):
    code = (code_brut or "").strip().upper()
    if len(code) == 5 and "." not in code:
        return f"{code[:2]}.{code[2:]}"
    return code


# --- AUTHENTIFICATION GOOGLE SERVICES (OPTIONNELLE) ---
def initialiser_connexion_google():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        if "gcp_service_account" in st.secrets:
            info = json.loads(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(info, scopes=scopes)
            return gspread.authorize(creds), build("drive", "v3", credentials=creds)
        elif os.path.exists("credentials.json"):
            creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
            return gspread.authorize(creds), build("drive", "v3", credentials=creds)
    except Exception:
        pass
    return None, None


def obtenir_code_postal_par_ville(nom_ville):
    url_geo = "https://geo.api.gouv.fr/communes"
    params = {"nom": nom_ville, "fields": "codesPostaux", "boost": "population", "limit": 1}
    try:
        response = requests.get(url_geo, params=params, timeout=10)
        if response.status_code == 200:
            resultats = response.json()
            if resultats and "codesPostaux" in resultats[0]:
                return resultats[0]["codesPostaux"][0], resultats[0]["nom"]
    except Exception as e:
        st.error(f"Erreur recherche ville : {e}")
    return None, None


def obtenir_ou_creer_dossier(drive_service, nom_dossier):
    requete = f"name = '{nom_dossier}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    resultats = drive_service.files().list(q=requete, spaces="drive", fields="files(id, name)").execute()
    fichiers = resultats.get("files", [])
    if fichiers:
        return fichiers[0]["id"]
    metadonnees = {"name": nom_dossier, "mimeType": "application/vnd.google-apps.folder"}
    dossier = drive_service.files().create(body=metadonnees, fields="id").execute()
    return dossier.get("id")


def nettoyer_nom_fichier(nom_brut, code_postal):
    nom = nom_brut.strip()
    if not nom:
        return f"entreprises_{code_postal}"
    if nom.lower().endswith(".xlsx"):
        nom = nom[:-5]
    nom = re.sub(r'[\\/*?:"<>|]', "_", nom)
    return nom


# --- INTERFACE UTILISATEUR ---
st.title("⚡ Extracteur de Prospection Automatique")
st.markdown("Saisissez une ville ou un code postal pour extraire les entreprises actives triées par effectif.")

if not LIBELLES_NAF:
    st.warning("⚠️ Fichier `codes_naf_rev2.json` introuvable à côté du script : seuls les libellés renvoyés par l'API seront affichés.")

with st.form("form_extraction"):
    col1, col2 = st.columns(2)
    with col1:
        ville_ou_cp = st.text_input("Ville ou Code Postal", value="37350", placeholder="Ex: 37000 ou Tours")
    with col2:
        nom_fichier_perso = st.text_input(
            "Nom du fichier de sortie (sans .xlsx)",
            placeholder="Ex: Prospection_Tours_2026 (laisser vide pour auto)"
        )

    with st.expander("⚙️ Options Google Drive (Facultatif)"):
        sauvegarder_drive = st.checkbox("Enregistrer aussi dans Google Drive (nécessite une clé Google Cloud)", value=False)
        nom_dossier_drive = st.text_input("Nom du dossier Google Drive", value="Liste Prospection")

    bouton_lancer = st.form_submit_button("🚀 Lancer l'extraction")

if bouton_lancer:
    saisie = ville_ou_cp.strip()
    dossier_nom = nom_dossier_drive.strip()

    if not saisie:
        st.warning("⚠️ Veuillez remplir le champ 'Ville ou Code Postal'.")
    else:
        with st.spinner("🔍 Analyse de la localisation..."):
            if saisie.isdigit():
                code_postal_final = saisie
                ville_prioritaire = None
            else:
                code_postal_final, ville_prioritaire = obtenir_code_postal_par_ville(saisie)

        if not code_postal_final:
            st.error(f"❌ Impossible de trouver un code postal pour '{saisie}'.")
        else:
            nom_base_fichier = nettoyer_nom_fichier(nom_fichier_perso, code_postal_final)
            cible_texte = f" ({ville_prioritaire})" if ville_prioritaire else ""
            st.info(f"📍 Cible : Code Postal {code_postal_final}{cible_texte} | 📄 Fichier : `{nom_base_fichier}.xlsx`")

            with st.spinner("⏳ Extraction en cours depuis l'API nationale des entreprises..."):
                url = "https://recherche-entreprises.api.gouv.fr/search"
                page = 1
                per_page = 25
                resultats_finaux = []
                placeholder_progression = st.empty()

                while True:
                    params = {
                        "code_postal": code_postal_final,
                        "etat_administratif": "A",
                        "page": page,
                        "per_page": per_page
                    }

                    # --- BLOC SÉCURISÉ (GESTION ERREUR 429 ET COUPURES RÉSEAU) ---
                    response = None
                    for tentative in range(4):
                        try:
                            response = requests.get(url, params=params, timeout=15)
                            if response.status_code == 429:
                                attente = 2 * (tentative + 1)
                                placeholder_progression.text(f"⏳ Limite API atteinte (429), pause de {attente}s avant nouvelle tentative...")
                                time.sleep(attente)
                                continue
                            break
                        except requests.exceptions.RequestException:
                            if tentative == 3:
                                st.error("❌ Connexion interrompue avec l'API après plusieurs tentatives.")
                                response = None
                                break
                            time.sleep(2)

                    if response is None or response.status_code == 429:
                        st.error("Erreur API (Code 429) : trop de requêtes envoyées. Patiente quelques instants puis relance.")
                        break

                    if response.status_code != 200:
                        st.error(f"Erreur API (Code {response.status_code})")
                        break
                    # -----------------------------------------------------------

                    data = response.json()
                    entreprises = data.get("results", [])
                    total_pages = data.get("total_pages", 1)

                    if not entreprises:
                        break

                    for ent in entreprises:
                        siege = ent.get("siege", {})
                        if siege.get("code_postal") != str(code_postal_final):
                            continue

                        finances = ent.get("finances", {})
                        ca, resultat_net, annee_bilan = "Non publié", "Non publié", ""
                        if finances:
                            derniere_annee = max(finances.keys())
                            annee_bilan = derniere_annee
                            ca = finances[derniere_annee].get("ca", "Non publié")
                            resultat_net = finances[derniere_annee].get("resultat_net", "Non publié")

                        code_eff = ent.get("tranche_effectif_salarie") or "NN"
                        score_tri = int(code_eff) if code_eff.isdigit() else -1
                        libelle_commune = siege.get("libelle_commune", "")

                        raison_sociale = ent.get("nom_raison_sociale") or ent.get("nom_complet") or ""
                        enseignes = siege.get("liste_enseignes") or []
                        enseigne_str = ", ".join(enseignes) if isinstance(enseignes, list) else str(enseignes)

                        code_ape = normaliser_code_ape(
                            ent.get("activite_principale") or siege.get("activite_principale")
                        )
                        libelle_activite = (
                            ent.get("libelle_activite_principale")
                            or siege.get("libelle_activite_principale")
                            or LIBELLES_NAF.get(code_ape)
                            or (f"Code APE {code_ape}" if code_ape else "Non renseignée")
                        )

                        code_nj = str(ent.get("nature_juridique") or "").strip()
                        forme_juridique = FORMES_JURIDIQUES.get(code_nj, f"Code {code_nj}" if code_nj else "Non renseignée")

                        resultats_finaux.append({
                            "SIREN": ent.get("siren"),
                            "Raison Sociale": raison_sociale,
                            "Enseigne": enseigne_str,
                            "Code APE": code_ape,
                            "Activité": libelle_activite,
                            "Adresse": siege.get("adresse"),
                            "Code Postal": siege.get("code
