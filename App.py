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

# Principaux codes NAF / APE (tu peux en ajouter autant que tu veux)
LIBELLES_NAF = {
    "01.13Z": "Culture de légumes, de melons, de racines et de tubercules",
    "02.20Z": "Exploitation forestière",
    "43.31Z": "Travaux de plâtrerie",
    "43.99C": "Travaux de maçonnerie générale et gros œuvre de bâtiment",
    "49.32Z": "Transports de voyageurs par taxis",
    "56.29B": "Autres services de restauration n.c.a.",
    "71.12B": "Ingénierie, études techniques",
    "81.30Z": "Services d'aménagement paysager",
    "84.11Z": "Administration publique générale",
}

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
            return gspread.authorize(creds), build('drive', 'v3', credentials=creds)
        elif os.path.exists("credentials.json"):
            creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
            return gspread.authorize(creds), build('drive', 'v3', credentials=creds)
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
                return resultats[0]["codesPostaux"][0], resultats[0]['nom']
    except Exception as e:
        st.error(f"Erreur recherche ville : {e}")
    return None, None

def obtenir_ou_creer_dossier(drive_service, nom_dossier):
    requete = f"name = '{nom_dossier}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    resultats = drive_service.files().list(q=requete, spaces='drive', fields='files(id, name)').execute()
    fichiers = resultats.get('files', [])
    if fichiers:
        return fichiers[0]['id']
    else:
        metadonnees = {'name': nom_dossier, 'mimeType': 'application/vnd.google-apps.folder'}
        dossier = drive_service.files().create(body=metadonnees, fields='id').execute()
        return dossier.get('id')

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
            st.info(f"📍 Cible : Code Postal {code_postal_final}" + (f" ({ville_prioritaire})" if ville_prioritaire else "") + f" | 📄 Fichier : `{nom_base_fichier}.xlsx`")

            with st.spinner("⏳ Extraction en cours depuis l'API nationale des entreprises..."):
                url = "https://recherche-entreprises.api.gouv.fr/search"
                page = 1
                per_page = 25
                resultats_finaux = []
                placeholder_progression = st.empty()

                while True:
                    params = {"code_postal": code_postal_final, "etat_administratif": "A", "page": page, "per_page": per_page}
                    
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
                        code_ape = ent.get("activite_principale") or siege.get("activite_principale") or ""
                        libelle_activite = (
                            ent.get("libelle_activite_principale")
                            or siege.get("libelle_activite_principale")
                            or LIBELLES_NAF.get(code_ape, "")
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
                            "Code Postal": siege.get("code_postal"),
                            "Ville": libelle_commune,
                            "Date de création": ent.get("date_creation"),
                            "Catégorie": ent.get("categorie_entreprise") or "EI / Non classé",
                            "Tranche Effectif": EFFECTIFS_LIBELLES.get(code_eff, "Inconnu"),
                            "Score Tri": score_tri,
                            "Année Bilan": annee_bilan,
                            "Chiffre d'Affaires (€)": ca,
                            "Résultat Net (€)": resultat_net,
                            "Statut": "Actif",
                            "Forme Juridique": forme_juridique
                        })

                    placeholder_progression.text(f"Extraction en cours... Page {page}/{total_pages} ({len(resultats_finaux)} entreprises trouvées)")

                    if page >= total_pages:
                        break
                    page += 1
                    time.sleep(0.35)

                placeholder_progression.empty()

                df = pd.DataFrame(resultats_finaux)
                if not df.empty:
                    if ville_prioritaire:
                        nom_cible_lower = ville_prioritaire.strip().lower()
                        df["Priorité Ville"] = df["Ville"].apply(lambda x: 1 if str(x).strip().lower() == nom_cible_lower else 2)
                        df = df.sort_values(by=["Priorité Ville", "Score Tri", "Raison Sociale"], ascending=[True, False, True])
                        df = df.drop(columns=["Priorité Ville", "Score Tri"])
                    else:
                        df = df.sort_values(by=["Score Tri", "Raison Sociale"], ascending=[False, True])
                        df = df.drop(columns=["Score Tri"])

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Entreprises")

                    st.success(f"✅ {len(df)} entreprises actives trouvées !")
                    
                    st.download_button(
                        label=f"📥 Télécharger {nom_base_fichier}.xlsx",
                        data=buffer.getvalue(),
                        file_name=f"{nom_base_fichier}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    if sauvegarder_drive:
                        gc, drive_service = initialiser_connexion_google()
                        if gc and drive_service:
                            with st.spinner("💾 Enregistrement dans Google Drive..."):
                                id_dossier_cible = obtenir_ou_creer_dossier(drive_service, dossier_nom)
                                nom_sheet = nom_base_fichier

                                requete_fichier = f"
