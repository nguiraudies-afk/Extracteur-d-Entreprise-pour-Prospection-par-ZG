from datetime import date, datetime
import io
import json
from pathlib import Path
from typing import Any, Optional

import gspread
from google.oauth2.service_account import Credentials
import openpyxl
import pandas as pd
import requests
import streamlit as st


# ==============================================================================
# CONFIGURATION
# ==============================================================================
API_URL = "https://recherche-entreprises.api.gouv.fr/search"
FICHIER_CUIVRE = Path("calendrier_cuivre_orange.json")
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ==============================================================================
# UTILITAIRES & NORMALISATION
# ==============================================================================
def normaliser_code_ape(code: Any) -> str:
    """Nettoie et uniformise un code APE / NAF (ex: '62.01z' -> '62.01Z')."""
    if not code:
        return ""
    return str(code).strip().upper()


def parser_date_flexible(valeur: Any) -> Optional[date]:
    """Convertit une date (ISO AAAA-MM-JJ, FR JJ/MM/AAAA, date ou datetime) en objet date."""
    if valeur is None:
        return None
    if isinstance(valeur, datetime):
        return valeur.date()
    if isinstance(valeur, date):
        return valeur

    texte = str(valeur).strip()
    if not texte or texte.lower() in {
        "non renseignée",
        "non renseignee",
        "none",
        "nat",
        "nan",
        "null",
    }:
        return None

    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(texte[:10], fmt).date()
        except ValueError:
            continue
    return None


def formater_date_fr(valeur: Any, repli: str = "Non renseignée") -> str:
    """Formate une date en JJ/MM/AAAA pour l'affichage."""
    d = parser_date_flexible(valeur)
    return d.strftime("%d/%m/%Y") if d else repli


def normaliser_dates_cuivre(dates_brutes: Any) -> str:
    """Trie chronologiquement et formate les dates cuivre (gère plusieurs dates séparées par ' / ')."""
    if dates_brutes is None:
        return "Non renseignée"

    texte = str(dates_brutes).strip()
    if not texte or texte.lower() in {"non renseignée", "non renseignee"}:
        return "Non renseignée"

    morceaux = [m.strip() for m in texte.split("/") if m.strip()]
    # Cas d'une seule date au format JJ/MM/AAAA : on ne scinde pas les '/' de la date
    if len(morceaux) == 3 and all(p.isdigit() for p in morceaux):
        d = parser_date_flexible(texte)
        return d.strftime("%d/%m/%Y") if d else texte

    parties_multiples = [p.strip() for p in texte.split(" / ") if p.strip()]
    dates_parsees: list[date] = []
    autres: list[str] = []

    for p in parties_multiples:
        d = parser_date_flexible(p)
        if d:
            dates_parsees.append(d)
        else:
            autres.append(p)

    formatees = [d.strftime("%d/%m/%Y") for d in sorted(set(dates_parsees))]
    for a in autres:
        if a not in formatees:
            formatees.append(a)

    return " / ".join(formatees) if formatees else "Non renseignée"


# ==============================================================================
# CHARGEMENT DU CALENDRIER RÉSEAU CUIVRE
# ==============================================================================
@st.cache_data(show_spinner=False)
def charger_calendrier_cuivre(chemin: Path = FICHIER_CUIVRE) -> tuple[dict[str, tuple[str, str, str, str]], dict[str, list[str]]]:
    """
    Charge le fichier calendrier_cuivre_orange.json.
    Retourne :
      - communes : dict[code_insee, (code_insee, nom_commune, date_fermeture, lot)]
      - par_cp : dict[code_postal, list[code_insee]]
    """
    communes: dict[str, tuple[str, str, str, str]] = {}
    par_cp: dict[str, list[str]] = {}

    if not chemin.exists():
        return communes, par_cp

    with chemin.open("r", encoding="utf-8") as f:
        data = json.load(f)

    items = data if isinstance(data, list) else data.get("communes", [])

    for item in items:
        if not isinstance(item, dict):
            continue

        code_insee = str(item.get("code_insee") or item.get("insee") or "").strip()
        nom_commune = str(item.get("nom_commune") or item.get("commune") or "").strip()
        date_fermeture = normaliser_dates_cuivre(
            item.get("date_fermeture")
            or item.get("date_fin_cuivre")
            or item.get("date")
            or "Non renseignée"
        )
        lot = str(item.get("lot") or item.get("lot_cuivre") or "Non renseigné").strip()

        cps_bruts = item.get("codes_postaux") or item.get("code_postal") or []
        if isinstance(cps_bruts, (str, int)):
            cps = [str(cps_bruts).strip()]
        elif isinstance(cps_bruts, list):
            cps = [str(cp).strip() for cp in cps_bruts if str(cp).strip()]
        else:
            cps = []

        cle_commune = code_insee or f"{nom_commune}_{'_'.join(cps)}"
        if cle_commune:
            communes[cle_commune] = (code_insee, nom_commune, date_fermeture, lot)
            for cp in cps:
                par_cp.setdefault(cp, [])
                if cle_commune not in par_cp[cp]:
                    par_cp[cp].append(cle_commune)

    return communes, par_cp


def trouver_info_cuivre(
    code_insee: Optional[str],
    code_postal: Optional[str],
    communes: dict[str, tuple[str, str, str, str]],
    par_cp: dict[str, list[str]],
) -> tuple[str, str]:
    """Retourne (date_fin_cuivre_formatee, lot_cuivre) à partir du code INSEE ou du code postal."""
    insee_clean = str(code_insee or "").strip()
    if insee_clean and insee_clean in communes:
        _, _, date_fermeture, lot = communes[insee_clean]
        return normaliser_dates_cuivre(date_fermeture), lot or "Non renseigné"

    cp_clean = str(code_postal or "").strip()
    if cp_clean and cp_clean in par_cp and par_cp[cp_clean]:
        dates = sorted(
            {
                communes[i][2]
                for i in par_cp[cp_clean]
                if i in communes and communes[i][2] and communes[i][2] != "Non renseignée"
            }
        )
        lots = sorted(
            {
                communes[i][3]
                for i in par_cp[cp_clean]
                if i in communes and communes[i][3] and communes[i][3] != "Non renseigné"
            }
        )
        dates_str = normaliser_dates_cuivre(" / ".join(dates)) if dates else "Non renseignée"
        lots_str = " / ".join(lots) if lots else "Non renseigné"
        return dates_str, lots_str

    return "Non renseignée", "Non renseigné"


# ==============================================================================
# RECHERCHE & FORMATAGE DES ENTREPRISES
# ==============================================================================
def rechercher_entreprises_api(
    requete: str = "",
    code_postal: str = "",
    departement: str = "",
    code_ape: str = "",
    max_pages: int = 3,
    per_page: int = 25,
) -> list[dict[str, Any]]:
    """Appelle l'API Recherche d'Entreprises et retourne la liste brute des résultats."""
    params_base: dict[str, Any] = {
        "per_page": per_page,
        "etat_administratif": "A",  # Entreprises actives
    }
    if requete.strip():
        params_base["q"] = requete.strip()
    if code_postal.strip():
        params_base["code_postal"] = code_postal.strip()
    if departement.strip():
        params_base["departement"] = departement.strip()
    if code_ape.strip():
        params_base["activite_principale"] = normaliser_code_ape(code_ape)

    resultats: list[dict[str, Any]] = []
    for page in range(1, max_pages + 1):
        params = {**params_base, "page": page}
        resp = requests.get(API_URL, params=params, timeout=20)
        resp.raise_for_status()
        payload = resp.json()
        batch = payload.get("results", [])
        if not batch:
            break
        resultats.extend(batch)
        if page >= payload.get("total_pages", page):
            break

    return resultats


def formater_entreprise(
    ent: dict[str, Any],
    communes: dict[str, tuple[str, str, str, str]],
    par_cp: dict[str, list[str]],
) -> dict[str, Any]:
    """Transforme un résultat brut de l'API en ligne structurée avec dates normalisées."""
    siege = ent.get("siege") or {}
    code_insee = siege.get("commune") or ""
    code_postal = siege.get("code_postal") or ""

    date_fin_cuivre, lot_cuivre = trouver_info_cuivre(
        code_insee=code_insee,
        code_postal=code_postal,
        communes=communes,
        par_cp=par_cp,
    )

    date_creation_formatee = formater_date_fr(ent.get("date_creation"), repli="Non renseignée")
    date_fin_cuivre_formatee = normaliser_dates_cuivre(date_fin_cuivre)

    return {
        "SIREN": ent.get("siren") or "",
        "SIRET siège": siege.get("siret") or "",
        "Dénomination": ent.get("nom_complet") or ent.get("nom_raison_sociale") or "",
        "Code APE": normaliser_code_ape(ent.get("activite_principale") or siege.get("activite_principale")),
        "Adresse siège": siege.get("geo_adresse") or siege.get("adresse") or "",
        "Code postal": code_postal,
        "Commune": siege.get("libelle_commune") or "",
        "Code INSEE": code_insee,
        "Date de création": date_creation_formatee,
        "Fin Réseau Cuivre": date_fin_cuivre_formatee,
        "Lot Cuivre": lot_cuivre,
    }


# ==============================================================================
# EXPORTS (EXCEL & GOOGLE SHEETS)
# ==============================================================================
def generer_excel_bytes(df: pd.DataFrame) -> bytes:
    """Génère un classeur Excel (.xlsx) avec de vraies dates au format JJ/MM/AAAA."""
    df_excel = df.copy()
    if "Date de création" in df_excel.columns:
        df_excel["Date de création"] = df_excel["Date de création"].apply(parser_date_flexible)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_excel.to_excel(writer, index=False, sheet_name="Entreprises")
        ws = writer.sheets["Entreprises"]

        if "Date de création" in df_excel.columns:
            col_idx = df_excel.columns.get_loc("Date de création") + 1
            for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2):
                for c in cell:
                    if c.value is not None:
                        c.number_format = "DD/MM/YYYY"

        # Ajustement léger de la largeur des colonnes
        for column_cells in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in column_cells)
            col_letter = column_cells[0].column_letter
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    buffer.seek(0)
    return buffer.getvalue()


def enregistrer_dans_drive(df: pd.DataFrame, spreadsheet_id: str, nom_onglet: str = "Export") -> None:
    """Envoie le DataFrame (chaînes 'JJ/MM/AAAA' conservées) vers une feuille Google Sheets."""
    creds_info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    client = gspread.authorize(creds)

    sh = client.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(nom_onglet)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=nom_onglet, rows=max(len(df) + 10, 100), cols=max(len(df.columns) + 5, 15))

    df_clean = df.fillna("")
    lignes = [df_clean.columns.tolist()] + df_clean.astype(str).values.tolist()
    ws.update("A1", lignes)


# ==============================================================================
# INTERFACE STREAMLIT
# ==============================================================================
def main() -> None:
    st.set_page_config(page_title="Recherche Entreprises & Réseau Cuivre", layout="wide")
    st.title("Recherche d'entreprises & calendrier Réseau Cuivre")

    communes, par_cp = charger_calendrier_cuivre(FICHIER_CUIVRE)
    if not communes:
        st.warning(
            f"Le fichier `{FICHIER_CUIVRE}` est introuvable ou vide. "
            "La colonne « Fin Réseau Cuivre » affichera « Non renseignée »."
        )

    with st.form("formulaire_recherche"):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            requete = st.text_input("Mot-clé / Raison sociale / SIREN", placeholder="Ex: boulangerie")
        with col2:
            code_postal = st.text_input("Code postal", placeholder="Ex: 37000")
        with col3:
            departement = st.text_input("Département", placeholder="Ex: 37")
        with col4:
            code_ape = st.text_input("Code APE", placeholder="Ex: 10.71C")

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            date_creation_min = st.date_input(
                "Créées à partir du (optionnel)",
                value=None,
                format="DD/MM/YYYY",
            )
        with col_f2:
            masquer_cuivre_inconnu = st.checkbox("Masquer les lignes sans date cuivre connue", value=False)

        soumis = st.form_submit_button("Rechercher", type="primary")

    if soumis:
        try:
            with st.spinner("Recherche en cours..."):
                bruts = rechercher_entreprises_api(
                    requete=requete,
                    code_postal=code_postal,
                    departement=departement,
                    code_ape=code_ape,
                )
                lignes = [formater_entreprise(ent, communes, par_cp) for ent in bruts]
                df_final = pd.DataFrame(lignes)

                if not df_final.empty:
                    if date_creation_min:
                        dates_dt = df_final["Date de création"].apply(parser_date_flexible)
                        df_final = df_final[dates_dt.apply(lambda d: d is not None and d >= date_creation_min)]

                    if masquer_cuivre_inconnu:
                        df_final = df_final[df_final["Fin Réseau Cuivre"] != "Non renseignée"]

                st.session_state["df_final"] = df_final
        except requests.RequestException as exc:
            st.error(f"Erreur lors de l'appel à l'API : {exc}")

    df_final = st.session_state.get("df_final")
    if df_final is not None:
        if df_final.empty:
            st.info("Aucune entreprise active trouvée pour ce critère.")
        else:
            st.success(f"{len(df_final)} entreprise(s) trouvée(s).")
            st.dataframe(df_final, use_container_width=True, hide_index=True)

            excel_bytes = generer_excel_bytes(df_final)
            st.download_button(
                label="Télécharger l'export Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"entreprises_cuivre_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            spreadsheet_id = st.text_input(
                "ID du Google Sheet cible (optionnel)",
                placeholder="Ex: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
            )
            if st.button("Envoyer vers Google Sheets"):
                if not spreadsheet_id.strip():
                    st.warning("Renseigne l'ID du Google Sheet avant l'envoi.")
                else:
                    try:
                        with st.spinner("Envoi vers Google Sheets..."):
                            enregistrer_dans_drive(df_final, spreadsheet_id.strip())
                        st.success("Export envoyé avec succès vers Google Sheets !")
                    except Exception as exc:
                        st.error(f"Erreur lors de l'envoi vers Google Sheets : {exc}")


if __name__ == "__main__":
    main()
