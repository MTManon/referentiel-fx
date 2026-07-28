#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_fx.py — Alimentation automatique du referentiel de taux de change MAD/EUR.

Produit et maintient un fichier master data au format CSV :
    - une ligne par mois
    - taux moyen mensuel et taux de fin de mois
    - source, statut (provisoire / definitif) et horodatage de mise a jour

Le script NE TOUCHE JAMAIS a un classeur Excel : il ecrit uniquement le CSV.
Ce choix est deliberé — il evite tout conflit de verrou sur un fichier ouvert
ou en cours de synchronisation OneDrive, et preserve les saisies manuelles.

Usage
-----
    python maj_fx.py                          # mise a jour incrementale
    python maj_fx.py --depuis 2024-01         # (re)construction depuis un mois
    python maj_fx.py --source stooq           # force une source
    python maj_fx.py --comparer 2026-06       # reconcilie toutes les sources
    python maj_fx.py --controles              # affiche les controles seuls

Dependance : requests   (pip install requests)
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import logging
import os
import statistics
import sys
from collections import defaultdict

try:
    import requests
except ImportError:
    sys.exit("Dependance manquante. Executer : pip install requests")

# =====================================================================
# PARAMETRES — seule section a adapter
# =====================================================================

# Dossier de sortie. En deploiement local, pointer vers le dossier
# OneDrive synchronise consomme par Power Query.
DOSSIER_SORTIE = os.environ.get("FX_DOSSIER_SORTIE", os.path.dirname(os.path.abspath(__file__)))

FICHIER_CSV = "master_data_fx.csv"
FICHIER_CONTROLES = "controles_fx.json"
FICHIER_LOG = "journal_fx.log"

DEVISE_BASE = "EUR"
DEVISE_CIBLE = "MAD"
MOIS_DEBUT = "2024-01"          # debut de l'historique reconstitue
DECIMALES = 4
DELIMITEUR = ";"          # separateur du CSV — ";" pour un Excel en locale francaise

# Controles
BORNE_BASSE = 9.50             # plausibilite du taux MAD/EUR
BORNE_HAUTE = 11.50
JOURS_COTES_MIN = 15           # sous ce seuil, un mois clos reste provisoire
ECART_SOURCES_MAX = 0.005      # 0,5 % — seuil de materialite entre deux sources
ANCIENNETE_MAJ_MAX = 40        # jours — au-dela, la donnee est reputee perimee

TIMEOUT = 30
ENTETE_HTTP = {"User-Agent": "MyTower-FX-Referentiel/1.0"}

# Ordre de priorite des sources
SOURCES_PAR_DEFAUT = ["yahoo", "stooq"]

COLONNES = [
    "mois",
    "id_mois",
    "devise_base",
    "devise_cible",
    "taux_mad_par_eur_moyen",
    "taux_mad_par_eur_fin_mois",
    "nb_jours_cotes",
    "dernier_jour_cote",
    "source",
    "statut",
    "date_maj",
]

log = logging.getLogger("fx")


# =====================================================================
# OUTILS DE DATE
# =====================================================================

def debut_de_mois(d: dt.date) -> dt.date:
    return d.replace(day=1)


def mois_suivant(d: dt.date) -> dt.date:
    return (d.replace(day=28) + dt.timedelta(days=4)).replace(day=1)


def fin_de_mois(d: dt.date) -> dt.date:
    return mois_suivant(d) - dt.timedelta(days=1)


def id_mois(d: dt.date) -> str:
    return f"{d.year:04d}-{d.month:02d}"


def parse_mois(txt: str) -> dt.date:
    """Accepte 'AAAA-MM' ou 'AAAA-MM-JJ'."""
    txt = txt.strip()
    if len(txt) == 7:
        txt += "-01"
    return dt.date.fromisoformat(txt)


# =====================================================================
# ADAPTATEURS DE SOURCE
# Chacun renvoie une liste de couples (date, taux) en cours quotidiens.
# Ajouter une source = ajouter une fonction et l'inscrire dans ADAPTATEURS.
# =====================================================================

def source_yahoo(debut: dt.date, fin: dt.date) -> list[tuple[dt.date, float]]:
    """Cours quotidiens EUR/MAD. Marche (mid). Sans cle d'API."""
    p1 = int(dt.datetime.combine(debut, dt.time.min).timestamp())
    p2 = int(dt.datetime.combine(fin, dt.time.min).timestamp()) + 86400
    url = "https://query1.finance.yahoo.com/v8/finance/chart/EURMAD=X"
    rep = requests.get(
        url,
        params={"period1": p1, "period2": p2, "interval": "1d"},
        headers=ENTETE_HTTP,
        timeout=TIMEOUT,
    )
    rep.raise_for_status()
    bloc = rep.json()["chart"]["result"][0]
    horodates = bloc["timestamp"]
    clotures = bloc["indicators"]["quote"][0]["close"]
    return _assembler(horodates, clotures)


def source_stooq(debut: dt.date, fin: dt.date) -> list[tuple[dt.date, float]]:
    """Cours quotidiens EUR/MAD en CSV. Marche. Sans cle d'API."""
    url = "https://stooq.com/q/d/l/"
    rep = requests.get(
        url,
        params={
            "s": "eurmad",
            "i": "d",
            "d1": debut.strftime("%Y%m%d"),
            "d2": fin.strftime("%Y%m%d"),
        },
        headers=ENTETE_HTTP,
        timeout=TIMEOUT,
    )
    rep.raise_for_status()
    lignes = rep.text.strip().splitlines()
    if not lignes or "Date" not in lignes[0]:
        raise ValueError("Reponse stooq inattendue : " + rep.text[:120])
    lecteur = csv.DictReader(lignes)
    sortie = []
    for ligne in lecteur:
        try:
            sortie.append((dt.date.fromisoformat(ligne["Date"]), float(ligne["Close"])))
        except (ValueError, KeyError, TypeError):
            continue
    return sortie


def source_erapi(debut: dt.date, fin: dt.date) -> list[tuple[dt.date, float]]:
    """Taux spot du jour uniquement. Dernier recours : ne documente que le mois en cours."""
    rep = requests.get(
        "https://open.er-api.com/v6/latest/EUR", headers=ENTETE_HTTP, timeout=TIMEOUT
    )
    rep.raise_for_status()
    donnees = rep.json()
    taux = float(donnees["rates"]["MAD"])
    jour = dt.date.today()
    return [(jour, taux)]


def source_bam(debut: dt.date, fin: dt.date) -> list[tuple[dt.date, float]]:
    """
    Fixing officiel Bank Al-Maghrib — NON IMPLEMENTE.

    Bank Al-Maghrib ne publie pas d'API. La procedure de raccordement est
    decrite dans le guide, section « Passage a la source officielle ».
    Une fois la requete identifiee, implementer ici et inscrire "bam" en
    tete de SOURCES_PAR_DEFAUT : le reste du script est inchange.
    """
    raise NotImplementedError(
        "Adaptateur BAM non implemente — voir la section dediee du guide."
    )


ADAPTATEURS = {
    "yahoo": source_yahoo,
    "stooq": source_stooq,
    "erapi": source_erapi,
    "bam": source_bam,
}


def _assembler(horodates, valeurs) -> list[tuple[dt.date, float]]:
    sortie = []
    for ts, val in zip(horodates, valeurs):
        if val is None or ts is None:
            continue
        sortie.append((dt.datetime.fromtimestamp(ts, dt.timezone.utc).date(), float(val)))
    return sortie


# =====================================================================
# AGREGATION MENSUELLE
# =====================================================================

def agreger_par_mois(quotidiens: list[tuple[dt.date, float]]) -> dict[str, dict]:
    """Cours quotidiens -> taux moyen et taux de fin de mois, par mois."""
    paniers: dict[str, list[tuple[dt.date, float]]] = defaultdict(list)
    for jour, taux in quotidiens:
        if taux is None or taux <= 0:
            continue
        paniers[id_mois(jour)].append((jour, taux))

    resultat = {}
    for cle, cours in paniers.items():
        cours.sort(key=lambda x: x[0])
        taux_seuls = [t for _, t in cours]
        resultat[cle] = {
            "taux_mad_par_eur_moyen": round(statistics.fmean(taux_seuls), DECIMALES),
            "taux_mad_par_eur_fin_mois": round(cours[-1][1], DECIMALES),
            "nb_jours_cotes": len(cours),
            "dernier_jour_cote": cours[-1][0].isoformat(),
        }
    return resultat


def determiner_statut(cle_mois: str, nb_jours: int, aujourdhui: dt.date) -> str:
    """Un mois n'est definitif que clos et suffisamment cote."""
    premier = parse_mois(cle_mois)
    if fin_de_mois(premier) >= aujourdhui:
        return "provisoire"
    if nb_jours < JOURS_COTES_MIN:
        return "provisoire"
    return "definitif"


# =====================================================================
# COLLECTE : premiere source qui repond, dans l'ordre de priorite
# =====================================================================

def collecter(debut: dt.date, fin: dt.date, sources: list[str]) -> tuple[str, dict]:
    dernier_incident = None
    for nom in sources:
        adaptateur = ADAPTATEURS.get(nom)
        if adaptateur is None:
            log.warning("Source inconnue, ignoree : %s", nom)
            continue
        try:
            log.info("Interrogation de la source %s (%s -> %s)", nom, debut, fin)
            quotidiens = adaptateur(debut, fin)
            if not quotidiens:
                raise ValueError("aucun cours renvoye")
            mensuels = agreger_par_mois(quotidiens)
            log.info("Source %s : %d cours quotidiens, %d mois",
                     nom, len(quotidiens), len(mensuels))
            return nom, mensuels
        except NotImplementedError as exc:
            log.info("Source %s indisponible : %s", nom, exc)
            dernier_incident = exc
        except Exception as exc:
            log.warning("Echec de la source %s : %s", nom, exc)
            dernier_incident = exc
    raise RuntimeError(f"Aucune source n'a repondu. Dernier incident : {dernier_incident}")


# =====================================================================
# PERSISTANCE
# =====================================================================

def chemin(nom: str) -> str:
    return os.path.join(DOSSIER_SORTIE, nom)


def lire_csv() -> dict[str, dict]:
    fichier = chemin(FICHIER_CSV)
    if not os.path.exists(fichier):
        return {}
    with open(fichier, "r", encoding="utf-8-sig", newline="") as flux:
        return {ligne["id_mois"]: ligne
                for ligne in csv.DictReader(flux, delimiter=DELIMITEUR)}


def ecrire_csv(lignes: dict[str, dict]) -> None:
    fichier = chemin(FICHIER_CSV)
    temporaire = fichier + ".tmp"
    ordre = sorted(lignes.keys())
    with open(temporaire, "w", encoding="utf-8-sig", newline="") as flux:
        redacteur = csv.DictWriter(flux, fieldnames=COLONNES, delimiter=DELIMITEUR)
        redacteur.writeheader()
        for cle in ordre:
            redacteur.writerow({col: lignes[cle].get(col, "") for col in COLONNES})
    os.replace(temporaire, fichier)   # ecriture atomique : jamais de CSV tronque
    log.info("Ecrit : %s (%d mois)", fichier, len(ordre))


def fusionner(existant: dict[str, dict], nouveaux: dict[str, dict],
              source: str, aujourdhui: dt.date) -> tuple[dict[str, dict], list[str]]:
    """
    Regle de mise a jour : un mois deja marque definitif n'est jamais reecrit.
    Tout le reste est rafraichi. Les modifications sont journalisees.
    """
    horodatage = dt.datetime.now().isoformat(timespec="seconds")
    fusion = dict(existant)
    mouvements = []

    for cle, valeurs in sorted(nouveaux.items()):
        precedent = existant.get(cle)
        if precedent and precedent.get("statut") == "definitif":
            continue

        statut = determiner_statut(cle, valeurs["nb_jours_cotes"], aujourdhui)
        ligne = {
            "mois": parse_mois(cle).isoformat(),
            "id_mois": cle,
            "devise_base": DEVISE_BASE,
            "devise_cible": DEVISE_CIBLE,
            "taux_mad_par_eur_moyen": f"{valeurs['taux_mad_par_eur_moyen']:.{DECIMALES}f}",
            "taux_mad_par_eur_fin_mois": f"{valeurs['taux_mad_par_eur_fin_mois']:.{DECIMALES}f}",
            "nb_jours_cotes": valeurs["nb_jours_cotes"],
            "dernier_jour_cote": valeurs["dernier_jour_cote"],
            "source": source,
            "statut": statut,
            "date_maj": horodatage,
        }
        fusion[cle] = ligne

        if precedent is None:
            mouvements.append(f"{cle} cree ({statut}, {ligne['taux_mad_par_eur_fin_mois']})")
        elif (precedent.get("taux_mad_par_eur_fin_mois") != ligne["taux_mad_par_eur_fin_mois"]
              or precedent.get("statut") != statut):
            mouvements.append(
                f"{cle} maj {precedent.get('taux_mad_par_eur_fin_mois')} -> "
                f"{ligne['taux_mad_par_eur_fin_mois']} ({precedent.get('statut')} -> {statut})"
            )
    return fusion, mouvements


# =====================================================================
# CONTROLES
# =====================================================================

def controler(lignes: dict[str, dict], aujourdhui: dt.date) -> dict:
    """Recoupements d'integrite. Aucun taux n'est diffuse si un voyant est rouge."""
    resultats = []

    def ajouter(code, libelle, valeur, ok):
        resultats.append({
            "code": code, "libelle": libelle,
            "valeur": valeur, "statut": "OK" if ok else "ALERTE",
        })

    cles = sorted(lignes.keys())
    ajouter("C1", "Nombre de mois au referentiel", len(cles), len(cles) > 0)

    # Continuite de la serie
    trous = []
    for precedent, suivant in zip(cles, cles[1:]):
        attendu = id_mois(mois_suivant(parse_mois(precedent)))
        if attendu != suivant:
            trous.append(attendu)
    ajouter("C2", "Serie mensuelle continue", trous or "aucun trou", not trous)

    # Plausibilite
    hors_bornes = []
    for cle in cles:
        for champ in ("taux_mad_par_eur_moyen", "taux_mad_par_eur_fin_mois"):
            try:
                taux = float(lignes[cle][champ])
            except (TypeError, ValueError):
                hors_bornes.append(f"{cle}/{champ}: illisible")
                continue
            if not BORNE_BASSE <= taux <= BORNE_HAUTE:
                hors_bornes.append(f"{cle}/{champ}: {taux}")
    ajouter("C3", f"Taux dans les bornes [{BORNE_BASSE} ; {BORNE_HAUTE}]",
            hors_bornes or "tous dans les bornes", not hors_bornes)

    # Rupture de serie : variation mensuelle anormale
    sauts = []
    for precedent, suivant in zip(cles, cles[1:]):
        try:
            avant = float(lignes[precedent]["taux_mad_par_eur_fin_mois"])
            apres = float(lignes[suivant]["taux_mad_par_eur_fin_mois"])
        except (TypeError, ValueError):
            continue
        if avant and abs(apres - avant) / avant > 0.05:
            sauts.append(f"{precedent}->{suivant}: {avant} -> {apres}")
    ajouter("C4", "Variation mensuelle inferieure a 5 %", sauts or "aucune rupture", not sauts)

    # Fraicheur : le mode d'echec d'une automatisation est le silence
    anciennete = None
    if cles:
        horodatages = [lignes[c]["date_maj"] for c in cles if lignes[c].get("date_maj")]
        if horodatages:
            derniere = max(dt.datetime.fromisoformat(h) for h in horodatages)
            anciennete = (dt.datetime.now() - derniere).days
    ajouter("C5", f"Anciennete de la mise a jour inferieure a {ANCIENNETE_MAJ_MAX} jours",
            anciennete, anciennete is not None and anciennete <= ANCIENNETE_MAJ_MAX)

    # Couverture : le mois precedent doit etre au referentiel
    mois_attendu = id_mois(debut_de_mois(aujourdhui) - dt.timedelta(days=1))
    ajouter("C6", f"Mois clos {mois_attendu} present", mois_attendu in lignes,
            mois_attendu in lignes)

    voyant = "OK" if all(r["statut"] == "OK" for r in resultats) else "ALERTE"
    return {
        "voyant_global": voyant,
        "date_controle": dt.datetime.now().isoformat(timespec="seconds"),
        "controles": resultats,
    }


def ecrire_controles(rapport: dict) -> None:
    with open(chemin(FICHIER_CONTROLES), "w", encoding="utf-8") as flux:
        json.dump(rapport, flux, ensure_ascii=False, indent=2)


def afficher_controles(rapport: dict) -> None:
    print(f"\nVoyant global : {rapport['voyant_global']}")
    print("-" * 72)
    for ctrl in rapport["controles"]:
        marque = "OK   " if ctrl["statut"] == "OK" else "ALERTE"
        print(f"  [{marque}] {ctrl['code']} {ctrl['libelle']} : {ctrl['valeur']}")
    print("-" * 72)


# =====================================================================
# RECONCILIATION MULTI-SOURCES
# =====================================================================

def comparer_sources(cle_mois: str) -> int:
    """Interroge toutes les sources sur un mois et documente les ecarts."""
    premier = parse_mois(cle_mois)
    dernier = fin_de_mois(premier)
    releves = {}
    for nom, adaptateur in ADAPTATEURS.items():
        try:
            mensuels = agreger_par_mois(adaptateur(premier, dernier))
            if cle_mois in mensuels:
                releves[nom] = mensuels[cle_mois]
        except Exception as exc:
            print(f"  {nom:8s} : indisponible ({exc})")

    if not releves:
        print("Aucune source disponible.")
        return 1

    print(f"\nReconciliation des sources — {cle_mois}")
    print("-" * 72)
    print(f"  {'source':10s} {'moyen':>10s} {'fin de mois':>13s} {'jours':>7s}")
    for nom, val in releves.items():
        print(f"  {nom:10s} {val['taux_mad_par_eur_moyen']:>10.4f} "
              f"{val['taux_mad_par_eur_fin_mois']:>13.4f} {val['nb_jours_cotes']:>7d}")

    fins = [v["taux_mad_par_eur_fin_mois"] for v in releves.values()]
    if len(fins) > 1:
        ecart = (max(fins) - min(fins)) / min(fins)
        verdict = "OK" if ecart <= ECART_SOURCES_MAX else "ALERTE"
        print("-" * 72)
        print(f"  Ecart maximal entre sources : {ecart:.4%}  [{verdict}]")
    return 0


# =====================================================================
# POINT D'ENTREE
# =====================================================================

def configurer_journal() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s %(message)s",
        handlers=[
            logging.FileHandler(chemin(FICHIER_LOG), encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def main(argv=None) -> int:
    analyseur = argparse.ArgumentParser(
        description="Alimentation du referentiel de taux de change MAD/EUR."
    )
    analyseur.add_argument("--depuis", default=None,
                           help="mois de debut (AAAA-MM). Par defaut : les 3 derniers mois.")
    analyseur.add_argument("--source", action="append", default=None,
                           choices=list(ADAPTATEURS),
                           help="force une source. Repetable pour fixer l'ordre.")
    analyseur.add_argument("--comparer", metavar="AAAA-MM", default=None,
                           help="reconcilie toutes les sources sur un mois, sans ecriture.")
    analyseur.add_argument("--controles", action="store_true",
                           help="affiche les controles sur le CSV existant, sans appel reseau.")
    args = analyseur.parse_args(argv)

    os.makedirs(DOSSIER_SORTIE, exist_ok=True)
    configurer_journal()
    aujourdhui = dt.date.today()

    if args.comparer:
        return comparer_sources(args.comparer)

    if args.controles:
        rapport = controler(lire_csv(), aujourdhui)
        afficher_controles(rapport)
        ecrire_controles(rapport)
        return 0 if rapport["voyant_global"] == "OK" else 2

    # Fenetre d'interrogation : large en reconstruction, glissante en routine
    if args.depuis:
        debut = parse_mois(args.depuis)
    else:
        existant_initial = lire_csv()
        if existant_initial:
            debut = debut_de_mois(aujourdhui)
            for _ in range(3):
                debut = debut_de_mois(debut - dt.timedelta(days=1))
        else:
            debut = parse_mois(MOIS_DEBUT)
    fin = aujourdhui

    sources = args.source or SOURCES_PAR_DEFAUT
    log.info("=== Mise a jour du referentiel FX %s/%s ===", DEVISE_CIBLE, DEVISE_BASE)

    try:
        source, mensuels = collecter(debut, fin, sources)
    except RuntimeError as exc:
        log.error("%s", exc)
        log.error("Le referentiel n'est PAS a jour. Le CSV precedent est conserve.")
        return 1

    existant = lire_csv()
    fusion, mouvements = fusionner(existant, mensuels, source, aujourdhui)

    if mouvements:
        for mouvement in mouvements:
            log.info("  %s", mouvement)
    else:
        log.info("  aucun mouvement — referentiel deja a jour")

    ecrire_csv(fusion)
    rapport = controler(fusion, aujourdhui)
    ecrire_controles(rapport)
    afficher_controles(rapport)

    if rapport["voyant_global"] != "OK":
        log.warning("Voyant global en alerte : verifier avant toute diffusion.")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
