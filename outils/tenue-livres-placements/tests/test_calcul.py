# -*- coding: utf-8 -*-
"""Tests du moteur de calcul avec des taux synthétiques (aucune connexion)."""
import sys
from datetime import date, timedelta
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from calcul_placements import (  # noqa: E402
    ACHAT, AJUST_PBR, DIVIDENDE, FRACTIONNEMENT, INTERET, REMB_CAPITAL, SOLDE_OUVERTURE,
    TRANSFERT_ENCAISSE, TRANSFERT_TITRES, VENTE, ENCAISSE, METHODE_MOYENNE, CONVENTION_REGLEMENT,
    Compte, Ligne1099, LigneReleve, Parametres, Taux, TauxManquant, Transaction, calculer,
)

C1 = Compte("A", "Courtier X")
C2 = Compte("B", "Courtier X")


def taux_constant(valeur=1.30, debut=date(2024, 1, 1), fin=date(2025, 12, 31)):
    q = {}
    d = debut
    while d <= fin:
        if d.weekday() < 5:
            q[d] = valeur
        d += timedelta(days=1)
    return Taux(q)


def taux_par_date(par_date: dict, defaut=1.30):
    """Taux constant sauf aux dates indiquées (jours ouvrables seulement)."""
    t = taux_constant(defaut)
    for d, v in par_date.items():
        assert d.weekday() < 5, "utiliser un jour ouvrable"
        t.quotidiens[d] = v
    t.dates = sorted(t.quotidiens)
    return t


def tx(compte, type_, titre, d, **kw):
    return Transaction(compte=compte, type=type_, titre=titre, date_tx=d, **kw)


def params(annee=2025, **kw):
    return Parametres(client="Test", annee=annee, **kw)


# --- Taux ---------------------------------------------------------------------

def test_taux_fin_de_semaine_utilise_jour_ouvrable_precedent():
    t = taux_par_date({date(2025, 3, 7): 1.44})       # vendredi
    assert t.au(date(2025, 3, 8)) == 1.44            # samedi
    assert t.au(date(2025, 3, 9)) == 1.44            # dimanche
    assert t.utilises[date(2025, 3, 9)][0] == date(2025, 3, 7)


def test_taux_manquant_trop_longtemps_leve_une_erreur():
    t = Taux({date(2025, 1, 2): 1.40})
    with pytest.raises(TauxManquant):
        t.au(date(2025, 2, 15))
    with pytest.raises(TauxManquant):
        t.au(date(2024, 12, 1))


def test_moyenne_annuelle_calculee_si_non_publiee():
    t = Taux({date(2025, 1, 2): 1.40, date(2025, 1, 3): 1.42}, {2024: 1.37})
    assert t.moyenne(2024) == 1.37
    assert t.moyenne(2025) == pytest.approx(1.41)


# --- PBR moyen et effet de change ----------------------------------------------

def test_pbr_moyen_en_cad_et_gain_de_change():
    # 100 @ 10 USD à 1.30, 100 @ 12 USD à 1.40, vente 100 @ 11 USD à 1.35
    # (le 2e achat est à plus de 30 jours de la vente : pas de perte apparente)
    t = taux_par_date({date(2025, 1, 6): 1.30, date(2025, 1, 20): 1.40, date(2025, 3, 3): 1.35})
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),
        tx("A", ACHAT, "XYZ", date(2025, 1, 20), quantite=100, prix_usd=12),
        tx("A", VENTE, "XYZ", date(2025, 3, 3), quantite=100, prix_usd=11),
    ]
    r = calculer(txs, [C1], t, params())
    assert not r.erreurs()
    d = r.dispositions[0]
    assert d.pbr_cad == pytest.approx((1300 + 1680) / 2)        # 1490
    assert d.produit_cad == pytest.approx(1100 * 1.35)            # 1485
    assert d.gain_cad == pytest.approx(-5.0)
    assert d.gain_usd == pytest.approx(0.0)                       # 1100 - 1100
    p = r.pools_fin["XYZ"]
    assert p.quantite == 100
    assert p.cout_cad == pytest.approx(1490)
    assert p.cout_usd == pytest.approx(1100)


def test_commission_ajoutee_au_cout_et_deduite_du_produit():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=10, prix_usd=100, commission_usd=10),
        tx("A", VENTE, "XYZ", date(2025, 2, 3), quantite=10, prix_usd=120, commission_usd=10),
    ]
    r = calculer(txs, [C1], t, params())
    d = r.dispositions[0]
    assert d.pbr_cad == pytest.approx(1010 * 1.30)
    assert d.produit_cad == pytest.approx(1190 * 1.30)
    assert d.gain_cad == pytest.approx(180 * 1.30)
    assert r.etats_fin["A"].encaisse == pytest.approx(-1010 + 1190)


def test_vente_superieure_aux_titres_detenus_est_une_erreur():
    t = taux_constant()
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=10, prix_usd=100),
        tx("A", VENTE, "XYZ", date(2025, 2, 3), quantite=20, prix_usd=100),
    ]
    r = calculer(txs, [C1], t, params())
    assert any("détient" in a.message for a in r.erreurs())
    assert not r.dispositions


# --- Perte apparente ----------------------------------------------------------

def test_perte_apparente_totale_refusee_et_ajoutee_au_pbr():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),
        tx("A", VENTE, "XYZ", date(2025, 3, 3), quantite=100, prix_usd=8),
        tx("A", ACHAT, "XYZ", date(2025, 3, 13), quantite=100, prix_usd=8),  # rachat dans les 30 jours
    ]
    r = calculer(txs, [C1], t, params())
    d = r.dispositions[0]
    assert d.gain_cad_avant == pytest.approx(-200 * 1.30)
    assert d.perte_refusee_cad == pytest.approx(260)
    assert d.gain_cad == pytest.approx(0)
    assert r.pools_fin["XYZ"].cout_cad == pytest.approx(800 * 1.30 + 260)
    assert any("Perte apparente" in a.message for a in r.alertes)


def test_perte_apparente_partielle():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),
        tx("A", VENTE, "XYZ", date(2025, 3, 3), quantite=100, prix_usd=8),
        tx("A", ACHAT, "XYZ", date(2025, 3, 20), quantite=40, prix_usd=8),
    ]
    r = calculer(txs, [C1], t, params())
    d = r.dispositions[0]
    assert d.perte_refusee_cad == pytest.approx(260 * 0.40)
    assert d.gain_cad == pytest.approx(-260 * 0.60)


def test_perte_apparente_par_achat_dans_un_autre_compte():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),
        tx("A", VENTE, "XYZ", date(2025, 3, 3), quantite=100, prix_usd=8),
        tx("B", ACHAT, "XYZ", date(2025, 2, 20), quantite=100, prix_usd=9),  # achat 11 jours avant, autre compte
    ]
    r = calculer(txs, [C1, C2], t, params())
    d = r.dispositions[0]
    assert d.perte_refusee_cad > 0
    assert d.gain_cad == pytest.approx(0)


def test_pas_de_perte_apparente_sans_rachat():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),
        tx("A", VENTE, "XYZ", date(2025, 3, 3), quantite=100, prix_usd=8),
        tx("A", ACHAT, "XYZ", date(2025, 5, 5), quantite=100, prix_usd=8),  # plus de 30 jours
    ]
    r = calculer(txs, [C1], t, params())
    assert r.dispositions[0].perte_refusee_cad == 0
    assert r.dispositions[0].gain_cad == pytest.approx(-260)


# --- Revenus ------------------------------------------------------------------

def test_dividende_avec_retenue_taux_quotidien():
    t = taux_par_date({date(2025, 6, 2): 1.37})
    txs = [tx("A", DIVIDENDE, "XYZ", date(2025, 6, 2), brut_usd=100, retenue_usd=15)]
    r = calculer(txs, [C1], t, params())
    rv = r.revenus[0]
    assert rv.brut_cad == pytest.approx(137)
    assert rv.retenue_cad == pytest.approx(20.55)
    assert r.etats_fin["A"].encaisse == pytest.approx(85)
    assert not any(a.niveau == "AVERTISSEMENT" for a in r.alertes)


def test_dividende_taux_moyen_annuel():
    t = taux_par_date({date(2025, 6, 2): 1.50})
    t.annuels[2025] = 1.36
    txs = [tx("A", DIVIDENDE, "XYZ", date(2025, 6, 2), brut_usd=100, retenue_usd=15)]
    r = calculer(txs, [C1], t, params(methode_revenus=METHODE_MOYENNE))
    assert r.revenus[0].taux == 1.36
    assert r.revenus[0].brut_cad == pytest.approx(136)
    assert r.taux_moyen == 1.36


def test_retenue_superieure_a_15_pourcent_declenche_un_avertissement():
    t = taux_constant()
    txs = [tx("A", DIVIDENDE, "XYZ", date(2025, 6, 2), brut_usd=100, retenue_usd=30)]
    r = calculer(txs, [C1], t, params())
    assert any("15 %" in a.message and a.niveau == "AVERTISSEMENT" for a in r.alertes)


def test_interet_sur_encaisse_va_au_t1135_du_compte():
    t = taux_constant(1.30)
    txs = [tx("A", INTERET, ENCAISSE, date(2025, 12, 31), brut_usd=100)]
    r = calculer(txs, [C1], t, params())
    ligne = next(l for l in r.t1135 if l.element == "ENCAISSE A")
    assert ligne.revenu_cad == pytest.approx(130)


# --- Remboursement de capital, fractionnement, ajustement -----------------------

def test_remboursement_de_capital_reduit_le_pbr():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),
        tx("A", REMB_CAPITAL, "XYZ", date(2025, 6, 2), brut_usd=50),
    ]
    r = calculer(txs, [C1], t, params())
    assert r.pools_fin["XYZ"].cout_cad == pytest.approx(1300 - 65)
    assert r.pools_fin["XYZ"].cout_usd == pytest.approx(950)
    assert not r.dispositions


def test_remboursement_de_capital_superieur_au_pbr_cree_un_gain_repute():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=1),
        tx("A", REMB_CAPITAL, "XYZ", date(2025, 6, 2), brut_usd=150),
    ]
    r = calculer(txs, [C1], t, params())
    assert r.pools_fin["XYZ"].cout_cad == pytest.approx(0)
    assert r.dispositions[0].gain_cad == pytest.approx(50 * 1.30)


def test_fractionnement_double_la_quantite_sans_changer_le_cout():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),
        tx("A", FRACTIONNEMENT, "XYZ", date(2025, 6, 2), quantite=100),
        tx("A", VENTE, "XYZ", date(2025, 7, 7), quantite=200, prix_usd=6),
    ]
    r = calculer(txs, [C1], t, params())
    assert r.dispositions[0].pbr_cad == pytest.approx(1300)
    assert r.dispositions[0].gain_cad == pytest.approx((1200 - 1000) * 1.30)


def test_ajustement_pbr_distribution_fantome():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),
        tx("A", AJUST_PBR, "XYZ", date(2025, 12, 31), brut_usd=20),
    ]
    r = calculer(txs, [C1], t, params())
    assert r.pools_fin["XYZ"].cout_cad == pytest.approx(1300 + 26)


# --- Transferts inter-comptes et regroupement du PBR -------------------------------

def test_transfert_de_titres_ne_change_pas_le_pbr_regroupe():
    t = taux_par_date({date(2025, 1, 6): 1.30, date(2025, 2, 3): 1.40, date(2025, 4, 1): 1.50, date(2025, 5, 5): 1.35})
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),
        tx("B", ACHAT, "XYZ", date(2025, 2, 3), quantite=100, prix_usd=12),
        tx("A", TRANSFERT_TITRES, "XYZ", date(2025, 4, 1), quantite=100, contrepartie="B"),
        tx("B", VENTE, "XYZ", date(2025, 5, 5), quantite=50, prix_usd=11),
    ]
    r = calculer(txs, [C1, C2], t, params())
    assert not r.erreurs()
    assert r.etats_fin["A"].quantites["XYZ"] == 0
    assert r.etats_fin["B"].quantites["XYZ"] == 150
    d = r.dispositions[0]
    assert d.pbr_cad == pytest.approx((1300 + 1680) / 200 * 50)     # moyen du client, pas du compte
    assert r.pools_fin["XYZ"].quantite == 150


def test_transfert_encaisse_entre_comptes():
    t = taux_constant()
    txs = [
        tx("A", SOLDE_OUVERTURE, ENCAISSE, date(2025, 1, 6), brut_usd=1000),
        tx("A", TRANSFERT_ENCAISSE, ENCAISSE, date(2025, 2, 3), brut_usd=400, contrepartie="B"),
    ]
    r = calculer(txs, [C1, C2], t, params())
    assert r.etats_fin["A"].encaisse == pytest.approx(600)
    assert r.etats_fin["B"].encaisse == pytest.approx(400)


def test_transfert_vers_compte_inconnu_est_une_erreur():
    t = taux_constant()
    txs = [tx("A", TRANSFERT_TITRES, "XYZ", date(2025, 2, 3), quantite=10, contrepartie="Z")]
    r = calculer(txs, [C1], t, params())
    assert r.erreurs()


# --- Solde d'ouverture, années antérieures, convention de date ----------------------

def test_solde_ouverture_et_historique_avant_annee():
    t = taux_par_date({date(2025, 3, 3): 1.40})
    txs = [
        tx("A", SOLDE_OUVERTURE, "XYZ", date(2024, 1, 2), quantite=100, brut_usd=1000, pbr_cad_ouverture=1250),
        tx("A", ACHAT, "XYZ", date(2024, 6, 3), quantite=100, prix_usd=15),          # 2024 : 1.30 -> 1950
        tx("A", VENTE, "XYZ", date(2025, 3, 3), quantite=100, prix_usd=20),
    ]
    r = calculer(txs, [C1], t, params(2025))
    assert r.dispositions[0].pbr_cad == pytest.approx((1250 + 1950) / 2)
    assert r.dispositions[0].produit_cad == pytest.approx(2000 * 1.40)
    # la disposition de 2025 est la seule; le journal contient tout l'historique
    assert len(r.journal) == 3


def test_solde_ouverture_sans_pbr_cad_est_une_erreur():
    t = taux_constant()
    txs = [tx("A", SOLDE_OUVERTURE, "XYZ", date(2025, 1, 2), quantite=100, brut_usd=1000)]
    r = calculer(txs, [C1], t, params())
    assert r.erreurs()


def test_convention_date_de_reglement():
    t = taux_par_date({date(2025, 3, 3): 1.30, date(2025, 3, 5): 1.50})
    txs = [tx("A", ACHAT, "XYZ", date(2025, 3, 3), date_regl=date(2025, 3, 5), quantite=10, prix_usd=100)]
    r1 = calculer([txs[0]], [C1], t, params())
    assert r1.journal[0].taux == 1.30
    r2 = calculer([txs[0]], [C1], t, params(convention=CONVENTION_REGLEMENT))
    assert r2.journal[0].taux == 1.50


# --- Rapprochements et T1135 ---------------------------------------------------------

def test_rapprochement_releve_detecte_un_ecart():
    t = taux_constant()
    txs = [
        tx("A", SOLDE_OUVERTURE, ENCAISSE, date(2025, 1, 2), brut_usd=5000),
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=10, prix_usd=100),
        tx("A", ACHAT, "XYZ", date(2025, 2, 10), quantite=5, prix_usd=100),  # après le relevé de janvier
    ]
    releves = [
        LigneReleve("A", date(2025, 1, 31), "XYZ", 10),
        LigneReleve("A", date(2025, 1, 31), ENCAISSE, 4000),
        LigneReleve("A", date(2025, 2, 28), "XYZ", 20),     # relevé fictif faux : 15 en réalité
    ]
    r = calculer(txs, [C1], t, params(), releves)
    statuts = {(l.date_fin, l.titre): (l.statut, l.ecart) for l in r.rapprochement_releves}
    assert statuts[(date(2025, 1, 31), "XYZ")] == ("OK", 0)
    assert statuts[(date(2025, 1, 31), ENCAISSE)][0] == "OK"
    assert statuts[(date(2025, 2, 28), "XYZ")] == ("ÉCART", 5)


def test_rapprochement_1099():
    t = taux_constant()
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=10, prix_usd=100),
        tx("A", DIVIDENDE, "XYZ", date(2025, 6, 2), brut_usd=100, retenue_usd=15),
        tx("A", VENTE, "XYZ", date(2025, 9, 8), quantite=10, prix_usd=110, commission_usd=5),
    ]
    f = [Ligne1099("A", 2025, dividendes_usd=100, retenue_usd=15, interets_usd=0, produits_usd=1095)]
    r = calculer(txs, [C1], t, params(), lignes_1099=f)
    assert all(l.statut == "OK" for l in r.rapprochement_1099)
    f2 = [Ligne1099("A", 2025, dividendes_usd=120, retenue_usd=15, interets_usd=0, produits_usd=1095)]
    r2 = calculer(txs, [C1], t, params(), lignes_1099=f2)
    ecarts = [l for l in r2.rapprochement_1099 if l.statut == "ÉCART"]
    assert len(ecarts) == 1 and ecarts[0].ecart == pytest.approx(20)


def test_t1135_cout_maximal_et_fin_annee():
    t = taux_constant(1.30)
    txs = [
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=100, prix_usd=10),    # 1300
        tx("A", ACHAT, "XYZ", date(2025, 3, 3), quantite=100, prix_usd=10),    # 2600 (max)
        tx("A", VENTE, "XYZ", date(2025, 9, 8), quantite=150, prix_usd=10),    # reste 650
    ]
    r = calculer(txs, [C1], t, params())
    ligne = next(l for l in r.t1135 if l.element == "XYZ")
    assert ligne.cout_max_cad == pytest.approx(2600)
    assert ligne.cout_fin_cad == pytest.approx(650)
    assert ligne.gain_cad == pytest.approx(0)


def test_encaisse_negative_declenche_un_avertissement_une_fois():
    t = taux_constant()
    txs = [
        tx("A", SOLDE_OUVERTURE, ENCAISSE, date(2025, 1, 2), brut_usd=500),
        tx("A", ACHAT, "XYZ", date(2025, 1, 6), quantite=10, prix_usd=100),   # -500
        tx("A", ACHAT, "XYZ", date(2025, 1, 7), quantite=1, prix_usd=100),    # -600, pas de 2e alerte
    ]
    r = calculer(txs, [C1], t, params())
    msgs = [a for a in r.alertes if "Encaisse de A négative" in a.message]
    assert len(msgs) == 1 and msgs[0].origine == ""
