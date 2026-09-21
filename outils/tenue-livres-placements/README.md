# Tenue de livres de placements américains (contribuable canadien)

Outil local, sans intelligence artificielle, pour tenir le PBR, calculer les gains et
pertes en capital, les revenus de placement, l'impôt étranger retenu et les données du
T1135 pour des comptes américains détenus par un contribuable canadien. Les montants
sont produits en CAD et en USD.

La seule communication externe est la requête de taux de change à l'API Valet de la
Banque du Canada, qui ne contient que des dates. Aucune donnée client ne quitte le poste.

## Deux façons de l'utiliser

- **Sans rien installer** : la macro Excel du dossier `excel-vba/`. Voir `excel-vba/LISEZMOI.md`.
  C'est la version à utiliser sur un poste où Python ne peut pas être installé.
- **Avec Python** : le script `calcul_placements.py`, décrit ci-dessous.

Les deux appliquent les mêmes règles et partagent les mêmes cas de test.

## Fichiers

| Fichier | Rôle |
| --- | --- |
| `calcul_placements.py` | Moteur de calcul et ligne de commande |
| `creer_gabarit.py` | Crée le gabarit de saisie (vide ou avec des données fictives) |
| `tests/test_calcul.py` | Tests unitaires (28) avec des taux synthétiques |
| `Gabarit_Client_Exemple.xlsx` | Démo : gabarit rempli de données fictives |
| `Sortie_Client_Exemple_données_fictives_2025.xlsx` | Démo : résultat du calcul |
| `taux_fictifs_demo.csv` | Taux de change inventés pour la démo, à ne jamais utiliser sur un vrai dossier |
| `excel-vba/TenueLivresPlacements.bas` | Le même moteur en macro Excel, avec ses tests intégrés |
| `excel-vba/LISEZMOI.md` | Installation et utilisation de la macro |

## Installation

Python 3.10 ou plus récent.

```
pip install openpyxl requests
pip install pytest        # facultatif, pour les tests
```

## Utilisation

1. Créer un gabarit vide pour un client, avec un onglet par compte :

   ```
   python creer_gabarit.py --client "Nom du client" --annee 2025 --comptes RCM-1001 RCM-1002
   ```

2. Remplir les cellules jaunes à partir des relevés papier. L'onglet Instructions du
   gabarit décrit chaque type d'opération.

3. Lancer le calcul :

   ```
   python calcul_placements.py Gabarit_Client.xlsx --annee 2025
   ```

   Les taux sont téléchargés de la Banque du Canada et conservés dans
   `taux_bdc_cache.csv` à côté du classeur. Les exécutions suivantes n'ont plus besoin
   de connexion tant que la période est couverte.

   Poste sans accès Internet : fournir un fichier de taux avec `--taux fichier.csv`
   (colonnes `date,serie,valeur`, série `FXUSDCAD` pour le quotidien et `FXAUSDCAD`
   pour la moyenne annuelle).

4. Ouvrir `Sortie_<client>_<année>.xlsx`. L'onglet Alertes se lit en premier.

Pour reproduire la démo :

```
python creer_gabarit.py --demo
python calcul_placements.py Gabarit_Client_Exemple.xlsx --annee 2025 --taux taux_fictifs_demo.csv
python -m pytest tests
```

## Structure du gabarit

Un fichier par client, c'est-à-dire par contribuable. Le PBR est regroupé par titre pour
le contribuable, tous comptes confondus (article 47 LIR). Les comptes d'un conjoint, d'une
société ou d'une fiducie vont dans un autre fichier.

| Onglet | Contenu |
| --- | --- |
| Paramètres | Client, année, convention de date, méthode de taux pour les revenus |
| Comptes | Un code par compte ; chaque code a son onglet |
| `<code de compte>` | Une ligne par opération |
| Relevés | Quantités et solde d'encaisse selon chaque relevé de fin de mois |
| 1099 | Cases du 1099 composite par compte et par année |

Les transferts entre comptes se saisissent une seule fois, dans l'onglet du compte
source, avec le compte de destination dans la colonne Compte contrepartie. Le script
crée l'autre côté. Un transfert en nature ne change pas le PBR.

L'historique complet reste dans le fichier d'une année à l'autre. Le PBR se reconstruit
depuis le début et seule l'année demandée est rapportée. Les positions acquises avant le
début du suivi entrent par une ligne SOLDE_OUVERTURE avec leur PBR CAD connu.

## Classeur de sortie

| Onglet | Contenu |
| --- | --- |
| Sommaire | Totaux de l'année en CAD et en USD prêts pour l'annexe 3, l'annexe G, le T2209 et le TP-772 |
| Alertes | Erreurs (opération rejetée), avertissements (à vérifier), informations |
| Par titre | Dividendes, intérêts, retenues et gains regroupés par placement, en CAD et en USD |
| Dispositions | Chaque vente : produit, PBR, gain avant et après perte apparente, en CAD et en USD |
| Revenus | Dividendes, intérêts et remboursements de capital avec retenue, plus sous-totaux par titre |
| Frais | Honoraires et frais, déductibilité à évaluer |
| Positions | PBR regroupé par titre au 31 décembre et détail par compte |
| Journal | Toutes les opérations avec le taux utilisé et le PBR courant après chacune |
| Rapprochement relevés | Écart entre les relevés saisis et le calcul, par mois |
| Rapprochement 1099 | Écart entre le 1099 et le côté USD du calcul |
| T1135 | Coût maximal dans l'année, coût de fin d'année, revenu et gain par bien |
| Taux | Chaque date d'opération et le taux BdC retenu |

## Règles appliquées

- Taux quotidien de la Banque du Canada à la date de l'opération. Sans taux publié ce
  jour-là, le jour ouvrable précédent le plus rapproché. Au-delà de sept jours sans taux,
  l'opération est rejetée avec une erreur.
- Convention de date au choix : date de transaction ou date de règlement. Un seul choix
  par fichier.
- Gains en capital toujours au taux quotidien. Revenus au taux quotidien ou à la moyenne
  annuelle BdC, selon le paramètre.
- PBR en coût moyen pondéré, en CAD, par titre et par contribuable. Commission d'achat
  ajoutée au coût, commission de vente déduite du produit.
- Perte apparente : perte refusée en proportion des titres rachetés dans les 30 jours
  avant ou après et encore détenus 30 jours après, tous comptes confondus. Le montant
  refusé est ajouté au PBR.
- Remboursement de capital : réduit le PBR. Un remboursement supérieur au PBR crée un
  gain en capital réputé.
- Retenue américaine sur dividendes supérieure à 15 % : avertissement, l'excédent n'est
  pas créditable.
- Le gain en USD est indicatif seulement. Seul le gain en CAD tient compte des règles
  fiscales.

## Ce que l'outil ne fait pas

- Gains ou pertes de change sur l'encaisse USD elle-même (paragraphe 39(2), exemption
  de 200 $ pour un particulier).
- Placements privés, sociétés en commandite et K-1, fonds monétaires, produits
  structurés.
- Choix de la catégorie T1135 par bien.
- Options, ventes à découvert, titres à revenu fixe achetés avec intérêts courus.

Ces éléments se traitent à part, à partir du journal produit.

## Vérification des résultats

Les tests couvrent le coût moyen, l'effet du change sur le gain, la perte apparente
totale, partielle et via un autre compte, les remboursements de capital, les
fractionnements, les transferts, les soldes d'ouverture, les deux conventions de date,
les rapprochements et le T1135. Ils ne remplacent pas la revue d'un dossier réel par la
personne responsable : la première année, comparer avec le calcul actuel sur un client
déjà produit.
