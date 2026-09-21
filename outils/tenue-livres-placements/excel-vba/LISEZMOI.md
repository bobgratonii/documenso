# Version Excel (macro VBA) : aucune installation

Cette version fait tout dans Excel. Le moteur de calcul est une macro (fichier
`TenueLivresPlacements.bas`) que vous importez une fois dans le gabarit. Rien à
installer, aucune intelligence artificielle, aucune donnée client qui sort du poste.
La seule connexion est vers la Banque du Canada pour les taux, et elle ne transmet que
des dates.

## Installation, une seule fois par fichier client

1. Ouvrir `Gabarit_Client_Exemple.xlsx` (ou un gabarit vide) dans Excel.
2. Appuyer sur **Alt+F11** pour ouvrir l'éditeur Visual Basic.
3. Menu **Fichier**, **Importer un fichier**, choisir `TenueLivresPlacements.bas`.
   Un module « TenueLivresPlacements » apparaît dans le projet, sous Modules.
4. Fermer l'éditeur (Alt+Q).
5. **Fichier**, **Enregistrer sous**, choisir le type **Classeur Excel (prenant en charge
   les macros) (*.xlsm)**. C'est ce fichier .xlsm que vous garderez pour ce client.

Si le fichier .bas a été téléchargé ou reçu par courriel, Windows peut le marquer comme
provenant d'Internet. Clic droit sur le fichier, Propriétés, cocher **Débloquer**, OK.

## Utilisation

- **Alt+F8**, choisir **TestsAutomatiques**, Exécuter. La macro rejoue une quinzaine de
  cas connus (coût moyen, effet de change, perte apparente, transfert, rapprochements)
  avec des taux synthétiques et affiche le résultat. Tout doit réussir. À faire une
  fois après l'import, et à refaire après toute modification de la macro.
- **Alt+F8**, choisir **Calculer**, Exécuter. La macro lit les onglets de saisie, va
  chercher les taux manquants à la Banque du Canada, les garde dans l'onglet Taux, puis
  crée ou remplace les onglets de sortie : Sommaire, Alertes, Par titre, Dispositions,
  Revenus, Frais, Positions, Journal, Rapprochement relevés, Rapprochement 1099, T1135
  et Taux utilisés.

Les onglets de saisie ne sont jamais modifiés par la macro. Les onglets de sortie sont
recréés à chaque calcul, donc ne rien y écrire à la main.

## Si la Banque du Canada est injoignable

L'onglet Alertes le dira. Deux solutions :

- Demander au service informatique d'autoriser `www.bankofcanada.ca` pour Excel.
- Télécharger le CSV à la main, une fois par année, à partir de cette adresse
  (remplacer les dates), puis coller les lignes date et valeur dans l'onglet Taux avec
  `FXUSDCAD` dans la colonne serie :

  ```
  https://www.bankofcanada.ca/valet/observations/FXUSDCAD/csv?start_date=2024-11-15&end_date=2025-12-31
  ```

  La moyenne annuelle, si vous choisissez cette méthode pour les revenus, est la série
  `FXAUSDCAD`. Sans elle, la macro calcule la moyenne des taux quotidiens de l'année, ce
  qui donne le même résultat.

## Pour un nouveau client

Copier le gabarit .xlsm d'un client existant, vider les onglets de comptes, de relevés
et de 1099, ajuster Paramètres et Comptes, et renommer les onglets de comptes pour
qu'ils correspondent exactement aux codes de l'onglet Comptes. La macro est déjà dans
le fichier.

## Équivalence avec la version Python

La macro reproduit le moteur Python (`calcul_placements.py`) règle pour règle. Les deux
versions partagent les mêmes cas de test. La version Python reste le point de référence
pour toute évolution et sert à produire le gabarit.
