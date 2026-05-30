# Magazin-ci

Logiciel local de gestion et de vente pour commerces en Cote d'Ivoire.

Developpe par Datadev-ci Tel : +225 0714351471 Abidjan - datadev.wps@gmail.com

## Modules inclus

- Splash/loading au demarrage avec logo.
- Login securise par PIN sans affichage des informations de connexion.
- Interface graphique tactile modernisee.
- Dashboard commercial
- POS tactile avec tuiles produits
- Liste produit avec recherche et categories
- Panier client
- Paiement rapide: especes, Mobile Money, carte et credit
- Generation de recu texte dans `data/recus`
- Produits, categories, prix et stock minimum
- Ventes avec panier et encaissement
- Clients
- Fournisseurs
- Entrees, sorties et corrections de stock
- Inventaire physique avec ajustement des ecarts
- Caisse: recettes, depenses, apports et retraits
- Statistiques avec graphiques: ventes, top produits, paiements, valeur stock
- Rapports avec filtres par date, paiement, statut et recherche client/numero
- Export Microsoft Excel `.xls` et CSV
- Facturation professionnelle en FCFA avec informations fiscales
- Export FNE JSON de preparation
- Parametres du commerce

## Lancement

Installer Python 3 sur Windows, puis double-cliquer sur:

`lancer-magazin-ci.bat`

Les donnees sont stockees dans:

`data/magazin_ci.db`

## Facturation et FNE

Les ventes peuvent generer une facture professionnelle avec:

- numero de facture;
- client, contact, NCC;
- details produits, quantites, prix unitaires et total;
- total paye, reste a payer et mode de paiement;
- statut FNE et reference FNE quand disponible.

L'export FNE cree un fichier JSON dans `data/fne/`. Ce fichier sert a preparer la certification officielle via la plateforme DGI/FNE, l'API FNE ou un TERNE.

## Acces et securite

Les codes d'acces ne sont pas affiches sur la page de connexion.

## Propositions de fonctions a implementer ensuite

- gestion utilisateurs: admin, vendeur, magasinier, comptable;
- roles et PIN par utilisateur;
- cloture de caisse journaliere avec rapport imprimable;
- scan code-barres;
- achats fournisseurs et factures fournisseurs;
- bons de commande et receptions partielles;
- credit client avec echeancier et paiements partiels;
- alertes de rupture stock et proposition de reapprovisionnement;
- multi-boutiques et transfert de stock;
- sauvegarde/restauration automatique;
- impression thermique ESC/POS directe;
- tableau de bord marge brute, rentabilite produits et rotation stock;
- etiquettes produits code-barres;
- export comptable;
- signature/cachet sur facture.
