# Magazin-ci

Logiciel local de gestion et de vente pour commerces en Cote d'Ivoire.

Developpe par Datadev-ci Tel : +225 0714351471 Abidjan - datadev.wps@gmail.com

## Modules inclus

- Splash/loading au demarrage avec logo.
- Login securise par PIN sans affichage des informations de connexion.
- Utilisateurs avec roles et PIN: Admin et Vendeur par defaut.
- Interface graphique tactile modernisee.
- Dashboard commercial
- POS tactile avec tuiles produits
- Liste produit avec recherche et categories
- Fiche produit enrichie: Part Number, S/N, IMEI, modele, couleur, taille et image
- Panier client
- Paiement rapide: especes, Mobile Money, carte et credit
- Scan code-barres au POS avec ajout direct au panier, compatible lecteurs USB/Bluetooth en mode clavier
- Generation de recu texte dans `data/recus`
- Produits, categories, prix et stock minimum
- Ventes avec panier et encaissement
- Clients
- Fournisseurs
- Achats fournisseurs et approvisionnement
- Retours client avec remise en stock
- Entrees, sorties et corrections de stock
- Inventaire physique avec ajustement des ecarts
- Caisse: ouverture, cloture, recettes, depenses, apports, retraits et rapport de cloture
- Comptabilite: journal financier, ecritures manuelles, filtres, resultat et export CSV
- Statistiques avec graphiques: ventes, top produits, paiements, valeur stock
- Rapports avec filtres par date, paiement, statut et recherche client/numero
- Export Microsoft Excel `.xls` et CSV
- Facturation professionnelle en FCFA avec informations fiscales
- Export FNE JSON de preparation
- Devis client avec impression et transformation en vente
- Bons de livraison lies aux ventes/factures
- Generation d'etiquettes produits en PDF A4, a l'unite ou par lot
- Parametres du commerce

## Devis et bons de livraison

Le module `Devis` permet de preparer une offre client avant vente:

- creation de devis avec client, validite, note et lignes produits;
- modification des lignes tant que le devis n'est pas converti;
- statuts `Brouillon`, `Envoye`, `Accepte`, `Refuse`, `Converti`;
- apercu et impression du devis;
- transformation en vente avec creation automatique de facture, mouvement de stock, caisse et credit client si paiement partiel.

Le module `Bons livraison` permet de generer un BL depuis une vente/facture:

- reprise automatique du client et des articles vendus;
- statuts `Prepare`, `En livraison`, `Livre`, `Annule`;
- apercu et impression du bon de livraison;
- signature client et magasin sur le document imprime.

## Comptabilite

Le module `Comptabilite` consolide les mouvements de caisse et les ecritures manuelles:

- recettes, depenses, resultat et solde especes;
- filtres par date, type et mode de paiement;
- categories de charges et recettes hors ventes;
- journal avec utilisateur et source;
- export CSV comptable.

Les ventes encaissees remontent via la caisse. Les charges, recettes hors ventes et ajustements peuvent etre saisis manuellement.

## Achats et retours

Le module `Achats` permet d'enregistrer un approvisionnement fournisseur:

- choix du fournisseur;
- choix du produit;
- quantite recue;
- cout unitaire;
- montant paye et mode de paiement;
- entree automatique en stock;
- mise a jour du cout d'achat;
- depense de caisse quand un montant est paye.

Le module `Retours` permet de traiter un retour client:

- recherche par numero de vente;
- choix de la ligne vendue;
- quantite retournee;
- montant rembourse;
- motif du retour;
- remise automatique en stock;
- sortie de caisse si remboursement.

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

Le module `Factures/FNE` centralise la gestion professionnelle des factures:

- filtres par periode, statut de facture, statut FNE et recherche client/numero/reference;
- apercu facture texte et generation PDF dans `data/factures/`;
- export FNE unitaire et export FNE par lot pour les factures actives;
- mise a jour du suivi FNE: `A exporter`, `Exportee`, `Certifiee`, `Rejetee`;
- saisie obligatoire de la reference FNE quand une facture est marquee `Certifiee`;
- annulation/restauration logique de facture avec PIN admin et motif, sans suppression fiscale.

Les ventes restent le point de creation des factures. La vue `Factures/FNE` sert au suivi, a la correction controlee, a l'impression et a la preparation fiscale.

## Etiquettes produits

Depuis le module `Produits`, selectionner un produit puis utiliser:

- `Etiquette produit`: genere une ou plusieurs etiquettes pour le produit selectionne;
- `Lot etiquettes A4`: genere une planche A4 pour tous les produits actifs.

Les fichiers PDF sont crees dans `data/etiquettes/` avec 24 etiquettes par page A4.

## Images et informations produit

Le formulaire produit contient des champs optionnels:

- Part Number;
- S/N;
- IMEI;
- modele;
- couleur;
- taille;
- image produit.

Les images importees sont copiees dans `data/uploads/produits/`.

## Lecteurs code-barres

Les lecteurs USB ou Bluetooth de toutes marques sont supportes quand ils fonctionnent en mode clavier HID, le mode standard de la plupart des scanners. Le lecteur doit envoyer le code puis `Entree` ou `Tab`.

Dans le POS tactile:

- le champ `Scan` peut recevoir directement le code;
- le scan est aussi capte globalement quand le curseur n'est pas dans un autre champ;
- la recherche se fait sur `Code barre`, `Part Number`, `S/N` et `IMEI`;
- le produit trouve est ajoute automatiquement au panier.

## Acces et securite

Les codes d'acces ne sont pas affiches sur la page de connexion.

Au premier lancement sur un poste Windows/EXE, l'application demande un PIN d'installation pour activer localement le logiciel:

- PIN d'installation: `05535350`

Une fois l'installation activee, ce PIN n'est plus redemande sur le meme poste sauf remise a zero de la base locale.

Comptes par defaut:

- Admin: PIN `0714`
- Vendeur: PIN `1234`

L'administrateur peut ajouter/modifier/desactiver les utilisateurs depuis `Parametres`.

## Propositions de fonctions a implementer ensuite

- permissions fines par role: magasinier, comptable, gerant;
- journal d'audit des actions sensibles;
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
