# Factur-x-generate
Outils de génération de Factur-X

L'idée est de générer un XML à partir des données d'une base de données ou un tableur générateur de factures en PDF,
puis de l'y insérer en utilisant Python et la bibliothèque Factur-X
au lieu de reconstituer le XML à partir du PDF

L'outil de génération de facture (tableur ou base de donnée), à l'aide d'un script intégré (cf. code ACCESS VBA):
remplace les valeurs dans une copie du modèle,
efface les lignes inutiles,
duplique les blocks si besoin (par exemple les lignes de facture),

puis lance le script Facture-X_Insert.py
