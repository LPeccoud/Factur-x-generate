# Factur-x-generate
Outils de génération de Factur-X pour tableur ou base de donnée

L'idée est:

(au lieu de reconstituer le XML à partir d'un PDF existant)
Partant d'une base de données (ou un tableur):

	générer une facture en PDF,
  
  générer un XML à partir des données de la base,
  
  insérer dans le PDF en utilisant Python et la bibliothèque Factur-X
  

Le tableur ou la base de donnée, à l'aide d'un script intégré (cf. code ACCESS VBA):

  remplace les valeurs dans une copie du modèle,
  
  efface les lignes inutiles,
  
  duplique les blocks si besoin (par exemple les lignes de facture),
  
  puis lance le script Facture-X_Insert.py
