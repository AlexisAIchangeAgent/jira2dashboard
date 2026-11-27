# language: fr
Fonctionnalité: Le bouton « VBA code for Rovo » écrit tout le code du classeur dans un seul fichier texte
  Rovo lit un xlsx mais pas les macros d'un xlsm (AL, 2026-11-27) : le bouton écrit vba-code.txt dans le dossier rovo-debug à côté du
  classeur, avec le code de chaque module livré et la formule de chaque colonne des tables, pour que Rovo lise le code derrière un
  chiffre. Le code vient de la feuille cachée VbaCode, écrite au build depuis vba\*.bas : le classeur ne lit jamais son propre projet
  VBA. Le fichier attendu est vba_code.file_text, calculé en Python depuis les mêmes fichiers. Il remplace le Debug prompt for Rovo.

  @excel
  Scénario: Le bouton écrit le code des modules livrés et les formules des tables, sans rien de la connexion
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Et un ancien fichier 07.md et une réponse de Rovo enregistrée dans le dossier rovo-debug
    Quand je clique sur le bouton VBA code for Rovo
    Alors le fichier vba-code.txt du dossier rovo-debug est le code des modules livrés et les formules des tables, caractère par caractère
    Et l'ancien fichier 07.md est effacé et la réponse de Rovo est gardée
    Et la feuille VbaCode est cachée
    Et l'état du tableau de bord donne le fichier, ses modules, ses lignes et le dossier
    Et le journal local décrit le run du bouton "VBA code for Rovo", sans aucun chemin
