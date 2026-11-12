# language: fr
Fonctionnalité: Le bouton « Prompt for Rovo » prépare le texte à donner à Rovo pour trouver les valeurs à saisir
  Le bouton écrit la feuille RovoPrompt : un prompt qui demande à Rovo toutes les valeurs à renseigner dans le classeur
  (connexion, métriques de flux, projets, itérations, PI, workflows, transitions) et une réponse en tableau exportable en
  CSV. Les paramètres de Config sont lus dans le classeur, mais jamais leur valeur : le texte ne contient ni adresse, ni
  e-mail, ni jeton. Rien n'est envoyé : l'utilisateur copie le texte lui-même.

  @excel
  Scénario: Le prompt liste chaque paramètre de Config, chaque feuille à remplir et le format CSV, sans aucune valeur saisie
    Soit un faux Jira avec 120 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Quand je clique sur le bouton Prompt for Rovo
    Alors la feuille RovoPrompt liste chaque paramètre de la feuille Config avec son caractère obligatoire
    Et la feuille RovoPrompt nomme chaque feuille à remplir avec ses colonnes
    Et la feuille RovoPrompt demande une réponse en CSV avec les colonnes Sheet, Row, Column, Value et Note
    Et la feuille RovoPrompt ne contient aucune valeur saisie dans la feuille Config
    Et l'état du tableau de bord dit combien de lignes sont sélectionnées pour Rovo
    Et le journal local décrit le run du bouton "Prompt for Rovo", avec ses paramètres et le temps de son étape

  @excel
  Scénario: Le CSV de Rovo remplit la feuille Config sans jamais toucher à la connexion et remplace les lignes des feuilles qu'il contient
    Soit un faux Jira avec 120 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et une réponse de Rovo en CSV qui donne FillRateThreshold, une adresse de connexion, deux projets, une itération, un jeton à saisir et une valeur inconnue
    Quand je charge ce CSV avec le bouton Load settings CSV
    Alors la feuille Config a reçu FillRateThreshold et aucun paramètre de connexion
    Et la feuille Projects contient les deux lignes du CSV et plus aucune autre
    Et la feuille Iterations contient la ligne du CSV et garde sa colonne Example
    Et les feuilles Workflows et Transitions n'ont pas changé
    Et l'état du tableau de bord résume ce qui est chargé et les valeurs inconnues
    Et le journal local décrit le run du bouton "Load settings CSV", avec ses paramètres et le temps de son étape

  @excel
  Scénario: Le CSV de configuration du faux Jira remplit le classeur livré comme la copie de manual-test, sauf la connexion
    Soit le classeur à livrer
    Et le CSV de configuration du faux Jira, dans un dossier à part
    Quand je charge ce CSV avec le bouton Load settings CSV
    Alors la feuille Config a reçu la date de référence et les champs du faux Jira, et aucun paramètre de connexion
    Et les feuilles Iterations et PIs ont les itérations et le PI du faux Jira
