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
