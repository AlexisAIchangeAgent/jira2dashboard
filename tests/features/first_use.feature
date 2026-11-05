# language: fr
Fonctionnalité: Première utilisation du classeur
  La première fois, l'utilisateur doit savoir dans quel ordre cliquer sur les boutons : le Dashboard le dit, et
  l'onglet Manual explique le reste. La feuille Config dit quels paramètres sont obligatoires (colonne Mandatory) :
  un paramètre obligatoire vide arrête le bouton avant tout appel à Jira, avec un message qui le nomme.

  @excel
  Scénario: Le Dashboard donne l'ordre des boutons et l'onglet Manual suit le Dashboard
    Soit le classeur à livrer
    Alors le Dashboard donne les premières étapes, avec « Read the Jira configuration » avant « Extract from Jira and compute »
    Et l'onglet Manual suit le Dashboard et reproduit le manuel
    Et la feuille Config a la colonne Mandatory

  @excel
  Scénario: Un paramètre obligatoire vide arrête l'extraction avant tout appel à Jira
    Soit un faux Jira avec 30 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Quand je vide le paramètre "FeatureTypes" de la feuille Config
    Et je clique sur le bouton Extraire et calculer
    Alors l'état du tableau de bord nomme les paramètres obligatoires vides "FeatureTypes"
    Et le faux Jira n'a reçu aucune requête

  @excel
  Scénario: Le bouton Lire la configuration Jira ne demande que les paramètres de connexion
    Soit un faux Jira avec 30 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Quand je vide le paramètre "AsOf" de la feuille Config
    Et je vide le paramètre "Email" de la feuille Config
    Et je clique sur le bouton Lire la configuration Jira
    Alors l'état du tableau de bord nomme les paramètres obligatoires vides "Email"
    Et le faux Jira n'a reçu aucune requête
