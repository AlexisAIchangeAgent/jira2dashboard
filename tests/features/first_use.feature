# language: fr
Fonctionnalité: Première utilisation du classeur
  La première fois, l'utilisateur doit savoir dans quel ordre cliquer sur les boutons : le Dashboard le dit, et
  l'onglet Manual explique le reste. La feuille Config dit quels paramètres sont obligatoires (colonne Mandatory) :
  un paramètre obligatoire vide arrête le bouton avant tout appel à Jira, avec un message qui le nomme.

  @excel
  Scénario: Le Dashboard donne l'ordre des boutons et l'onglet Manual suit le Dashboard
    Soit le classeur à livrer
    Alors le Dashboard donne les premières étapes, avec « Read the Jira configuration » avant « 1. Extract from Jira »
    Et l'onglet Manual suit le Dashboard et reproduit le manuel
    Et la feuille Config a les colonnes Example et Mandatory

  @excel
  Scénario: Chaque feuille du classeur dit d'où viennent ses données
    Soit le classeur à livrer
    Alors chaque feuille du classeur porte le suffixe de sa source
    Et les feuilles à saisir sont _from_user, celles lues dans Jira sont _from_Jira et celles calculées sont _from_excel

  @excel
  Scénario: Un paramètre obligatoire vide arrête l'extraction avant tout appel à Jira
    Soit un faux Jira avec 30 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Quand je vide le paramètre "BlocksLinkType" de la feuille Config
    Et je clique sur le bouton « 1. Extract from Jira »
    Alors l'état du tableau de bord nomme les paramètres obligatoires vides "BlocksLinkType"
    Et le faux Jira n'a reçu aucune requête

  @excel
  Scénario: Un en-tête faux de la feuille Iterations arrête le calcul et le Dashboard dit quel en-tête taper
    Soit un faux Jira avec 30 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Quand l'en-tête de la colonne des libellés de la feuille Iterations est "Iteration label"
    Et je clique sur le bouton « 2. Compute the metrics »
    Alors l'état du tableau de bord donne le code J2D-CFG-014 et, sur une seule ligne, le message qui nomme l'en-tête à taper
    Et le faux Jira n'a reçu aucune requête

  @excel
  Scénario: Un classeur verrouillé arrête le bouton avant tout appel à Jira et dit comment le déverrouiller
    Soit un faux Jira avec 30 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et le classeur est rouvert en lecture seule
    Quand je clique sur le bouton Lire la configuration Jira
    Alors l'état du tableau de bord dit que le classeur est verrouillé et qu'il faut activer la modification
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
