# language: fr
Fonctionnalité: Le scénario art-metrics du faux Jira donne, par le chemin Python, les métriques calculées à la main
  Le scénario art-metrics : 25 Features, 4 PI de 6 itérations, 3 équipes, des Features reportées (champ PI vidé ou déplacé) et une
  Feature ajoutée en cours de PI. Chaque chiffre attendu est calculé à la main à partir du texte du scénario (tests/art_oracle.py),
  jamais par le code testé. La Feature Predictability d'un PI se juge sur ce qui était engagé à son début.

  Scénario: Chaque PI est jugé sur ce qui était engagé à son début, mode A
    Soit le faux Jira du scénario art-metrics
    Quand je lance l'extraction Python sur ce scénario
    Alors elle se termine avec le code 0
    Et la feuille PIMetrics donne pour les 4 PI 6 prévues et 6, 5, 3 et 4 faites
    Et la Feature Predictability des PI est 100 %, 83,33 %, 50 % et 66,67 %

  Scénario: Le mode B compte aussi la Feature ajoutée en cours de PI
    Soit le faux Jira du scénario art-metrics
    Et le paramètre "PredictabilityMode" vaut "B" dans le classeur de paramètres
    Quand je lance l'extraction Python sur ce scénario
    Alors elle se termine avec le code 0
    Et la feuille PIMetrics donne pour les 4 PI 6 prévues et 6, 5, 3 et 5 faites

  Scénario: La table IterationFeatures est celle calculée à la main, pour l'ART et pour chaque équipe
    Soit le faux Jira du scénario art-metrics
    Quand je lance l'extraction Python sur ce scénario
    Alors elle se termine avec le code 0
    Et la feuille IterationFeatures de l'ART est celle calculée à la main
    Et chaque équipe a ses lignes de PI-3 et PI-4 calculées à la main

  Scénario: La vélocité, le Flow Time et le Flow Efficiency de l'ART sont ceux calculés à la main
    Soit le faux Jira du scénario art-metrics
    Quand je lance l'extraction Python sur ce scénario
    Alors elle se termine avec le code 0
    Et la vélocité par itération est celle du scénario
    Et chaque itération qui livre a un Flow Time de 8 jours et le Flow Efficiency de son PI
    Et chaque statut dure en moyenne 1 jour, 2 jours pour le statut long du PI de la Feature

  Scénario: Sans statut de départ ni statut Wait, Flow Time, Flow Efficiency et Flow Load ne sont pas calculables et le résultat dit pourquoi
    Soit le faux Jira du scénario art-metrics
    Et le faux Jira refuse la lecture des workflows
    Et la feuille Workflows du classeur de paramètres est vide
    Quand je lance l'extraction Python sur ce scénario
    Alors elle se termine avec le code 0
    Et la vélocité par itération est celle du scénario
    Et Flow Time, Flow Efficiency et Flow Load de l'ART ne sont calculables dans aucune itération
    Et chaque Feature livrée est hors du Flow Time faute de statut de départ dans les règles
    Et le journal et la feuille Run disent pourquoi, avec les codes calculés à la main
    Et la Reconciliation demande de revoir le statut de départ et le statut Wait
