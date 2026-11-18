# language: fr
Fonctionnalité: Le scénario art-metrics du faux Jira donne, dans le classeur livré, les métriques calculées à la main
  Le classeur (VBA et formules Excel) lit le faux Jira du scénario art-metrics : 25 Features, 4 PI, 3 équipes, des reports et un
  ajout en cours de PI. Chaque chiffre attendu vient de tests/art_oracle.py, calculé à la main ; les tables sont aussi comparées,
  cellule par cellule, au calcul de référence en Python.

  @excel
  Scénario: Chaque PI est jugé sur ce qui était engagé à son début, puis le mode B compte aussi la Feature ajoutée
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors la feuille PIMetrics donne pour les 4 PI 6 prévues et 6, 5, 3 et 4 faites
    Et la Feature Predictability des PI est 100 %, 83,33 %, 50 % et 66,67 %
    Quand je passe le paramètre PredictabilityMode à "B"
    Alors la feuille PIMetrics donne pour les 4 PI 6 prévues et 6, 5, 3 et 5 faites

  @excel
  Scénario: La table IterationFeatures est celle calculée à la main et celle du calcul de référence
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors la table IterationFeatures de l'ART est celle calculée à la main
    Et chaque équipe a ses lignes de PI-3 et PI-4 calculées à la main
    Et la table IterationFeatures et la table PIMetrics sont celles du calcul de référence

  @excel
  Scénario: La vélocité, le Flow Time, le Flow Efficiency et la durée des statuts sont ceux calculés à la main
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors la vélocité par itération de l'ART est celle du scénario
    Et chaque itération qui livre a un Flow Time de 8 jours et le Flow Efficiency de son PI
    Et chaque statut dure en moyenne 1 jour, 2 jours pour le statut long du PI de la Feature

  @excel
  Scénario: Le classeur ne contient que les données du scénario, aucune équipe ni aucun workflow des paramètres par défaut
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors la feuille Projects ne contient que le projet du programme, et les feuilles Workflows et Transitions que le workflow Feature
    Et chaque graphique par équipe de la feuille Charts n'a que la série ART

  @excel
  Scénario: Le temps de cycle par statut de chaque PI est celui calculé à la main et celui du calcul de référence
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors la table PIStatusTime donne pour chaque PI ses Features livrées, la durée moyenne et la part de chaque statut calculées à la main
    Et la table PIStatusTime est celle du calcul de référence et ses totaux sont les jours actifs et d'attente des mêmes Features
    Et le graphique Cycle time by status empile les parts des statuts de chaque PI

  @excel
  Scénario: La table VelocityCheck retrouve la vélocité de l'ART dans les changements d'état du changelog
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors la table VelocityCheck donne pour chaque itération de l'ART autant de passages vers l'état de fin que de Features livrées, sans création ni réouverture
    Et la table VelocityCheck est celle du calcul de référence et son écart est 0
    Et les sous-tâches du Program sont dans l'équipe « ART (other types) » et aucun chiffre de l'ART ne les compte

  @excel
  Scénario: Une itération sans livraison montre sa vélocité à 0 et la note sous les graphiques dit pourquoi
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors le graphique Velocity écrit la valeur au-dessus des barres de l'ART et sa note compte 6 points à 0
    Et les graphiques Flow Time et Flow Efficiency ont 6 trous expliqués par l'absence de ticket terminé

  @excel
  Scénario: Les 5 graphiques de la feuille IterationCharts montrent les chiffres de la table
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors la feuille IterationCharts a ses 5 graphiques avec une série pour l'ART et une par équipe
    Et chaque graphique montre pour l'ART les chiffres calculés à la main
