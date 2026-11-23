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
    Et la table FlowTimeCheck de l'ART est celle calculée à la main

  @excel
  Scénario: Les statuts viennent de Jira : la feuille Workflows ne déclare que les trois statuts d'attente et les chiffres ne changent pas
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Et la feuille Workflows ne garde que les trois statuts d'attente comme exceptions
    Quand j'extrais et calcule
    Alors la vélocité par itération de l'ART est celle du scénario
    Et chaque itération qui livre a un Flow Time de 8 jours et le Flow Efficiency de son PI
    Et chaque statut dure en moyenne 1 jour, 2 jours pour le statut long du PI de la Feature
    Et la feuille StatusRules donne la catégorie, le début et la fin de chaque statut du scénario avec leur source

  @excel
  Scénario: Sans aucune exception tout vient de Jira : le Flow Efficiency n'est pas calculable faute de statut d'attente, le reste ne change pas
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Et la feuille Workflows est vide
    Quand j'extrais et calcule
    Alors la vélocité par itération de l'ART est celle du scénario
    Et chaque itération qui livre a un Flow Time de 8 jours et un Flow Efficiency non calculable
    Et la feuille StatusRules ne contient que des règles venues de Jira, sans statut d'attente

  @excel
  Scénario: Le classeur ne contient que les données du scénario, aucune équipe ni aucun workflow des paramètres par défaut
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors la feuille Projects ne contient que le projet du programme, la feuille Workflows que les trois statuts Wait, la feuille Transitions rien, et les feuilles lues dans Jira que le workflow du programme
    Et chaque graphique par équipe de la feuille Charts n'a que la série ART

  @excel
  Scénario: Le contrôle DC28 juge les passages d'après le workflow Jira du ticket quand la feuille Transitions est vide
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Alors aucun passage n'est hors du workflow Jira de son ticket
    Et un passage retiré du workflow Jira lève DC28 sur les Features qui l'ont pris et le remettre l'efface

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
    Et chaque ticket est dans un seul compte de ItemFlow et les comptes sont ceux calculés à la main
    Et les listes de la feuille MetricItems sont celles calculées à la main et leur nombre de lignes est la vélocité
    Et tous les contrôles de la feuille Reconciliation sont OK avec les chiffres calculés à la main et le recensement des types est le bon

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

  @excel
  Scénario: Sans statut de départ ni statut Wait, Flow Time, Flow Efficiency et Flow Load ne sont pas calculables et le classeur dit pourquoi
    Soit le faux Jira du scénario art-metrics
    Et le faux Jira refuse la lecture des workflows
    Et le classeur à livrer rempli pour ce scénario
    Et la feuille Workflows est vide
    Quand j'extrais et calcule
    Alors la vélocité par itération de l'ART est celle du scénario
    Et Flow Time, Flow Efficiency et Flow Load de l'ART ne sont calculables dans aucune itération
    Et chaque Feature livrée est hors du Flow Time faute de statut de départ dans les règles
    Et la table FlowTimeCheck de l'ART est celle calculée à la main, sans statut de départ
    Et le journal et les champs de santé disent pourquoi, avec les codes calculés à la main
    Et la Reconciliation demande de revoir le statut de départ et le statut Wait

  @excel
  Scénario: Corriger la feuille Workflows puis recalculer sans relire Jira : le Flow Efficiency suit la correction, dans les deux sens
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Quand j'extrais et calcule
    Et je vide la feuille Workflows puis je calcule et je dessine
    Alors chaque itération qui livre a un Flow Time de 8 jours et un Flow Efficiency non calculable
    Et la feuille StatusRules ne contient que des règles venues de Jira, sans statut d'attente
    Et le faux Jira n'a reçu aucune requête pendant le calcul et le dessin
    Quand je remets les trois statuts d'attente puis je calcule et je dessine
    Alors chaque itération qui livre a un Flow Time de 8 jours et le Flow Efficiency de son PI
    Et le faux Jira n'a reçu aucune requête pendant le calcul et le dessin
