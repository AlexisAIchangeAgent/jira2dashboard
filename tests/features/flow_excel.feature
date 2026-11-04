# language: fr
Fonctionnalité: Les formules Excel donnent les mêmes chiffres que la référence Python
  Chemin réel : faux Jira, extraction VBA (feuilles Issues et StatusIntervals), formules Excel (ItemFlow,
  DailyLoad, Metrics), puis comparaison cellule par cellule avec le calcul de référence en Python sur les
  mêmes données. Les paramètres (mode, date de référence, projets, itérations) sont lus dans le classeur.

  @excel
  Scénario: Mêmes résultats en mode A puis en mode B après changement du paramètre dans Config
    Soit un faux Jira avec 250 tickets générés avec les projets des paramètres
    Et un classeur dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les itérations de 14 jours depuis le "2026-06-01" jusqu'à la date de référence dans la feuille Periods
    Quand le VBA extrait les données de flux depuis le faux Jira
    Et les feuilles de flux sont construites avec des formules
    Alors aucune cellule des feuilles de flux ne montre une erreur Excel
    Et la feuille StatusIntervals reproduit les intervalles de la référence Python
    Et la table ItemFlow reproduit les calculs de la référence Python
    Et la table DailyLoad reproduit la référence Python en mode "A"
    Et la table Metrics reproduit la référence Python en mode "A"
    Quand le tableau de bord est construit
    Alors la feuille ChartData reproduit les métriques de la référence Python en mode "A", avec #N/A pour « not computable »
    Et chaque graphique a une série par équipe ou par type de travail avec les valeurs de ChartData
    Et le graphique Flow Distribution suit l'équipe choisie dans la cellule B2
    Et les graphiques sont exportés en PNG
    Quand je passe FlowLoadMode à "B" dans la feuille Config
    Alors la table DailyLoad reproduit la référence Python en mode "B"
    Et la table Metrics reproduit la référence Python en mode "B"
    Et la feuille ChartData reproduit les métriques de la référence Python en mode "B", avec #N/A pour « not computable »
