# language: fr
Fonctionnalité: Le classeur livré extrait Jira et calcule les mêmes chiffres que la référence Python
  Chemin réel sur le classeur à livrer (sans Python) : faux Jira, macro Extraire et calculer (extraction VBA, tables
  et formules Excel construites par le VBA depuis les feuilles FlowFormulas et ChartSpec, graphiques), puis
  comparaison cellule par cellule avec le calcul de référence en Python sur les mêmes données. Les paramètres (mode,
  date de référence, projets, itérations) sont lus dans le classeur.

  @excel
  Scénario: Mêmes résultats en mode A puis en mode B après changement du paramètre dans Config
    Soit un faux Jira avec 250 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les itérations de 14 jours depuis le "2026-06-01" jusqu'à la date de référence dans la feuille Periods
    Quand j'exécute la macro Extraire et calculer du classeur
    Alors aucune cellule des feuilles de flux ne montre une erreur Excel
    Et la feuille StatusIntervals reproduit les intervalles de la référence Python
    Et la table ItemFlow reproduit les calculs de la référence Python
    Et la table DailyLoad reproduit la référence Python en mode "A"
    Et la table Metrics reproduit la référence Python en mode "A"
    Et la feuille ChartData reproduit les métriques de la référence Python en mode "A", avec #N/A pour « not computable »
    Et chaque graphique a une série par équipe ou par type de travail avec les valeurs de ChartData
    Et le graphique Flow Distribution suit l'équipe choisie dans la cellule B2
    Et les graphiques sont exportés en PNG
    Quand je passe FlowLoadMode à "B" dans la feuille Config
    Alors la table DailyLoad reproduit la référence Python en mode "B"
    Et la table Metrics reproduit la référence Python en mode "B"
    Et la feuille ChartData reproduit les métriques de la référence Python en mode "B", avec #N/A pour « not computable »

  @excel
  Scénario: Les contrôles de données DC01, DC04 à DC10 et DC28 trouvent les anomalies connues et rien d'autre
    Soit un faux Jira avec 150 tickets générés et les anomalies de données connues
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les itérations de 14 jours depuis le "2026-08-17" jusqu'à la date de référence dans la feuille Periods
    Quand j'exécute la macro Extraire et calculer du classeur
    Alors aucune cellule des feuilles de flux ne montre une erreur Excel
    Et la feuille StatusIntervals reproduit les intervalles de la référence Python
    Et la table ItemFlow reproduit les calculs de la référence Python
    Et la table DataChecks et les colonnes DC de ItemFlow reproduisent les contrôles de la référence Python
    Et chaque anomalie connue est trouvée par les contrôles qui lui correspondent et seulement par eux

  @excel
  Scénario: Un second calcul avec moins de données remplace le premier sans rien laisser
    Soit un faux Jira avec 120 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les itérations de 14 jours depuis le "2026-08-17" jusqu'à la date de référence dans la feuille Periods
    Quand j'exécute la macro Extraire et calculer du classeur
    Alors aucune cellule des feuilles de flux ne montre une erreur Excel
    Et les feuilles de flux ont exactement la taille des nouvelles données
    Et la feuille StatusIntervals reproduit les intervalles de la référence Python
    Et la table ItemFlow reproduit les calculs de la référence Python
    Et la table DailyLoad reproduit la référence Python en mode "A"
    Et la table Metrics reproduit la référence Python en mode "A"
    Et la feuille ChartData reproduit les métriques de la référence Python en mode "A", avec #N/A pour « not computable »
    Et chaque graphique a une série par équipe ou par type de travail avec les valeurs de ChartData
    Et le graphique Flow Distribution suit l'équipe choisie dans la cellule B2
