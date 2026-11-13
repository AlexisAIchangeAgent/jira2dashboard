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
    Et les itérations de 14 jours depuis le "2026-06-01" jusqu'à la date de référence dans la feuille Iterations
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
    Et la feuille Changelog a une ligne par changement du faux Jira, sans auteur
    Et chaque feuille du classeur porte le suffixe de sa source
    Quand je passe FlowLoadMode à "B" dans la feuille Config
    Alors la table DailyLoad reproduit la référence Python en mode "B"
    Et la table Metrics reproduit la référence Python en mode "B"
    Et la feuille ChartData reproduit les métriques de la référence Python en mode "B", avec #N/A pour « not computable »

  @excel
  Scénario: Les contrôles de données trouvent les anomalies connues et rien d'autre
    Soit un faux Jira avec 150 tickets générés et les anomalies de données connues
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les champs Jira des contrôles et la fenêtre du PI sont renseignés dans les feuilles Config et PIs
    Et les itérations de 14 jours depuis le "2026-08-17" jusqu'à la date de référence dans la feuille Iterations
    Quand j'exécute la macro Extraire et calculer du classeur
    Alors aucune cellule des feuilles de flux ne montre une erreur Excel
    Et la table ItemFlow reproduit les calculs de la référence Python
    Et la table FillRates reproduit les taux de remplissage de la référence Python
    Et la table DataChecks et les colonnes DC de ItemFlow reproduisent les contrôles de la référence Python
    Et chaque anomalie connue est trouvée par les contrôles qui lui correspondent et seulement par eux
    Et la table PIMetrics et le graphique Feature Predictability donnent pour chaque PI les Features faites avant sa fin sur les Features planifiées

  @excel
  Scénario: Un contrôle dont le champ Jira n'est pas configuré est « not computable » et ne signale personne
    Soit un faux Jira avec 150 tickets générés et les anomalies de données connues
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les champs Jira des contrôles et la feuille PIs sont vides
    Et les itérations de 14 jours depuis le "2026-08-17" jusqu'à la date de référence dans la feuille Iterations
    Quand j'exécute la macro Extraire et calculer du classeur
    Alors aucune cellule des feuilles de flux ne montre une erreur Excel
    Et la table FillRates reproduit les taux de remplissage de la référence Python
    Et la table DataChecks et les colonnes DC de ItemFlow reproduisent les contrôles de la référence Python
    Et les contrôles qui lisent un champ non configuré ne signalent aucun ticket

  @excel
  Scénario: Un second calcul avec moins de données remplace le premier sans rien laisser
    Soit un faux Jira avec 120 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les itérations de 14 jours depuis le "2026-08-17" jusqu'à la date de référence dans la feuille Iterations
    Et un premier calcul sur un faux Jira de 250 tickets a déjà rempli le classeur
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

  @excel
  Scénario: Le bouton extrait par tranches : la barre avance, Excel répond entre deux tranches, puis les résultats sont complets
    Soit un faux Jira lent avec 150 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les itérations de 14 jours depuis le "2026-08-17" jusqu'à la date de référence dans la feuille Iterations
    Quand je clique sur le bouton Extraire et calculer sans attendre la fin
    Alors la barre de progression avance sans reculer, le classeur répond entre deux tranches et l'extraction se termine
    Et le journal local décrit l'exécution étape par étape, avec ses paramètres et le temps de chaque étape, sans jeton, adresse ni réponse de Jira
    Et la feuille Steps montre chaque étape avec son statut et son temps
    Et aucune cellule des feuilles de flux ne montre une erreur Excel
    Et la table ItemFlow reproduit les calculs de la référence Python
    Et la table Metrics reproduit la référence Python en mode "A"
    Et l'état du tableau de bord ne signale ni ticket sans équipe, ni statut inconnu, ni cellule en erreur

  @excel
  Scénario: Annuler pendant l'extraction garde les résultats du calcul précédent
    Soit un faux Jira lent avec 150 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les itérations de 14 jours depuis le "2026-08-17" jusqu'à la date de référence dans la feuille Iterations
    Quand j'exécute la macro Extraire et calculer du classeur
    Et je clique sur le bouton Extraire et calculer sans attendre la fin
    Et j'annule dès que la barre a avancé
    Alors les résultats du premier calcul sont conservés et l'état dit « Cancelled »

  @excel
  Scénario: L'état de l'extraction dit pourquoi les graphiques sont vides quand les itérations ne couvrent pas les tickets terminés
    Soit un faux Jira avec 120 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les itérations de 14 jours depuis le "2026-08-17" jusqu'à la date de référence dans la feuille Iterations
    Et seule la première itération reste dans la feuille Iterations
    Quand je clique sur le bouton Extraire et calculer sans attendre la fin
    Et j'attends la fin de l'extraction du bouton
    Alors l'état du tableau de bord dit combien de tickets terminés sont hors des itérations
