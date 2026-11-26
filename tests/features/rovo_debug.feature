# language: fr
Fonctionnalité: Le bouton « Debug prompt for Rovo » écrit ce qu'il faut à Rovo pour comparer les chiffres du classeur avec Jira
  Le bouton écrit des fichiers 01.md, 02.md... dans le dossier rovo-debug à côté du classeur : les consignes, le contexte et les
  contrôles de la Reconciliation, puis les tickets de chaque itération (au plus les 18 dernières commencées) avec les requêtes JQL qui
  ne ramènent que les écarts. Rovo ne compte pas de façon fiable : il lance ces requêtes telles quelles et liste des clés, chacune avec
  une cause d'une liste fermée. Rien n'est envoyé : l'utilisateur dépose les fichiers dans Rovo Chat. Les clés attendues viennent de
  tests/art_oracle.py, calculé à la main (AL, 2026-11-26).

  @excel
  Scénario: Après le calcul, les fichiers donnent les tickets de chaque itération et des requêtes JQL valides, sans rien de la connexion
    Soit le faux Jira du scénario art-metrics
    Et le classeur à livrer rempli pour ce scénario
    Et un ancien fichier 07.md et une réponse de Rovo enregistrée dans le dossier rovo-debug
    Quand j'extrais et calcule
    Et le paramètre Filter de la feuille Config vaut « labels != "équipe" »
    Et je clique sur le bouton Debug prompt for Rovo
    Alors le dossier rovo-debug a les nouveaux fichiers 01.md, 02.md..., chacun de 12 000 caractères au plus en UTF-8, plus d'ancien fichier et la réponse de Rovo gardée
    Et les fichiers donnent les 18 dernières itérations, aucune de IT-1, chacune avec les clés livrées calculées à la main
    Et les sous-tâches TTT-26 et TTT-27 sont laissées de côté dans IT-2-1 et TTT-28 dans IT-3-4, comme des tickets qui ne sont pas des Features
    Et une itération sans ticket livré a des requêtes JQL sans liste vide
    Et le premier fichier donne les consignes, le Filter, les statuts de fin, les contrôles et l'avertissement sur le fuseau horaire, sans aucun code de contrôle
    Et les fichiers ne contiennent ni le jeton, ni l'e-mail, ni l'adresse de Jira
    Et l'état du tableau de bord dit combien de fichiers sont écrits et lequel déposer d'abord
    Et le journal local décrit le run du bouton "Debug prompt for Rovo", avec ses paramètres en clair dont le Filter, sans aucune clé de ticket

  @excel
  Scénario: Avant 2. Compute the metrics, le bouton s'arrête et dit de calculer d'abord
    Soit le classeur à livrer
    Quand je clique sur le bouton Debug prompt for Rovo
    Alors l'état du tableau de bord donne le code J2D-ROVO-004 et dit de cliquer sur 2. Compute the metrics

  @excel
  Scénario: Sur les données générées, chaque équipe a ses requêtes JQL sur ses projets avec les clés livrées, et un fichier trop gros est signalé
    Soit un faux Jira avec 250 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les itérations de 14 jours depuis le "2026-06-01" jusqu'à la date de référence dans la feuille Iterations
    Et l'équipe du projet AAA s'appelle « Alpha » dans la feuille Projects
    Quand j'exécute la macro Extraire et calculer du classeur
    Et la limite des fichiers de la feuille RovoDebugSpec passe à 3000 caractères
    Et je clique sur le bouton Debug prompt for Rovo
    Alors chaque équipe de chaque itération a ses requêtes JQL sur ses projets, avec les clés livrées de la liste Done de MetricItems
    Et l'état du tableau de bord signale le fichier qui dépasse la limite de 3000 caractères
