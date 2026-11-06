# language: fr
Fonctionnalité: Un type de ticket peut suivre son propre workflow dans un projet
  Dans Jira, le schéma de workflows d'un projet associe chaque type de ticket à un workflow : dans un même projet
  d'équipe, les Bugs peuvent suivre un autre workflow que les Stories. Aujourd'hui le classeur choisit le workflow
  par le rôle du projet seulement (Program = Feature, Team = Team), comme le dit
  docs/safe-flow-metrics-and-data-checks.md, section 1.

  À VENIR : scénarios écrits avant la vérification chez le client (Paramètres du projet > Workflows, ou
  GET /rest/api/3/project/{clé}/statuses). Ce fichier n'est relié à aucune étape (aucun appel scenarios()) :
  ni `test` ni `test-excel` ne le lancent. Le relier seulement si la vérification confirme le besoin.

  Le workflow « Bug » ci-dessous est un exemple de test, pas celui du client. Par défaut aucun type n'a de workflow
  propre : chaque projet garde le workflow de son rôle, et les chiffres d'aujourd'hui ne changent pas.

  Règle: La référence Python lit le workflow du type du ticket, et celui du rôle du projet sinon

    Contexte:
      Soit le workflow "Bug" décrit dans les paramètres :
        | Order | Status          | Category | IsStart | IsEnd | IsFree |
        | 1     | Nouveau         | Backlog  | 0       | 0     | 0      |
        | 2     | En analyse      | Wait     | 0       | 0     | 0      |
        | 3     | En correction   | Active   | 1       | 0     | 0      |
        | 4     | À livrer        | Wait     | 0       | 0     | 0      |
        | 5     | Corrigé         | Done     | 0       | 1     | 0      |
        | -     | Rejeté          | Removed  | 0       | 0     | 0      |
      Et ses transitions dans les paramètres :
        | From          | To            | Type    |
        | Nouveau       | En analyse    | forward |
        | En analyse    | En correction | forward |
        | En correction | À livrer      | forward |
        | À livrer      | Corrigé       | forward |
        | À livrer      | En correction | back    |
        | En analyse    | Rejeté        | exit    |
      Et le type "Bug" des projets d'équipe suit le workflow "Bug" dans les paramètres
      Et ces tickets, créés en statut Nouveau :
        | ticket | type    | équipe | créé  | changements de statut                                                                                 |
        | BUG-1  | Bug     | Alpha  | 10-01 | 10-02 En analyse ; 10-05 En correction ; 10-09 À livrer ; 10-10 Corrigé                               |
        | BUG-2  | Bug     | Alpha  | 10-01 | 10-02 En analyse ; 10-03 Corrigé                                                                      |
        | STO-1  | Story   | Alpha  | 10-01 | 10-02 En analyse ; 10-04 Prêt ; 10-05 Planifié ; 10-06 En développement ; 10-08 En test ; 10-09 À déployer ; 10-10 Terminé |
        | FEA-1  | Feature | ART    | 10-01 | 10-02 En analyse ; 10-20 Terminé                                                                      |

    Scénario: Un Bug suit le workflow de son type, pas celui de son projet
      Quand je calcule la ligne du ticket "BUG-1"
      Alors son workflow est "Bug"
      Et aucun de ses statuts n'est non classé
      Et son démarrage est le "10-05"
      Et son Flow Time vaut 5 jours
      Et ses jours actifs valent 4 et ses jours d'attente 1

    Scénario: Un même nom de statut peut avoir un sens différent selon le workflow
      # « En analyse » démarre le Flow Time dans le workflow Team, c'est une attente dans le workflow Bug.
      Quand je calcule les lignes des tickets "STO-1" et "BUG-1"
      Alors le démarrage de "STO-1" est le "10-02", à son entrée en « En analyse »
      Et le démarrage de "BUG-1" est le "10-05", après son passage en « En analyse »

    Plan du scénario: Un type sans workflow propre garde le workflow du rôle de son projet
      Quand je calcule la ligne du ticket "<ticket>"
      Alors son workflow est "<workflow>"
      Et son Flow Time vaut <jours> jours

      Exemples:
        | ticket | workflow | jours |
        | STO-1  | Team     | 8     |
        | FEA-1  | Feature  | 18    |

    Scénario: Le contrôle des passages (DC28) suit les transitions du workflow du ticket
      Quand je calcule les contrôles de données des tickets
      Alors le contrôle "DC28" signale "BUG-2"
      Et le contrôle "DC28" ne signale pas "BUG-1"
      Et le contrôle "DC08" ne signale ni "BUG-1" ni "BUG-2"

  Règle: Les paramètres refusent une association incohérente et ne changent rien par défaut

    Scénario: Par défaut aucun type n'a de workflow propre
      Soit les fichiers de paramètres du dépôt
      Quand je les charge
      Alors aucun type n'a de workflow propre

    Scénario: Un type associé à un workflow absent de la feuille Workflows est refusé
      Soit les fichiers de paramètres du dépôt
      Quand j'associe le type "Bug" des projets d'équipe au workflow "Correctif", absent de "workflows.md"
      Et je les charge
      Alors le chargement est refusé avec "workflow 'Correctif'"

  Règle: Le classeur livré calcule chaque ticket avec le workflow de son type, comme la référence Python

    @excel
    Scénario: Les Bugs des projets d'équipe suivent le workflow Bug dans le classeur livré
      Soit un faux Jira avec 150 tickets générés dont les Bugs des projets d'équipe suivent le workflow "Bug"
      Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
      Et le workflow "Bug", ses transitions et l'association du type "Bug" sont dans les feuilles de paramètres du classeur
      Et les itérations de 14 jours depuis le "2026-08-17" jusqu'à la date de référence dans la feuille Iterations
      Quand j'exécute la macro Extraire et calculer du classeur
      Alors aucune cellule des feuilles de flux ne montre une erreur Excel
      Et chaque Bug a le workflow "Bug" dans les feuilles StatusIntervals et ItemFlow
      Et aucun statut d'un Bug n'est non classé
      Et la feuille StatusIntervals reproduit les intervalles de la référence Python
      Et la table ItemFlow reproduit les calculs de la référence Python
      Et la table Metrics reproduit la référence Python en mode "A"
      Et la table DataChecks et les colonnes DC de ItemFlow reproduisent les contrôles de la référence Python
