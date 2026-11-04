# language: fr
Fonctionnalité: Paramètres écrits dans le classeur
  Les statuts, les transitions et les champs ne sont jamais du code : ils sont lus dans params/*.md,
  écrits dans les feuilles du classeur, et le VBA les lit.

  Scénario: Les fichiers de paramètres décrivent deux workflows complets
    Soit les fichiers de paramètres du dépôt
    Quand je les charge
    Alors la feuille "Workflows" a 20 lignes
    Et la feuille "Transitions" a 26 lignes
    Et la feuille "FieldMapping" a 22 lignes

  Scénario: Une catégorie inconnue est refusée avec un message clair
    Soit les fichiers de paramètres du dépôt
    Quand je remplace "| Prêt | Wait |" par "| Prêt | Attente |" dans "workflows.md"
    Et je les charge
    Alors le chargement est refusé avec "Category 'Attente'"

  Scénario: Le token d'API ne peut pas être gardé dans le dépôt
    Soit les fichiers de paramètres du dépôt
    Quand je remplace "| ApiToken |  |" par "| ApiToken | abc123 |" dans "config.md"
    Et je les charge
    Alors le chargement est refusé avec "ApiToken must stay empty"

  Scénario: Les paramètres de connexion, de proxy et de filtre sont dans la feuille Config
    Soit les fichiers de paramètres du dépôt
    Quand je les charge
    Alors la feuille "Config" a 12 lignes

  Scénario: Les projets à extraire et le rôle de chacun sont dans la feuille Projects
    Soit les fichiers de paramètres du dépôt
    Quand je les charge
    Alors la feuille "Projects" a 9 lignes

  Scénario: Deux projets Program sont refusés
    Soit les fichiers de paramètres du dépôt
    Quand je remplace "| AAA | Team | AAA |" par "| AAA | Program |  |" dans "projects.md"
    Et je les charge
    Alors le chargement est refusé avec "exactly one Program project is needed, found 2"

  @excel
  Plan du scénario: Le classeur contient les paramètres du dépôt
    Soit le classeur construit avec les paramètres du dépôt
    Alors la feuille Excel "<feuille>" reproduit le fichier "<fichier>"

    Exemples:
      | feuille      | fichier          |
      | Config       | config.md        |
      | Projects     | projects.md      |
      | Periods      | periods.md       |
      | Workflows    | workflows.md     |
      | Transitions  | transitions.md   |
      | FieldMapping | field_mapping.md |

  @excel
  Plan du scénario: Le VBA lit la catégorie d'un statut dans la feuille Workflows
    Soit le classeur construit avec les paramètres du dépôt
    Quand le VBA donne la catégorie du statut "<statut>" du workflow "<workflow>"
    Alors la catégorie est "<categorie>"

    Exemples:
      | workflow | statut            | categorie    |
      | Team     | En analyse        | Active       |
      | Team     | Prêt              | Wait         |
      | Team     | Terminé           | Done         |
      | Team     | Annulé            | Removed      |
      | Team     | Dépriorisé        | Backlog      |
      | Feature  | En test / recette | Active       |
      | Team     | En test / recette | Unclassified |
      | Team     | Statut inconnu    | Unclassified |

  @excel
  Plan du scénario: Le VBA applique les transitions de la feuille Transitions
    Soit le classeur construit avec les paramètres du dépôt
    Quand le VBA teste la transition de "<depuis>" vers "<vers>" dans le workflow "<workflow>"
    Alors la transition est <verdict>

    Exemples:
      | workflow | depuis           | vers              | verdict   |
      | Team     | Nouveau          | En analyse        | autorisée |
      | Team     | Prêt             | En analyse        | autorisée |
      | Team     | Nouveau          | Prêt              | refusée   |
      | Team     | Nouveau          | Dépriorisé        | autorisée |
      | Team     | Dépriorisé       | Annulé            | autorisée |
      | Team     | Terminé          | Nouveau           | refusée   |
      | Feature  | En test / recette | À déployer       | autorisée |
      | Feature  | En test          | À déployer        | refusée   |
