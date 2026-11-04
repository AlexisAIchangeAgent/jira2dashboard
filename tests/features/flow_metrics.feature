# language: fr
Fonctionnalité: Les six métriques de flux sont exactes
  Calcul de référence en Python, sur neuf tickets écrits à la main dont les résultats sont connus.
  Une métrique qui ne peut pas être calculée affiche « not computable », jamais un faux chiffre.
  Les itérations sont P1 = [05/10, 19/10[ et P2 = [19/10, 02/11[. Date de référence : 02/11/2026.

  Contexte:
    Soit les neuf tickets de flux écrits à la main

  Scénario: Un ticket passé par Dépriorisé puis Nouveau voit son Flow Time repartir à zéro
    Quand je calcule la ligne du ticket "A-2"
    Alors son démarrage est le "10-09" et son premier démarrage le "10-02"
    Et son Flow Time vaut 9 jours
    Et il a 2 retours en backlog

  Scénario: Un ticket terminé sans jamais démarrer est exclu du Flow Time
    Quand je calcule la ligne du ticket "A-3"
    Alors il est un démarrage sauté
    Et son Flow Time est vide

  Scénario: Un ticket rouvert après Terminé est exclu du Flow Time
    Quand je calcule la ligne du ticket "A-4"
    Alors il a été rouvert 1 fois
    Et son Flow Time vaut 12 jours

  Plan du scénario: Flow Velocity, Flow Time, Flow Efficiency et Flow Load en mode A
    Quand je calcule les métriques de l'équipe "<équipe>" pour la période "<période>" en mode "A"
    Alors la métrique "<métrique>" vaut "<valeur>"

    Exemples:
      | équipe | période | métrique                  | valeur         |
      | Alpha  | P1      | Flow Velocity             | 4              |
      | Alpha  | P1      | Flow Time médiane         | 10             |
      | Alpha  | P1      | Flow Time P85             | 10.7           |
      | Alpha  | P1      | Flow Time exclus          | 2              |
      | Alpha  | P1      | Flow Efficiency           | 0.75           |
      | Alpha  | P1      | Flow Load moyenne         | 2.2857142857   |
      | Alpha  | P1      | Flow Load fin de période  | 1              |
      | Alpha  | P2      | Flow Velocity             | 0              |
      | Alpha  | P2      | Flow Time médiane         | not computable |
      | Alpha  | P2      | Flow Efficiency           | not computable |
      | Alpha  | P2      | Flow Load moyenne         | 1              |
      | Beta   | P2      | Flow Velocity             | 2              |
      | Beta   | P2      | Flow Time médiane         | 13             |
      | Beta   | P2      | Flow Time P85             | 15.8           |
      | Beta   | P2      | Flow Efficiency           | not computable |
      | Beta   | P2      | Flow Load fin de période  | 0              |
      | ART    | P2      | Flow Velocity             | 1              |
      | ART    | P2      | Flow Efficiency           | 0.8421052632   |
      | ART    | P1      | Flow Predictability       | not computable |

  Scénario: Le Flow Efficiency n'est pas calculable tant qu'un statut est non classé
    Quand je calcule les métriques de l'équipe "Beta" pour la période "P2" en mode "A"
    Alors la métrique "Flow Efficiency" vaut "not computable"
    Et la métrique "Flow Velocity" vaut "2"

  Plan du scénario: Le Flow Load jour par jour diffère entre le mode A et le mode B
    Quand je calcule le Flow Load de l'équipe "Alpha" en mode "<mode>" pour la période "P1"
    Alors la charge du "<jour>" vaut <charge>

    Exemples:
      | mode | jour  | charge |
      | A    | 10-05 | 3      |
      | B    | 10-05 | 2      |
      | A    | 10-08 | 2      |
      | B    | 10-08 | 3      |
      | A    | 10-11 | 4      |
      | B    | 10-11 | 4      |
      | A    | 10-18 | 1      |

  Scénario: La répartition des tickets terminés par type de travail
    Quand je calcule les métriques de l'équipe "Alpha" pour la période "P1" en mode "A"
    Alors la part "Feature" des tickets terminés vaut 25 %
    Et la part "Defect" des tickets terminés vaut 25 %
    Et la somme des parts des tickets terminés vaut 100 %

  Scénario: Un ticket sans type de travail est compté comme non mappé
    Quand je calcule les métriques de l'équipe "Beta" pour la période "P2" en mode "A"
    Alors la part "Unmapped" des tickets terminés vaut 50 %
    Et la part "Feature" des tickets terminés vaut 50 %

  Scénario: Les Features du Program sont rapportées sous l'équipe ART
    Quand je calcule les métriques de l'équipe "ART" pour la période "P2" en mode "A"
    Alors la métrique "Flow Time médiane" vaut "19"
    Et la part "Feature" des tickets terminés vaut 100 %

  Scénario: Inclus plus exclus égale terminés
    Quand je calcule les métriques de l'équipe "Alpha" pour la période "P1" en mode "A"
    Alors les tickets retenus plus les tickets exclus font la vitesse de flux
