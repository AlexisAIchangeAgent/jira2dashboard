# language: fr
Fonctionnalité: Faux Jira Cloud
  Un faux serveur Jira sur localhost sert les tests Python et le client VBA. La recherche et le changelog
  viennent des mêmes historiques : ils ne se contredisent jamais.

  Scénario: La recherche est paginée par jeton sans doublon ni oubli
    Soit un faux Jira avec 250 tickets générés
    Quand je lis toutes les pages de la recherche de "PRG,TA,TB" avec 100 par page
    Alors j'ai reçu 250 tickets en 3 pages sans doublon

  Scénario: Le client lit ce qui est revenu, pas ce qu'il a demandé
    Soit un faux Jira qui sert au plus 30 tickets par page avec 100 tickets générés
    Quand je lis toutes les pages de la recherche de "PRG,TA,TB" avec 100 par page
    Alors j'ai reçu 100 tickets en 4 pages sans doublon

  Scénario: Le statut actuel de chaque ticket est la dernière destination de son historique
    Soit un faux Jira avec 250 tickets générés
    Quand je lis toutes les pages de la recherche de "PRG,TA,TB" avec 100 par page
    Et je lis les historiques de tous les tickets reçus
    Alors le statut actuel de chaque ticket est la dernière destination de son historique

  Scénario: Un ticket passé par Dépriorisé puis Nouveau garde tout son historique
    Soit un faux Jira avec les cinq tickets écrits à la main
    Quand je lis les historiques du ticket "TA-2"
    Alors les statuts successifs sont "En analyse, Dépriorisé, Nouveau, En analyse"

  Scénario: Une limite de débit répond 429 avec Retry-After puis le service reprend
    Soit un faux Jira avec les cinq tickets écrits à la main
    Et la recherche répond 429 avec Retry-After "2" deux fois
    Quand je lance trois recherches
    Alors les réponses sont 429, 429, 200
    Et la première réponse porte Retry-After "2"

  Scénario: L'ancienne recherche est retirée
    Soit un faux Jira avec les cinq tickets écrits à la main
    Quand j'appelle "GET /rest/api/3/search?jql=project%3DTA"
    Alors la réponse est 410

  Scénario: Une recherche sans restriction est refusée
    Soit un faux Jira avec les cinq tickets écrits à la main
    Quand je lance la recherche "ORDER BY key ASC"
    Alors la réponse est 400

  Scénario: Le compte se présente comme dans le script PowerShell : Basic, e-mail et token nettoyé
    Soit un faux Jira avec les cinq tickets écrits à la main qui exige le compte "genericemail@domain.com" et le token "tok-123"
    Quand j'appelle "GET /rest/api/3/myself" avec le compte "genericemail@domain.com" et le token "tok-123"
    Alors la réponse est 200
    Quand j'appelle "GET /rest/api/3/myself" avec le compte "genericemail@domain.com" et le token "tok-123 "
    Alors la réponse est 401

  Scénario: Un proxy qui exige l'authentification répond 407 puis laisse passer
    Soit un faux Jira avec les cinq tickets écrits à la main derrière un proxy qui exige "Negotiate"
    Quand j'appelle "GET http://jira.mock.invalid/rest/api/3/myself" par le proxy
    Alors la réponse est 407
    Et l'en-tête de réponse "Proxy-Authenticate" vaut "Negotiate"
    Quand j'appelle "GET http://jira.mock.invalid/rest/api/3/myself" par le proxy avec "Negotiate YIIB"
    Alors la réponse est 200

  Scénario: HTTPS passe par un tunnel CONNECT
    Soit un faux Jira avec les cinq tickets écrits à la main derrière un proxy qui exige "NTLM"
    Quand j'ouvre un tunnel CONNECT vers "jira.mock.invalid:443" sans authentification
    Alors le tunnel est refusé avec 407
    Quand j'ouvre un tunnel CONNECT vers "jira.mock.invalid:443" avec "NTLM TlRMTVNT"
    Alors la requête dans le tunnel répond 200

  Scénario: Le faux Jira donne les projets demandés avec l'id et les types de ticket de chacun
    Soit un faux Jira avec 250 tickets générés
    Quand j'appelle "GET /rest/api/3/project/search?keys=PRG&keys=TA&expand=issueTypes"
    Alors la réponse est 200
    Et la réponse donne les projets "PRG, TA", chacun avec un id et les types de ses tickets

  Plan du scénario: Le faux Jira donne les statuts de chaque type de ticket d'un projet, avec leur catégorie Jira
    Soit un faux Jira avec 250 tickets générés
    Quand j'appelle "GET /rest/api/3/project/<projet>/statuses"
    Alors la réponse est 200
    Et chaque type de ticket de "<projet>" a les statuts du workflow "<workflow>" de la feuille Workflows

    Exemples:
      | projet | workflow |
      | TA     | Team     |
      | PRG    | Feature  |

  Plan du scénario: L'aperçu des workflows d'un projet redonne les statuts et les transitions des paramètres
    Soit un faux Jira avec 250 tickets générés
    Quand je demande l'aperçu des workflows du projet "<projet>" pour tous ses types de ticket
    Alors la réponse est 200
    Et l'aperçu donne le seul workflow "<workflow>", utilisé par tous les types de ticket de "<projet>"
    Et ses statuts sont ceux du workflow "<workflow>" de la feuille Workflows
    Et ses transitions dirigées sont celles du workflow "<workflow>" de la feuille Transitions
    Et ses transitions globales mènent aux statuts libres du workflow "<workflow>"

    Exemples:
      | projet | workflow |
      | TA     | Team     |
      | PRG    | Feature  |

  @excel
  Scénario: Le client VBA lit tous les tickets et tous les historiques du faux Jira, en direct et via le proxy
    Soit un faux Jira avec 250 tickets générés qui exige le compte "genericemail@domain.com" et le token "tok-123" derrière un proxy qui exige "Negotiate"
    Et le classeur de test ouvert sur ce faux Jira
    Quand le classeur lance ses tests contre le faux Jira
    Alors tous les tests VBA contre le faux Jira passent
    Et le faux Jira a reçu des requêtes adressées au site "jira.mock.invalid"
    Et le proxy a reçu l'authentification Windows du client
