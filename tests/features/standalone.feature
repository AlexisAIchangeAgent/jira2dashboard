# language: fr
Fonctionnalité: Extraction Python sans VBA
  Quand les macros du classeur ne marchent pas, scripts\run.ps1 extract fait leur travail en Python : il lit les
  paramètres du classeur rempli (comme un fichier, sans Excel), lit Jira avec WinHTTP, calcule les données et les
  résultats avec la référence Python et écrit jira2dashboard.xlsx : des valeurs et les graphiques en SVG, aucun VBA.
  Le journal décrit chaque appel à Jira qui échoue et ce que Jira a dit, jamais le token.

  Contexte:
    Soit un faux Jira avec 150 tickets générés et les anomalies de données connues
    Et un classeur de paramètres rempli pour ce faux Jira

  Scénario: L'extraction Python écrit les données et les résultats du faux Jira
    Quand je lance l'extraction Python
    Alors elle se termine avec le code 0
    Et la feuille Issues a une ligne par ticket du faux Jira
    Et les feuilles Metrics, PIMetrics, DailyLoad, ItemFlow, FillRates et DataChecks reproduisent la référence Python
    Et chaque anomalie connue est trouvée par les contrôles qui lui correspondent et seulement par eux
    Et la feuille Charts a les 8 graphiques du Dashboard en SVG
    Et la feuille Config des résultats masque le token

  Scénario: Une valeur obligatoire vide arrête l'extraction avant tout appel à Jira
    Soit le paramètre "AsOf" vidé dans le classeur de paramètres
    Quand je lance l'extraction Python
    Alors elle se termine avec le code 1
    Et le faux Jira n'a reçu aucun appel
    Et le journal dit "Fill these values of the Config sheet first (Mandatory = True): AsOf"

  Scénario: Une erreur de Jira qui dure est décrite dans le journal et le fichier précédent reste
    Soit un fichier de résultats précédent
    Et le faux Jira répond 500 "Jira is down" à toutes les recherches
    Quand je lance l'extraction Python
    Alors elle se termine avec le code 1
    Et le journal donne pour chaque tentative le statut 500, l'appel search et le message "Jira is down"
    Et le journal dit "HTTP 500 after 5 attempt(s). Jira is down"
    Et le journal ne contient ni le token ni l'en-tête Authorization
    Et le fichier de résultats précédent n'a pas changé

  Scénario: Une limite de débit est attendue, puis l'extraction aboutit
    Soit le faux Jira répond 429 avec Retry-After 7 à la première recherche
    Quand je lance l'extraction Python
    Alors elle se termine avec le code 0
    Et l'extraction a attendu 7 secondes avant de recommencer
    Et la feuille Errors garde l'avertissement du 429

  Scénario: À travers le proxy typé dans ProxyServer
    Soit un proxy devant le faux Jira, typé dans ProxyServer
    Quand je lance l'extraction Python
    Alors elle se termine avec le code 0
    Et la recherche et les historiques sont passés par le proxy

  @excel
  Scénario: La copie du classeur à livrer remplie par Excel suffit, et Excel ouvre le résultat
    Soit la copie d'essai manuel du classeur à livrer, remplie par Excel pour ce faux Jira
    Quand je lance l'extraction Python sur cette copie
    Alors elle se termine avec le code 0
    Et les feuilles Metrics, PIMetrics, DailyLoad, ItemFlow, FillRates et DataChecks reproduisent la référence Python
    Et Excel ouvre le fichier de résultats et ses 8 graphiques sont des images SVG
