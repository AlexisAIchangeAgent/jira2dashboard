# language: fr
Fonctionnalité: Diagnostic d'un blocage des macros
  Quand la protection du poste bloque les macros, scripts\run.ps1 diagnose écrit un rapport à coller dans le ticket :
  empreintes du classeur, origine (marque du web), Office, produits de sécurité, interpréteur du faux Jira et fin du
  journal VBA. Il n'ouvre pas le classeur et ne lit ni python.log ni le journal des événements de Windows.

  Scénario: Un classeur téléchargé d'Internet est signalé avec ses empreintes
    Soit un classeur téléchargé d'Internet (marque du web, zone 3)
    Quand je lance le diagnostic
    Alors le rapport indique la marque du web "yes" et la zone 3
    Et le rapport donne le MD5 et le SHA-256 du classeur

  Scénario: Le rapport ne cite que le journal VBA et aucun secret
    Soit un journal VBA et un python.log qui contient le jeton "tok-SECRET-123"
    Quand je lance le diagnostic
    Alors le rapport cite la fin du journal VBA
    Mais le rapport ne contient pas "tok-SECRET-123"

  Scénario: Le dossier de l'utilisateur est masqué
    Soit un classeur dans le dossier de l'utilisateur
    Quand je lance le diagnostic
    Alors le rapport ne contient pas le dossier de l'utilisateur
    Et le rapport contient "%USERPROFILE%"

  Scénario: Le nom stable du classeur est confirmé
    Soit un classeur dans le dossier de l'utilisateur
    Quand je lance le diagnostic
    Alors le rapport contient "Name: jira2dashboard.xlsm (the stable name)"

  Scénario: Un nom de fichier changé par un téléchargement en double est signalé
    Soit un classeur nommé "jira2dashboard (1).xlsm"
    Quand je lance le diagnostic
    Alors le rapport signale que le nom n'est pas le nom stable "jira2dashboard.xlsm"

  Scénario: Sans classeur, le rapport le dit et donne le reste
    Soit aucun classeur
    Quand je lance le diagnostic
    Alors le rapport dit qu'il n'y a pas de classeur et comment en donner un
    Et le rapport contient "[Office]"
