# language: fr
Fonctionnalité: Les macros du classeur s'exécutent sur ce poste
  scripts\run.ps1 check-machine ouvre le classeur à livrer avec les macros permises pour le test (la sécurité d'Office est
  relâchée) et exécute une macro en lecture seule. Si elle s'exécute, un blocage à la main vient d'Office (marque du web,
  dossier de confiance, règle) ; si elle ne s'exécute pas, il vient de l'extérieur d'Office : la protection du poste.

  @excel
  Scénario: Le classeur à livrer exécute sa macro
    Soit le classeur à livrer
    Quand je fais exécuter sa macro en lecture seule
    Alors la macro s'exécute et rend l'état du tableau de bord

  @excel
  Scénario: Un classeur sans cette macro est signalé avec le conseil sur la protection du poste
    Soit un classeur sans macro
    Quand je fais exécuter sa macro en lecture seule
    Alors la macro ne s'exécute pas
    Et le conseil parle de la protection du poste
