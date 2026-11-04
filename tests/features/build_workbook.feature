# language: fr
Fonctionnalité: Construire le classeur de test
  Walking skeleton : le code Python assemble les modules VBA du repo dans un .xlsm
  qui compile et dont les tests passent. Il fixe le cadre de la boucle externe (BDD).

  @excel
  Scénario: Le classeur construit compile et les tests hors ligne passent
    Soit les modules VBA du repo
    Quand je construis le classeur de test
    Alors le classeur compile
    Et les tests VBA hors ligne passent sans échec

  @excel @network
  Scénario: La couverture des marqueurs vaut 100 % avec les tests réseau
    Soit les modules VBA du repo
    Quand je construis le classeur de test
    Et je lance les tests VBA avec le réseau
    Alors tous les marqueurs de couverture sont atteints
