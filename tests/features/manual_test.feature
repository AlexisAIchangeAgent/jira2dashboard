# language: fr
Fonctionnalité: Essai manuel avec le faux Jira
  Pour essayer un classeur à la main, scripts\manual-test.ps1 en écrit une copie remplie pour le faux Jira
  (scripts\run.ps1 mock-workbook) : la feuille Config pointe sur le faux Jira, les feuilles Periods et PIs ont leurs
  lignes. Le classeur donné ne change pas : celui à livrer garde ses valeurs client vides.

  @excel
  Scénario: La copie d'essai pointe sur le faux Jira et le classeur à livrer reste vide
    Soit le classeur à livrer
    Quand je prépare la copie d'essai manuel pour le port 8099
    Alors la feuille Config de la copie pointe sur "http://127.0.0.1:8099" avec le compte, la date de référence et les champs du faux Jira
    Et chaque paramètre obligatoire de la copie a une valeur
    Et les feuilles Periods et PIs de la copie ont les itérations et le PI du faux Jira
    Et le classeur à livrer garde BaseUrl, Email, ApiToken et AsOf vides
