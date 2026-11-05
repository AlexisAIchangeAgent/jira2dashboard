# jira2dashboard

Extrait Jira, calcule les six métriques de flux SAFe et les affiche dans un tableau de bord avec graphiques
dans un classeur `.xlsm`. Le code Python (pywin32) construit le classeur et son VBA ; les mappings (statuts,
transitions, champs, FlowLoadMode...) sont des **paramètres écrits dans le classeur**, jamais du code.

Plateforme : Windows 11, PowerShell 5.1, Python 3.14. Exécution : **pywin32 + stdlib seulement** (pytest,
pytest-bdd, pytest-cov = extras de dev). Le repo est autonome : les processus Excel enfants importent
`jira2dashboard` via `PYTHONPATH=<repo>\src`, sans `pip install`.

## Arborescence

```
src/jira2dashboard/
  excel/com.py          infrastructure COM (Excel, VBA, classeur) ; fonctions pures testées en unitaire
  testbench/            build + compilation + exécution des tests VBA ; children/ = scripts lancés par chemin
  cli/check_machine.py  les 21 contrôles de la machine (Excel, confiance VBA...)
  params/               paramètres de l'appli : mdtable.py (lit les tables md), model.py (charge + valide),
                        write.py (écrit les feuilles Workflows, Transitions, FieldMapping dans le classeur)
  flow/                 les six métriques de flux (E4). Référence Python sans Excel : model.py (vocabulaire),
                        intervals.py (StatusIntervals), itemflow.py (ItemFlow), load.py (Flow Load A/B),
                        metrics.py (par période et par équipe ; « not computable » jamais 0).
                        Côté Excel : formulas.py (texte des formules, testé sans Excel) et sheets.py (tables
                        ItemFlow, DailyLoad, Metrics en formules visibles), blueprint.py (le plan écrit dans le
                        classeur : feuilles FlowFormulas et ChartSpec), dashboard_spec.py (graphiques et formules
                        de ChartData), sheets.py (COM : écrit le plan, lit les tables, exporte les PNG).
                        Les tables, formules et graphiques sont construits par le VBA (modFlowBuild.bas) depuis
                        ce plan ; modExtract.bas écrit les feuilles brutes Issues et StatusIntervals.
                        checks.py : les 29 contrôles de données DC01-DC29 (référence Python)
  build_workbook.py     construit le classeur à livrer out\release\jira2dashboard.xlsm (sans module de test)
  clean.py              vide out\inter avant un build (sauf les caches de test) ; out\release n'est jamais touché
  mock_workbook.py      copie d'essai manuel remplie pour le faux Jira : out\inter\jira2dashboard_mock.xlsm
  mockjira/             faux Jira Cloud sur localhost (E3) : dataset.py (historiques), generator.py (volume seedé),
                        jql.py, api.py (réponses pures), server.py (http.server + fautes injectables)
  paths.py              emplacements (repo, vba, params, out\release, out\inter, children)
params/                 workflows.md, transitions.md, field_mapping.md : les paramètres (statuts, transitions,
                        champs). On les modifie ici ou dans le classeur, jamais dans le code
vba/                    13 modules .bas (ASCII uniquement), artefact de premier rang ; modParams.bas lit les feuilles,
                        modFlowBuild.bas construit tables et graphiques, modActions.bas = le bouton du Dashboard
tests/unit/             Python pur, sans Excel (rapide)
tests/features/         scénarios Gherkin (*.feature, en français)
tests/steps/            steps pytest-bdd
tests/integration/      tests marqués `excel` (vide pour l'instant ; le walking skeleton est dans features/steps)
docs/                   les 7 md de contexte et de paramètres
scripts/                setup.ps1, activate.ps1, run.ps1, release.ps1, manual-test.ps1, start-mock.ps1, stop-mock.ps1
out/release/            le classeur à livrer (build) et les paquets de release (out\release\<version>)
out/inter/              tout le reste : classeurs de test, copie d'essai, exports, caches (vidé à chaque build)
archive/                ancien code et sondes (rien n'a été supprimé)
```

## Commandes

Chaque script et chaque commande expliquent leur fonctionnement et leurs arguments avec `--help` :
`.\scripts\run.ps1 --help` liste les commandes, `.\scripts\run.ps1 build --help` donne les options de `build`,
`.\scripts\manual-test.ps1 --help` (ou `-?`) l'aide d'un script.

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Dev    # crée .venv (py -3.14), installe ; -Proxy <url> si besoin
. .\scripts\activate.ps1                                           # optionnel : active le venv
.\scripts\run.ps1 test            # unitaires, sans Excel (boucle rapide)
.\scripts\run.ps1 coverage        # idem + couverture du code pur, échec sous 90 % (voir pyproject.toml)
.\scripts\run.ps1 test-excel      # scénarios qui pilotent Excel (hors réseau)
.\scripts\run.ps1 test-network    # scénarios qui pilotent Excel + réseau (couverture 100 %)
.\scripts\run.ps1 vba-tests       # build + compile + tests VBA ; --offline, --bisect, --out nom.xlsm
.\scripts\run.ps1 check-machine   # 21 contrôles de la machine
.\scripts\run.ps1 mock-jira       # faux Jira sur http://127.0.0.1:8099 (--anomalies : avec les anomalies connues), Ctrl+C pour arrêter
.\scripts\run.ps1 build           # construit le classeur à livrer : out\release\jira2dashboard.xlsm (--out nom.xlsm : dans out\inter)
.\scripts\run.ps1 clean           # vide out\inter sauf les caches de test (build et vba-tests le font d'abord)
.\scripts\manual-test.ps1         # essai manuel : faux Jira + copie remplie du classeur, ouverte dans un nouvel Excel
.\scripts\stop-mock.ps1           # arrête le faux Jira lancé par manual-test.ps1 ou start-mock.ps1
.\scripts\release.ps1 -Version v1.0.0            # essai à blanc : zip du code + classeur dans out\release\v1.0.0
.\scripts\release.ps1 -Version v1.0.0 -Publish   # idem, puis crée la release et son tag sur GitHub
```

Release : le script part d'un arbre propre (tout est commité). Le zip est le code du dernier commit, sans
`CLAUDE.md`, `claude-code-handoff.md`, `archive\`, `.claude\` ni `.venv\` (règles `export-ignore` de
`.gitattributes`, vérifiées par `tests\unit\test_release_contents.py`). Avec `-Publish`, le commit doit déjà être
poussé. Les liens « Source code » ajoutés par GitHub suivent les mêmes règles. Le classeur de la release est
`out\release\jira2dashboard.xlsm`, construit par le script à partir du même commit.

## Essai manuel avec le faux Jira

Pour essayer un classeur à la main, sans lancer les tests automatiques :

```powershell
.\scripts\manual-test.ps1                       # le classeur out\release\jira2dashboard.xlsm (construit s'il manque)
.\scripts\manual-test.ps1 -Workbook C:\x.xlsm   # une autre version du classeur ; -Port 8100 si 8099 est pris
.\scripts\stop-mock.ps1                         # à la fin : arrête le faux Jira
```

Le script :

1. écrit la copie `out\inter\jira2dashboard_mock.xlsm`, dont les feuilles Config, Periods et PIs sont remplies
   pour le faux Jira (adresse, compte, date de référence, champs) ;
2. lance le faux Jira (1200 tickets et les anomalies connues) dans une fenêtre réduite ;
3. ouvre la copie dans un nouvel Excel.

Le classeur donné ne change pas. Dans la copie, activer le contenu, puis cliquer **Read the Jira configuration**,
puis **Extract from Jira and compute**.

Attention : Excel peut planter juste après la dernière page de recherche. C'est le défaut connu de
l'extraction. La copie s'ouvre donc dans un Excel séparé : enregistrer son autre travail avant.

`stop-mock.ps1` ne tue que le faux Jira que ces scripts ont lancé : le numéro et l'heure de démarrage du
processus sont gardés dans `%TEMP%\jira2dashboard\`.

## Livraison : le classeur `out\release\jira2dashboard.xlsm` (E6)

Deux façons de livrer, au choix d'AL : **le fichier `.xlsm` tel quel**, ou **le code du projet**, à builder sur le
poste client (Python 3.14 : `scripts\setup.ps1`, puis `scripts\run.ps1 build`). Le classeur n'a pas besoin de Python
pour fonctionner : le VBA lit Jira et construit les tables, les formules et les graphiques.

Sur le poste client (le classeur le rappelle : les premières étapes dans l'ordre des boutons sont sur le
**Dashboard**, à partir de la cellule T1, et le manuel complet est dans l'onglet **Manual**, juste après) :

1. **Débloquer le fichier** s'il vient d'un téléchargement ou d'un mail (marque « Mark of the Web », sinon Excel
   bloque les macros) : clic droit sur le fichier > Propriétés > cocher « Débloquer », ou
   `Unblock-File .\jira2dashboard.xlsm` dans PowerShell. Ouvrir le classeur et cliquer sur « Activer le contenu ».
2. Feuille **Config** : `BaseUrl`, `Email`, `ApiToken`, `WorkTypeField` (id du champ type de travail), `AsOf` (date
   de l'extraction, UTC). `ProxyServer` vide : le proxy des réglages Windows est trouvé comme le fait le navigateur
   (détection automatique WPAD, script PAC, puis proxy fixe) ; `DIRECT` : pas de proxy ; sinon `hôte:port`.
   Feuille **Projects** : les projets Jira et lequel est le Program. Feuille **Periods** : les itérations (`Period | Start | End`, fin exclue).
   La colonne **Example** de `Config`, à côté de **Value**, montre la forme attendue (elle n'est jamais lue).
   La colonne **Mandatory** de `Config` dit si une valeur vide est permise. `True` : le bouton **Extract from Jira and
   compute** s'arrête avant d'appeler Jira et nomme les valeurs vides (le bouton **Read the Jira configuration** ne
   demande que `BaseUrl`, `Email` et `ApiToken`). `False` : une valeur vide est permise (valeur par défaut, pas de
   filtre, pas de proxy, ou contrôles « not computable »).
3. Feuille **Dashboard** : bouton **Extract from Jira and compute**. L'extraction se fait par petites tranches : la barre
   de progression de la feuille (à côté du bouton) avance, la cellule `A3` dit ce qui se passe, Excel reste utilisable entre
   deux tranches et le bouton **Cancel** arrête l'extraction avant l'écriture des feuilles brutes (les résultats
   précédents restent intacts). La première partie (lecture de Jira, sans taille connue) remplit 30 % de la barre : c'est une
   estimation. À la fin, les feuilles `ItemFlow`, `DailyLoad`,
   `Metrics`, `ChartData` et les graphiques sont **reconstruits à la taille des nouvelles données** : on peut
   relancer autant de fois que nécessaire. Pour l'équipe du graphique Flow Distribution : cellule `B2`.
4. **Le token API reste en clair dans la feuille Config** : l'effacer avant de transmettre le classeur à quelqu'un d'autre. La BaseUrl doit commencer par https://.
5. Les feuilles `FlowFormulas`, `ChartSpec` et `ChecksSpec` décrivent les formules, les graphiques et les contrôles : ne pas les modifier.
6. **Contrôles de données (feuille `DataChecks`)** : les 29 contrôles DC01-DC29, calculés par des formules sur les champs
   du ticket, l'historique des statuts, les liens « Blocks » (table `IssueLinks`) et les taux de remplissage (table
   `FillRates`). Chaque ligne donne le nombre de tickets en défaut (`Items`) et `OK`, `KO` ou `not computable`. Pour voir
   les tickets, filtrer la table `ItemFlow` sur la colonne `DC..` égale à 1. **Un contrôle qui lit un champ Jira non
   configuré est `not computable` et ne signale personne** (jamais « OK » sans avoir regardé). À renseigner dans `Config` :
   `StoryPointsField`, `BusinessValueField`, `PIField`, `TeamField`, `SprintField`, `TargetStartField`, `TargetEndField`
   (identifiants `customfield_NNNNN` de votre Jira) ; les types dans `FeatureTypes`, `SubTaskTypes`, `PointsRequiredTypes`,
   `ParentRequiredTypes` ; le lien de blocage dans `BlocksLinkType` ; le seuil de DC25 dans `FillRateThreshold` (90) ; et les
   fenêtres des PI dans la feuille `PIs` (PI, Start, End, noms de la forme `PI#01-2026`). Les interprétations retenues sont des
   propositions à valider avec les équipes (voir `claude-code-handoff.md`).
7. **Journal local** : chaque exécution écrit des lignes dans
   `%LOCALAPPDATA%\jira2dashboard\logs\jira2dashboard.log` (module `modLog`). Une ligne `RUN` au début (version
   d'Excel, nom du classeur) et à la fin (`end ok`, `end cancelled`), un `BEGIN` et un `END` autour de chaque étape
   (page de recherche, lot de changelogs, feuilles brutes, feuilles de flux), une ligne `ERROR` par erreur. Si Excel
   plante, le dernier `BEGIN` sans `END` donne l'étape en cours. Le journal ne contient ni le token, ni l'e-mail, ni
   l'URL, ni les réponses de Jira (les erreurs 602 et 607, qui citent la réponse, n'écrivent que leur numéro). Il reste
   sur le poste (pas dans OneDrive). Au-delà de 1 Mo, l'ancien fichier devient `jira2dashboard.1.log`. Relire le
   fichier avant de le transmettre. Paramètre `LogLevel` de `Config` : `Steps` (par défaut, toutes les étapes) ou
   `Errors` (seulement le début, la fin et les erreurs). Si une exécution ne s'est pas terminée (Excel fermé ou planté),
   la suivante écrit `RUN previous run did not end (started ...)` : un indicateur est gardé dans le registre de
   l'utilisateur (`HKCU\Software\VB and VBA Program Settings\jira2dashboard`).
8. **Lire la configuration Jira** (bouton **Read the Jira configuration** du `Dashboard`, module `modJiraConfig`) :
   lit Jira en lecture seule pour les projets de la feuille `Projects` et écrit des **feuilles de propositions**,
   jamais les feuilles de paramètres :
   - `JiraIssueTypes` : le workflow de chaque type de ticket de chaque projet (un projet peut avoir un workflow par
     type) ;
   - `JiraStatuses` : les statuts de chaque type, avec la catégorie Jira (il suffit du droit de consulter le projet) ;
   - `JiraWorkflows` : mêmes premières colonnes que `Workflows`. Catégorie proposée : à faire → `Backlog`, terminé →
     `Done` ; « en cours » reste vide (Active ou Wait à décider). `IsFree = 1` si Jira autorise ce statut depuis tous
     les autres. `Order = 1` pour le statut de création ;
   - `JiraTransitions` : mêmes premières colonnes que `Transitions`, une ligne par passage autorisé ;
   - `JiraFields` : les identifiants des champs, pour les lignes `...Field` de `Config`.

   Restent à décider à la main : `Order`, `IsStart`, `IsEnd`, Active ou Wait, `Removed` au lieu de `Done`, le `Type`
   d'un passage. Relire, puis recopier ce qui convient dans les feuilles de paramètres. Si le compte n'a pas le droit
   de lire les workflows (« View (read-only) workflow »), les autres feuilles sont quand même remplies et la cellule
   d'état le dit. Appels : `claude-code-handoff.md` section 8c. Rien ne sort du poste.

**Formules Excel (E4)** : le test `tests\features\flow_excel.feature` compare cellule par cellule les formules du
classeur à la référence Python (faux Jira, extraction VBA, formules, en mode A puis B). `Config.WorkTypeField` est
l'id du champ personnalisé du type de travail ; `Config.AsOf` est obligatoire pour l'extraction.

**Tableau de bord (E5)** : la feuille `Dashboard` montre huit graphiques natifs Excel (Velocity, Flow Time médian et
P85, Flow Load moyen, Flow Efficiency, Flow Distribution, Flow Predictability avec la bande cible 80-100 %, et la courbe
quotidienne de Flow Load), un graphique par métrique avec une série par équipe. La bande 80-100 % est une proposition (repère SAFe). Les points viennent de la feuille `ChartData` (une formule par point). Une valeur « not computable »
est un trou dans la courbe, jamais un 0, avec une note sous le graphique. Flow Distribution montre une série par
type de travail pour l'équipe choisie dans la cellule `B2` (`All teams` = toutes les équipes ensemble). Les
graphiques sont exportés en PNG dans `out\dashboard\` par le test `flow_excel.feature`.

**Métriques de flux (E4)** : `FlowLoadMode` (`A` par défaut) et `AsOf` (date de l'extraction en UTC, vide = le
moment du calcul) sont dans `Config` ; les itérations sont dans la feuille `Periods` (`Period | Start | End`,
vide dans le dépôt : calendrier du client, intervalles semi-ouverts [Start, End[). Les Features du Program sont
rapportées sous l'équipe « ART ». Le Flow Time et le Flow Efficiency laissent de côté les tickets terminés sans
démarrage valide ou rouverts (comptés et listés à part). Flow Load : photo à la fin de chaque jour UTC.

**Projets : feuille `Projects`** (écrite depuis `params\projects.md`) : une ligne par projet Jira à extraire avec
son rôle (`Program` pour le train qui porte les Features, un seul ; `Team` pour les autres) et le nom de l'équipe.
Par défaut les clés anonymisées du script d'extraction, `TTT` étant le Program : à remplacer dans le classeur.
Le paramètre `Filter` de `Config` ajoute une restriction JQL (par exemple `updated >= -90d`, sans `ORDER BY`).

**Connexion et proxy : dans la feuille `Config` du classeur** (écrite depuis `params\config.md`, jamais dans le
code) : `BaseUrl`, `Email`, `ApiToken` (à saisir dans le classeur, refusé dans le dépôt), `Filter`, `ProxyServer`
(`host:port` ou `http://host:port`), `ProxyBypass`, `ProxyWindowsAuth` (Oui = la session Windows répond au proxy,
Non = jamais envoyée), `TimeoutSeconds` (60), `MaxRetries` (4). Politique de reprise comme le script PowerShell :
429 et tous les 5xx, avec `Retry-After` s'il est présent.

Faux Jira pour le classeur : lancer `mock-jira`, puis mettre dans `Config` l'URL, l'e-mail et le token qu'il
affiche (Basic auth comme le script PowerShell : `base64("email:token")`, token nettoyé). Options calquées sur ce
script : `--projects TTT,AAA,...` (le premier est le Program ; par défaut ceux de `params\projects.md`), `--bv-field`, `--sp-field`, `--pi-field`,
`--email`, `--token`, `--no-auth`, `--issues`, `--seed`, `--port`. Il sert aussi de proxy HTTP (mettre
`host:port` dans `ProxyServer`) : requêtes en forme absolue et tunnel `CONNECT`, avec `--proxy-challenge Negotiate`
pour un proxy qui exige l'authentification (407 tant que la requête n'a pas `Proxy-Authorization` ; n'importe
quelle valeur est acceptée). TLS 1.2 dans le tunnel et l'échange NTLM/Kerberos réel ne sont pas simulés : à
vérifier chez le client.
Les tests VBA `M_*` de `modTests.bas` ne tournent que depuis les scénarios Python (`test-excel`) : ils ont
besoin du serveur que pytest démarre.

**Défaut ouvert** : Excel peut planter (`0xC0000005`) pendant l'extraction, bien plus souvent après plusieurs
extractions dans le même Excel. Voir `claude-code-handoff.md`, section 8b. Les scénarios Excel ouvrent donc un Excel
neuf par scénario ; `J2D_ONE_EXCEL=1` remet un seul Excel pour tout le module (mode « stress », pour tester un
correctif : il faut alors 5 passes vertes de suite).

Par défaut pytest exclut les marqueurs `excel` et `network` (`addopts`). On ne tue jamais un `EXCEL.EXE` qui
n'est pas celui du script.

Les commandes `check-machine` et `run-vba-tests` existent aussi comme entry points (`pip install -e .`, testé :
21/21 et COMPILED_OK). **Installation éditable uniquement** : `paths.py` retrouve `vba\` et `out\` à partir de
l'emplacement du code. Ce n'est pas nécessaire, `run.ps1` suffit.

La cible de 90 % s'applique au code pur : le code qui a besoin d'Excel, du registre ou d'un processus enfant est
exclu de la mesure (liste dans `pyproject.toml`, à compléter pour toute nouvelle fonction liée à Excel) et il est
exercé par les tests `excel`, `check-machine` et `vba-tests`.

## Boucles de travail : TDD imbriquée dans BDD

1. **Boucle externe (BDD)** : écrire un scénario Gherkin rouge qui décrit un comportement visible
   (`tests/features/*.feature`, steps dans `tests/steps/`).
2. **Boucle interne (TDD)** : test unitaire rouge (`tests/unit/`) -> code minimal -> vert -> refactor, jusqu'à
   ce que le scénario passe.
3. Pyramide : beaucoup d'unitaires purs, peu d'intégrations Excel, quelques scénarios de bout en bout.
4. Cibles : couverture `pytest-cov` >= 90 % sur le code pur ; marqueurs VBA 100 % (CoverageMap) ; aucun test
   Excel dans la boucle rapide.
5. KISS / YAGNI : un paquet n'est créé que quand il contient du code. Les paramètres vont dans le classeur
   (garde-fou : `tests/unit/test_parameters_not_in_code.py`).

Toute tâche longue en VBA suit `modBackground` / `modTests` : tranches planifiées par `Application.OnTime`,
barre de progression, résultats ligne par ligne, bouton Annuler.

## Documents

- `docs\` : fichiers de contexte et de paramètres de l'application.
- `rovo\` : les agents Rovo, à recopier dans Rovo Studio (ces fichiers sont la version de référence).
  - `rovo_settings_analyser.md` : agent « Settings Analyser » (types de ticket, totaux par type, statuts par type,
    champs personnalisés d'un projet) pour comparer Jira avec le classeur ;
  - `rovo_flow_metric.md` : agent « Flow Metrics by Type » (métriques de flux par type de ticket, chaque nombre
    marqué exact, approché ou échantillon, avec sa JQL) ; les chiffres officiels restent ceux du classeur ;
  - `rovo_train_steering.md` : agent « ART Flow Steering » pour le pilotage du train. Il contrôle les données avant
    l'extraction, explique les chiffres publiés et rédige une note de pilotage. Il ne calcule aucune métrique : le
    classeur reste le moteur. Ses compétences sont dans `skills\` (un `SKILL.md` par étape) ;
  - `confluence_flow_metrics_page.md` : modèle de la page Confluence « Flow metrics <PI> », où le RTE colle les
    tableaux `Metrics` et `DataChecks` et les graphiques du classeur ;
  - `jira_dashboard_data_quality.md` : tableau de bord Jira de filtres enregistrés `DQ DCxx`, monté à la main. Les
    comptes sont calculés par Jira ;
  - `rovo-guide.md` : copie à jour (2026-11-05) de `docs\rovo-guide.md`, qui reste figé. Son agent « Flow Data
    Steward » est remplacé par « ART Flow Steering ».
- `docs\cookbook_python.md` et `docs\vba-modules-tested.md` sont des **copies de code non maintenues**,
  périmées depuis la restructuration (le code vit dans `src\` et `vba\`).
- `claude-code-handoff.md` est mis à jour (nouveau dossier, nouvelles commandes, règle `bootstrap_test 15/15`
  retirée). Les sections 1-3 et 5-9 restent le contexte d'origine ; la source de commandes est ce README.
