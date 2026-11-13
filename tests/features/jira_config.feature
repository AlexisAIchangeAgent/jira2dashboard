# language: fr
Fonctionnalité: Le bouton « Read the Jira configuration » propose les paramètres à partir de Jira
  Le bouton lit Jira en lecture seule (projets, types de ticket, statuts, workflows, transitions, champs) et écrit des
  feuilles de propositions dans le classeur : JiraIssueTypes, JiraStatuses, JiraWorkflows, JiraTransitions, JiraFields,
  JiraProjectWorkflows (les workflows de chaque projet), JiraProjectFields (les champs de chaque type de ticket) et
  JiraFieldUsage (les champs remplis dans les tickets lus). Il ne change jamais les feuilles de paramètres : l'utilisateur relit les propositions, décide ce qui
  reste humain (ordre, début, fin, attente ou activité, type de passage) et recopie. Rien ne sort du poste.
  Le faux Jira est construit à partir des paramètres du dépôt : les propositions doivent redonner ces paramètres.

  @excel
  Scénario: Les propositions redonnent les statuts, les transitions et les champs du faux Jira
    Soit un faux Jira avec 120 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Quand je clique sur le bouton Lire la configuration Jira
    Alors l'état du tableau de bord dit que la configuration est lue
    Et la feuille JiraIssueTypes donne le workflow de chaque type de ticket de chaque projet de la feuille Projects
    Et la feuille JiraStatuses donne les statuts de chaque type de ticket avec leur catégorie Jira
    Et la feuille JiraWorkflows redonne les statuts de la feuille Workflows, avec une catégorie proposée et les statuts libres
    Et la feuille JiraTransitions redonne les passages de la feuille Transitions
    Et la feuille JiraFields liste les champs du faux Jira avec leur identifiant
    Et les feuilles JiraWorkflows et JiraTransitions disent quels projets utilisent chaque workflow
    Et la feuille JiraProjectWorkflows résume les workflows de chaque projet avec leurs types de ticket
    Et la feuille JiraProjectFields donne les champs de chaque type de ticket de chaque projet
    Et la feuille JiraFieldUsage compte les tickets où chaque champ est rempli
    Et les feuilles de paramètres n'ont pas changé
    Et le journal local décrit la lecture de la configuration, avec ses paramètres et le temps de chaque étape, sans jeton, adresse ni réponse de Jira

  @excel
  Scénario: Sans droit de lire les workflows, les autres propositions sont écrites et l'état le dit
    Soit un faux Jira avec 120 tickets générés avec les projets des paramètres
    Et le faux Jira refuse l'aperçu des workflows avec 401
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Quand je clique sur le bouton Lire la configuration Jira
    Alors l'état du tableau de bord dit que les workflows n'ont pas pu être lus
    Et la feuille JiraStatuses donne les statuts de chaque type de ticket avec leur catégorie Jira
    Et la feuille JiraFields liste les champs du faux Jira avec leur identifiant
    Et les feuilles JiraWorkflows et JiraTransitions n'ont que leurs en-têtes

  @excel
  Scénario: Le bouton « Propose values from Jira » remplit les cellules vides des feuilles à saisir et les marque à confirmer
    Soit un faux Jira avec 120 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les feuilles Workflows et Transitions ont perdu leur dernière ligne et la feuille Config n'a plus l'identifiant du champ Story Points
    Quand je clique sur le bouton Propose values from Jira
    Alors l'état du tableau de bord dit que des valeurs sont proposées
    Et le statut et le passage perdus sont revenus en jaune avec la note « To confirm »
    Et l'identifiant du champ Story Points est revenu en jaune
    Et les itérations des sprints du faux Jira sont proposées en jaune sans qu'aucun jour soit compté deux fois
    Et aucune cellule saisie n'a changé

  @excel
  Scénario: Avec IsUserProjectAdmin à False, le bouton ne lit pas les workflows et propose les statuts
    Soit un faux Jira avec 120 tickets générés avec les projets des paramètres
    Et le classeur à livrer dont la feuille Config pointe sur ce faux Jira avec FlowLoadMode "A"
    Et les feuilles Workflows et Transitions ont perdu leur dernière ligne et la feuille Config n'a plus l'identifiant du champ Story Points
    Et IsUserProjectAdmin vaut "False" dans la feuille Config
    Quand je clique sur le bouton Propose values from Jira
    Alors l'état du tableau de bord dit que les workflows n'ont pas été lus faute d'être administrateur
    Et le faux Jira n'a reçu aucun appel à l'aperçu des workflows
    Et le statut perdu est revenu en jaune mais pas le passage
