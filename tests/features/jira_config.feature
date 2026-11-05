# language: fr
Fonctionnalité: Le bouton « Read the Jira configuration » propose les paramètres à partir de Jira
  Le bouton lit Jira en lecture seule (projets, types de ticket, statuts, workflows, transitions, champs) et écrit des
  feuilles de propositions dans le classeur : JiraIssueTypes, JiraStatuses, JiraWorkflows, JiraTransitions et
  JiraFields. Il ne change jamais les feuilles de paramètres : l'utilisateur relit les propositions, décide ce qui
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
    Et les feuilles de paramètres n'ont pas changé
    Et le journal local décrit la lecture de la configuration, sans jeton, adresse ni réponse de Jira

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
