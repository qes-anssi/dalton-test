# Dalton

<img src="https://www.sgdsn.gouv.fr/files/styles/ds_image_paragraphe/public/files/Notre_Organisation/logo_anssi.png" alt="Logo ANSSI" width="30%">

## Agence nationale de la sécurité des systèmes d’information

[![badge_catégorie_doctrinal](https://img.shields.io/badge/catégorie-doctrinal-%23e9c7e7)](https://anssi-fr.github.io/#types-de-projets)
[![badge_ouverture_A](https://img.shields.io/badge/code.gouv.fr-contributif-blue)](https://documentation.ouvert.numerique.gouv.fr/les-parcours-de-documentation/ouvrir-un-projet-num%C3%A9rique/#niveau-ouverture)

*Ce projet est édité par l’[ANSSI](https://cyber.gouv.fr/). Pour en savoir plus, voir la [page dédiée à la stratégie open source de l’ANSSI](https://cyber.gouv.fr/open-source-lanssi). Vous pouvez également cliquer sur les badges pour en savoir plus sur leurs significations.*

## À propos
La démarche d’amélioration cyber, thématique et progressive (Dalton) est un projet de la division assistance technique (DAT) visant à recenser et organiser les mesures techniques permettant d’améliorer la sécurité des SI.

La doctrine de l'ANSSI est un terme parapluie regroupant l'ensemble des recommandations et positions de l'ANSSI. Elle est diffusée au travers de guides (les guides techniques, les "Fondamentaux", les "Essentiels"), de billets et notes techniques sur le site du CERT-FR, de présentations réalisées dans diverses instances publiques et privées, ou de publications open source.

La majorité de la doctrine est constituée de recommandation à l'état de l'art permettant de répondre à un attaquant de niveau stratégique (ayant des moyens, de l'expertise et de la détermination). Par opposition, les mesures actuellement cartographiées par Dalton ne constituent pas des recommandations à l'état de l'art de l’ANSSI. Si les mesures recensées permettent d’améliorer la sécurité des SI, cette amélioration peut être marginale au vu des modes opératoires d’attaquants et/ou des efforts nécessaires à leur mise en œuvre. 

Dalton ne se substitue pas aux guides publiés par l’ANSSI, son objectif est de faciliter l’orientation au sein de la doctrine technique (somme des guides et publications techniques).

## Philosophie Dalton
Dalton vise à organiser et vulgariser les mesures et recommandations SSI, à la façon d’une carte. 

Cette carte est destinée à des professionnels de la cyber, notamment pour leur permettre (à leur destination ou celle de leurs clients) de :
* Situer un SI sur la carte : quelles mesures sont déjà mises en œuvre ?
* Identifier des mesures accessibles permettant de sortir rapidement d'une situation de risque inacceptable, et laisser le temps de dérouler des projets plus ambitieux de sécurisation.

Une utilisation projetée, **mais non incluse dans Dalton**, est la création de feuilles de route par étape :
* Identifier la cible de sécurité à atteindre, idéalement une sélection de recommandations adossée à une analyse de risque.
* Identifier un ou plusieurs ensembles de mesures intermédiaires permettant de transiter de l’état courant vers la cible.

Cette feuille de route et son exécution sont comparables à la création et le suivi d’un trajet sur votre GPS favori :
* Les itinéraires (feuille de route) sont personnels : certains préférons l’autoroute (rapide, en peu d’étapes), et d’autre la nationale en privilégiant les arrêts fréquents (progression par de multiples petites transformations). 
* Sur un même itinéraire, prenons l’autoroute, chacun choisira son allure (le nombre de paliers franchis entre deux jalons) et les aires d’autoroute de son choix (les mesures sélectionnées) ;
* Enfin, il n’est pas nécessaire de s’arrêter à toutes les villes et toutes les aires d’autoroutes sur l’itinéraire : c’est-à-dire qu’il n’est pas nécessaire de mettre en place toutes les mesures d’une verticale Dalton.

Par ailleurs, si certaines des mesures Dalton se complètent, d’autres se contredisent : une mesure d’un palier supérieur peut demander de défaire une mesure d’un palier inférieur. Les paliers ne sont pas le reflet d’une organisation projet ou d’une progression SSI idéale. Ils ne font que refléter la position des mesures d’une verticale les unes par rapport aux autres.

### Dalton est actuellement en version de travail
Dalton est un projet émergeant et itératif proposant une nouvelle organisation de la doctrine ANSSI. C’est pourquoi nous avons souhaité mettre à disposition sur ce dépôt les résultats intermédiaires pour permettre de collecter les avis et les retours d’expérience des destinataires de la démarche.

Tous les fichiers et le contenu Dalton sont amenés à fortement évoluer à court terme. Le format sélectionné (CSV) et le mode de publication (Github) ont pour but de simplifier le suivi des modifications.

### Reste à faire déjà identifié pour Dalton
* Positionner la nature des mesures pour chaque entrée (M —, M+).
* Intégrer les mesures des guides ANSSI.
* Rendre les mesures Dalton atomiques.

### Contributions
Les contributions, retours d’expériences et avis sont les bienvenues sur le projet, en particulier :
* Pour les thèmes publiés : identifiez-vous des mesures supplémentaires ?
* L’organisation des mesures au sein des verticales vous semble-t-elle juste ? (« Est-ce que les Daltons sont dans l’ordre ? »)
* Les mesures sont-elles compréhensibles, pouvez-vous les mettre en œuvre ?
* Si vous n’utilisez pas encore Dalton, quels usages envisagez-vous ?
* Si vous l’utilisez déjà, quels usages en faites-vous et avec quel succès ?

## Contenu du dépôt
### Les mesures et les recommandations
#### Fichiers 
Le dossier Matrix du dépôt contient :
* Un fichier CSV comprenant toutes les mesures Dalton. Le fichier CSV est le fichier de référence (en cas de divergence entre les fichiers).
* Un fichier Excel par thématique reprenant les mesures du fichier CSV.

#### Organisation des mesures
* Thématique : champs d'action principal de la mesure (Administration, Accès distants, Gestion des utilisateurs, etc.)
* Verticale : second critère d'organisation. Le choix et l'organisation par verticale a pour objectif de donner de la lisibilité aux mesures au sein d'une verticale. En fonctions des thématiques, les verticales peuvent être motivées par le risque, l'usage, la structure technique des équipements, etc.
* Pallier : numérotation permettant de grouper les mesures et les organiser entre elles au sein d'une verticale. Le critère d'organisation reflète l'avis des expert sur l'ordre de recommandations des mesures.

#### Structure des données
| Colonne   | Description |
| :----:    | :----      |
| ID        | Un identifiant unique | 
|Thématique | Thématiques auxquelles se rattache la recommandation |
|Intitulé   | Titre décrivant de manière synthétique la mesure de sécurité. L’intitulé commence par un verbe d’action à l’infinitif (pas de « faire », « étudier la possibilité de », « s’assurer que », etc.) |
|Palier     | Numéro permettant d’organiser les mesures entre elles | 
|Verticale  | Second critère d’organisation des mesures |
|Description| Explication de la mesure : sa finalité, ce en quoi elle consiste et les actions à mettre en œuvre | 
|Conseils et astuces    | Des conseils et astuces pour la mise en œuvre de la mesure | 
|Justification/Risque | Vulnérabilité et risque couvert (même partiellement) par la mesure. Alternativement, la motivation cachée derrière la mesure, le contexte ou l’historique ayant conduit à la production de la mesure |
|Public visé            | A qui s’adresse la mesure | 
|Exemple/Cas d’usage  | Un exemple d’outil pour mettre en œuvre la mesure |
|Configuration ou Code  | Un exemple de commande, code ou configuration pour mettre en œuvre la mesure | 

### Représentation graphique
Le dossier Visualization contient :
* Un fichier Mermaid par thématique.
* Leur export au format SVG.

### Scripts d’exports
Le dossier Scripts_and_tools contient :
* La macro Excel permettant d’exporter les mesures d’une thématique vers un fichier mermaid. Cette macro a été réalisée à l’aide d’un LLM.
* Le prompt ayant permis de générer la macro Excel.