
# Gestion et facturation des heures pour les experts fribourgeois

Document excel permetant de gérér ses facturation d'heures pour le canton de Fribourg

Les facture sont automatiquement complétée et exportées sous format pdf.

## Compatibilité
En cas de stockage du fichier sur OneDrive, les fichirs d'exports sont créés sur le bureau.

Ne fonctionne pas avec des versions Excel 2016, Excel 2019 ou antérieures.

N'a pas été testé sur Mac ou Linux

## Installation

Téléchargez le fichier et sauvegardez le dans un dossier local :

https://github.com/GrubCaloz/ExpHeuresFR/blob/main/Fichiers/01_Heures%20expert.xlsm


    
## Permission des macros

La macro principale est disponnible dans le dossier "Code"

Si Windows bloque l'execusion des macros, autorisez les dans la propriété du fichier :

 ![image](https://github.com/GrubCaloz/ExpHeuresFR/assets/163901454/3c3a968d-675d-42ad-940b-1dc4190cd7d5|width=20)



Le tableau comporte 3 onglets :
1.	Tâches : Permet de saisir toutes les tâches effectuées
2.	Annexe : permet de préparer l’annexe à la facture de l’état.
3.	Paramètres : permet de saisir vos coordonnées ou de modifier certaines données du tableau

![image](https://github.com/GrubCaloz/ExpHeuresFR/assets/163901454/e067298b-2f17-41f0-b193-10b9697bc17b)


## Onglet Tâches


Cliquez sur le bouton Nouvelle tâche
La fenêtre ci-dessous s’ouvre

 ![image](https://github.com/GrubCaloz/ExpHeuresFR/assets/163901454/881d4461-a31a-4d3d-b68b-288d8153db03)


Entrez les informations de votre tâche.
Cliquez sur ajouter.
Vous pouvez ajouter plusieurs tâches d’affilées

![image](https://github.com/GrubCaloz/ExpHeuresFR/assets/163901454/12f27fb2-99e0-4809-9ea6-f7205695d119)

 
Vous pouvez modifier, après création, une tâche directement dans le tableau

## Onglet Annexe


 ![image](https://github.com/GrubCaloz/ExpHeuresFR/assets/163901454/92fa769a-20e7-437d-bc12-1d62da1a5796)

Sélectionnez **uniquement la profession et le type d’examen**, les données sont importées du tableau « Tâches » ce qui vous permet de vérifier votre annexe

**Assurez vous de ne pas avoir de facture (PDF) ou d’annexe ouverte avant d’en créer la nouvelle**

Cliquez sur « Générer la facture » (vous devez être connecté à internet à ce moment-là, le formulaire est récupéré en ligne)

Une Pop-Up vous demande si vous êtes le bénéficiaire du paiement

 ![image](https://github.com/GrubCaloz/ExpHeuresFR/assets/163901454/85cccd71-61df-46dd-9225-67ee01a04110)

Les documents sont alors générés en fonction de vos informations, ils sont enregistrés dans le dossier parent de votre fichier :

 ![image](https://github.com/GrubCaloz/ExpHeuresFR/assets/163901454/700a6077-06b2-4850-8b2e-8dc96f600fd7)
![image](https://github.com/GrubCaloz/ExpHeuresFR/assets/163901454/0a9cc8f7-f42e-4df9-8f33-1fc752f240e3)

 

Attention, les documents existants sont écrasés lors de la génération


## Onglet Paramètres

 

Complétez vos informations personnelles et/ou votre employeur.

Ces informations sont utilisées pour générer la facture.

![image](https://github.com/GrubCaloz/ExpHeuresFR/assets/163901454/be8b49a5-55fb-4a62-900b-693b94c83507)

Vous pouvez modifier la liste des professions, le n° de Finance correspond au n° d'ordonnance du Sefri:
https://www.becc.admin.ch/becc/public/bvz/beruf/showAllActive
## Authors

- [@GrubCaloz](https://www.github.com/GrubCaloz)

