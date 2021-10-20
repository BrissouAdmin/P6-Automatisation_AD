# P6-Automatisation_AD
Automatisation de tâches d'administration de Windows Server
Ce projet6 de la formation AIC a pour but d'écrire un script en Python pouvant automatiser certaines tâches répétitives

Ici nous avons choisi d'automatiser la configuration réseau de la machine et certaines tâches de l'Active Directory

Tests réalisés sous Windows Server 2016

# Prérequis
Vous devez placer le fichier base_utilisateurs.csv sur un lecteur réseau partagé avant de lancer le script. Avant d'executer le script, il faut prendre soin d'y modifier, la source du fichier csv

# Les différentes fonctions
Configuration du serveur

	Renommage de la machine
	Configuration du réseau

Gestion des utilisateurs

	Ajout d'utilisateurs via fichier csv
	Supression d'utilisateu via fichier csv
	
Gestion des groupes

	Création de groupes
	Ajout de membres via fichier csv
	Suppression de membres d'un groupe via fichier csv
	Suppression de tous les membres d'un groupe
	Supression de groupes


# Expérience utilisateurs
Un menu semi-interactif a été mis en place pour faciliter l'expérience et donc rendre l'utilisation du script un peu plus conviviale
