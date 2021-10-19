
from pyad import *

import csv

############################################################################
# Connexion au serveur Active Directory avec un utilisateur admin du domaine
############################################################################
pyad.set_defaults(ladp_server="WINDOWS-AD.akb.lab",username="python.demo",password="Azerty2016$")


############################################################################
# Fonction de renommage de la machine
############################################################################

import wmi

def renommer_pc():
        # Connexion sur une machine locale
        c = wmi.WMI()

        # Connexion sur une machine distante
        #c = wmi.WMI("WINDOWS-AD.akb.lab", user=r"python.demo", password="Azerty2016$")

        # Nouveau nom de la machine
        nouveau_nom_pc=input('Veuillez entrer le nouveau nom de la machine ')
        #nouveau_nom_pc=("WIN-IT")

        for system in c.Win32_ComputerSystem():
                system.Rename(nouveau_nom_pc)
        print()
        print("La machine a été renommée en " + nouveau_nom_pc)        
        print()



############################################################################
# Fonction de configuration du réseau
############################################################################

import wmi

def configurer_reseau():
    # Connexion sur une machine locale
    c = wmi.WMI()

    # Connexion sur une machine distante
    #c = wmi.WMI("WINDOWS-AD.akb.lab", user=r"python.demo", password="Azerty2016$")

    # obtention des configurations des adaptateurs réseau
    nic_configs = wmi.WMI().Win32_NetworkAdapterConfiguration(IPEnabled=True)

    # premier adaptateur réseau
    nic = nic_configs[0]

    # Adresse IP, Masque de sous-réseau et Passerelle devraient être des objets Unicode (u avant les adresses)
    #ip = u'192.168.0.12'
    #subnetmask = u'255.255.255.0'
    #gateway = u'192.168.0.1'

    ip = input(u'Saisir adresse ip: ')
    subnetmask = input(u'Saisir le masque de sous réseau: ')
    gateway = input(u"Saisir l'adresse de la passerelle: ")

    #print(nic)

    # Configuration de l'adresse IP , le masque de sous-réseau et la passerelle
    # Les méthodes EnableStatic() et SetGateways() demandent une "liste" de valeurs
    nic.EnableStatic(IPAddress=[ip],SubnetMask=[subnetmask])
    nic.SetGateways(DefaultIPGateway=[gateway])

    print()
    print("La carte réseau a été reconfigurée avec les données suivantes")
    print("Ip: " + ip + " /Masque de sous-réseau: " + subnetmask + " /Passerelle: " + gateway)
    print()


############################################################################
# Fonction d'ajout d'utilisateurs à partir d'un fichier csv
############################################################################

def ajouter_utilisateurs():

    # Déclaration de l'OU dans laquelle on veut travailler
    ou=pyad.adcontainer.ADContainer.from_dn("OU=Logistique,OU=Employés,DC=akb,DC=lab") 

    # Déclaration du fichier csv contenant la liste des utilisateurs
    with open(r"P:\base_utilisateurs.csv") as fichier_utilisateurs:
    
    # Lecture du fichier csv: les données sont délimitées par ","
        liste_utilisateurs=csv.reader(fichier_utilisateurs,delimiter=",")
    
        MDP_utilisateur="Logistique2021$"
        
        print()
        print("Liste des utilisateurs ajoutés")
        print() 
    # Utilisation d'une boucle pour traiter toutes les lignes du fichier
        i=0
        for utilistateur in liste_utilisateurs:
           
            if(i>0):
                employeeID=utilistateur[0]
                givenName=utilistateur[1]
                sn=utilistateur[2]
                department=utilistateur[3]
                mail=utilistateur[4]
                sAMAccountName=utilistateur[5]
            
                print(sAMAccountName)
            
                creer_utilisateur=pyad.aduser.ADUser.create(sAMAccountName, ou, password=MDP_utilisateur, upn_suffix=None, enable=True, optional_attributes={"employeeID":employeeID, "givenName":givenName, "sn":sn, "mail":mail, "department":department})               
            i=i+1


#########################################################################
# Fonction Suppression de plusieurs Utilisateurs via un fichier CSV
#########################################################################

def supprimer_utilisateurs():

# Déclaration du fichier csv contenant la liste des utilisateurs à suprimer
    with open(r"P:\base_utilisateurs.csv") as fichier_utilisateurs:
    
        # Lecture du fichier csv: les données sont délimitées par ","
        liste_utilisateurs=csv.reader(fichier_utilisateurs,delimiter=",")

        print()
        print("Liste des utilisateurs supprimés")
        print()

        # Utilisation d'une boucle pour traiter toutes les lignes du fichier
        i=0
        for utilistateur in liste_utilisateurs:
                
            if(i>0):
                sAMAccountName=utilistateur[5]
                pyad.aduser.ADUser.from_cn(sAMAccountName).delete()
                                 
                print("ID-utilisateur supprimé est" + utilistateur[0])             
            i=i+1



#########################################################################
# Fonction Création de groupes
#########################################################################

def creer_groupes():
    # nom du groupe à créer
    nom_groupe="test_group_python1"
    # l'OU du groupe à créer
    group_ou=pyad.adcontainer.ADContainer.from_dn("OU=Logistique,OU=Employés,DC=akb,DC=lab")
    # création du group
    creer_group=pyad.adgroup.ADGroup.create(nom_groupe, group_ou, security_enabled=True, scope='GLOBAL', optional_attributes={"description":"Créé via script Python"})

    print("Le groupe créé est: " + nom_groupe)


#########################################################################
# Fonction Ajout de membres à un groupe
#########################################################################

def ajouter_membres():

    # nom du groupe dans lequel ajouter les membres
    nom_groupe="test_group_python1"

    group=pyad.adgroup.ADGroup.from_cn(nom_groupe)

    # Déclaration du fichier csv contenant la liste des utilisateurs
    with open("base_utilisateurs.csv") as fichier_utilisateurs:
        
        # Lecture du fichier csv: les données sont délimitées par ","
            liste_utilisateurs=csv.reader(fichier_utilisateurs,delimiter=",")
                    
            print()
            print("Liste de membres ajoutés ")
            print() 
        # Utilisation d'une boucle pour traiter toutes les lignes du fichier
            i=0
            for utilistateur in liste_utilisateurs:
            
                if(i>0):
                    
                    sAMAccountName=utilistateur[5]
                                        
                    utilisateur=pyad.aduser.ADUser.from_cn(sAMAccountName)
                
                    group.add_members([utilisateur])
                    print(sAMAccountName)
                i=i+1



#########################################################################
# Fonction Suppression de membres d'un groupe via fichier csv
#########################################################################

def supprimer_membres():

    # nom du groupe dans lequel ajouter les membres
    nom_groupe="test_group_python1"

    group=pyad.adgroup.ADGroup.from_cn(nom_groupe)

    # Déclaration du fichier csv contenant la liste des utilisateurs
    with open("base_utilisateurs.csv") as fichier_utilisateurs:
        
        # Lecture du fichier csv: les données sont délimitées par ","
            liste_utilisateurs=csv.reader(fichier_utilisateurs,delimiter=",")
                    
            print()
            print("Liste de membres supprimés ")
            print() 
        # Utilisation d'une boucle pour traiter toutes les lignes du fichier
            i=0
            for utilistateur in liste_utilisateurs:
            
                if(i>0):
                    
                    sAMAccountName=utilistateur[5]
                                        
                    utilisateur=pyad.aduser.ADUser.from_cn(sAMAccountName)
                
                    group.remove_members([utilisateur])
                    
                    print(sAMAccountName)
                i=i+1




#########################################################################
# Fonction Suppression de tous les membres d'un groupe
#########################################################################

def supprimer_all_membres():
    # nom du groupe dont tous les membres seront supprimés
    nom_groupe="test_group_python1"

    # suppression de tous les membres du groupe
    group=pyad.adgroup.ADGroup.from_cn(nom_groupe)

    # affichage de message suite à la suppression de tous les membres
    print()
    print("Le groupe " + nom_groupe + " ne contient plus de membres")
    print()   
            
    group.remove_all_members()
                

#########################################################################
# Fonction Supression de groupes
#########################################################################

def supprimer_groupes():
    # nom du groupe à supprimer
    nom_groupe="test_group_python1"

    group=pyad.adgroup.ADGroup.from_cn(nom_groupe).delete()

    print("Le groupe supprimé est " + nom_groupe)





#########################################################################
# Menu pour le choix des tâches
#########################################################################

def menu():
    
    print()
    print("=============== CONFIGURATION DU SERVEUR ==================")
    print("[1] Renommage de la machine")
    print("[2] Configuration du réseau")
    print()
    print("=============== GESTION DES UTILISATEURS ==================")
    print("[3] Ajout d'utilisateurs à partir d'un fichier csv")
    print("[4] Supression d'utilisateurs à partir d'un fichier csv")
    print()
    print("================== GESTION DES GROUPES =====================")
    print("[5] Création de groupes")
    print("[6] Ajout de membres à un groupe")
    print("[7] Suppression de membres d'un groupe via fichier csv")
    print("[8] Suppression de tous les membres d'un groupe")
    print("[9] Supression de groupes")
    print()
    print("================== VERS LA SORTIE =================")
    print("[99] Sortie du Script")

print()
menu()
print()
choix = int(input("Entrer votre choix: "))
print()

while choix != 0:
    if choix == 1:
        renommer_pc()  
    elif choix == 2:
        configurer_reseau()
    elif choix == 3:
        print("Vous avez choisi la création d'utilisateurs ........")
        ajouter_utilisateurs()
    elif choix == 4:
        print("Vous avez choisi la suppression d'utilisateurs ........")
        supprimer_utilisateurs()
    elif choix == 5:
        print("Vous avez choisi la création de groupes ........")
        creer_groupes()
    elif choix == 6:
        print("Vous avez choisi Ajout de membres à un groupe........")
        ajouter_membres()  
    elif choix == 7:
        print("Vous avez choisi la Suppression de membres d'un groupe via fichier csv ........")
        supprimer_membres()
    elif choix == 8:
        print("Vous avez choisi la Suppression de tous les membres d'un groupe ........")
        supprimer_all_membres()
    elif choix == 9:
        print("Vous avez choisi la Suppression de groupes ........")
        supprimer_groupes()
       
    elif choix == 99:
        quit()

    else:
        print("Choix invalide")
    
    print()
    menu()
    print()
    choix = int(input("Entrer votre choix "))

print("Merci d'avoir utilisé le programme. Bye. ")


