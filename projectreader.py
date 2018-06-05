import os
import json
from xlsconverter import xlsx_doc, web_renderer


def clear():
    os.system('cls' if os.name=='nt' else 'clear')



class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


class session:
    
    def __init__(self, name, output):
        self.name = name
        self.session_file = output
        self.project_list = []
        self.handler_0 = project_handler()
    
    
    def selection(self):

        # Si fichier de sauvegarde de session existant : verifier son contenu.
        if os.path.isfile(self.session_file):
 
            # Si fichier de sauvegarde de session vide : création d'une nouvelle session.
            if os.stat(self.session_file).st_size == 0:  
                self.choix_session = 2
            
            # Detection d'une session précedente dans le fichier de sauvegarde de session.
            else:
                while True:
                    try:
                        self.choix_session = int(input('''\nATTENTION !
Une session précédente à été détectée, voulez vous reprendre cette session ?

    1. Reprendre la session.
    2. Effacer la session et en démarrer une nouvelle.

>> '''))
                        break
                    except ValueError:
                        input("\nEntrez un numéro de menu valide.")
                        clear()

            return self.choix_session
        
        # Fichier de sauvegarde inexistant : création d'une nouvelle session.
        else:
            if not os.path.exists("./session"):
                os.makedirs("./session")
            self.choix_session = 2



    def run(self):
    
        while self.choix_session != 0:
	
            # NOUVELLE SESSION
            if self.choix_session == 2: 

                clear()
                input("Le dossier projet actuel est : /projet\n\nAppuyez sur une touche pour charger les projets.")
                self.handler_0.load_projects(self)
                input("\nAppuyez sur une touche pour commencer l'evaluation.")
                clear()
                #self.print_project_list(self.project_list)
                self.handler_0.eval_all_projects(self)
                self.choix_session = 3
		
            # REPRISE SESSION
            elif self.choix_session == 1: 
                self.load()
                self.choix_session = 3
        
            # MENU GENERAL DE SESSION
            elif self.choix_session == 3: # Session de selection d'options après chargement   

                while True:
                    clear()
                    try:
                        choix_projets = int(input('''Que voulez vous faire ?\n
    1. Reprendre l'évaluations des projets.
    2. Enregistrer les évaluations sous la forme d'un fichier texte.
    3. Enregistrer les evaluations dans le classeur de section.
    4. Rechercher une evaluation.
    5. Quitter.

>> '''))
                        break
                    except ValueError:
                        input("\nEntrez un numéro de menu valide.")


                # MENU REPRISE EVALUATION     
                if choix_projets == 1:
                    self.handler_0.resume_eval(self)
        
                # MENU SORTIE TXT   
                elif choix_projets == 2: 
                    self.txt_file = "./sortie/Evaluations.txt"
                    
                    if not os.path.exists("./sortie"):
                        os.makedirs("./sortie")        
                    
                    self.handler_0.print_txt_file(self.txt_file, self.project_list)

                # MENU SORTIE CLASSEUR EXTERNE
                elif choix_projets == 3: 
                    self.xlsx_doc_0 = xlsx_doc("document")
                    choix_xlsx = 0
                    clear()
                    print("Veuillez charger un classeur afin d'y stocker les évaluation.\n")
                    self.xlsx_doc_0.load()

                    while choix_xlsx != 3:
                        clear()
                        print("Que souhaitez vous faire ?\n\n1. Enregistrer les évaluation.\n2. Enregistrer le classeur sous la forme d'une page web.\n3. Retour")
                        choix_xlsx = int(input("\n>> "))
                        
                        if choix_xlsx == 1:
                            self.xlsx_doc_0.map(self.project_list)
                            self.xlsx_doc_0.fill(self.project_list)
                            self.xlsx_doc_0.write()

                        if choix_xlsx == 2:
                            self.web_0 = web_renderer("web",self.xlsx_doc_0)
                            self.web_0.style_select()
                            self.web_0.load_templates()
                            self.web_0.render()

                # MENU RECHERCHE PROJET
                elif choix_projets == 4: 
                    self.handler_0.find_project(self)
        
                # MENU QUITTER
                elif choix_projets == 5:
                    self.choix_session = 0 
            
                # MENU AFFICHER DEBUG LIST
                elif choix_projets == 6:
                    input(self.project_list)

                else:
                    input("\nEntrez un numéro de menu valide.")
    
    
    def save(self):
        with open(self.session_file,"w+") as fichier:
            json.dump(self.project_list, fichier,indent=2)		
        print("\nSession sauvegardée\n")


    def load(self):
        with open(self.session_file,"r+") as fichier:
            self.project_list = json.load(fichier)
        input("\nSession chargée. Appuyez sur une touche pour continuer.\n")
        return(self.project_list)




class project_handler:

    def __init__(self):
        self.dir=input("Entrez le chemin d'accès du dossier de projets:\n> ")
        #self.dir = "./projets"
        self.projet = {"NOM" : "", "ETUDIANTS": [] , "SECTION" : "","FICHIER" : "", "NOTE": "Aucune note", "COMMENTAIRE" : "Aucun commentaire", "CHECKSUM": 0}


    def load_projects(self,session): # Chargement des fichiers projets et stockage dans les structures de donnees.
        
        session.project_list = []
        liste_etudiant_projet =[]
        fichiers=[fichier for fichier in os.listdir(self.dir)]

        for index,fichier in enumerate(fichiers):
            
            liste_etudiant_projet = fichier.split("_")[1]
            liste_etudiant_projet = liste_etudiant_projet.split("-")
        
            self.projet["NOM"] = fichier.split("_")[2]
            self.projet["ETUDIANTS"] = liste_etudiant_projet
            self.projet["SECTION"] = fichier.split("_")[0]
            self.projet["FICHIER"] = fichier
            session.project_list.append(self.projet.copy())
            
        print("Projets charges :", len(session.project_list))
        return session.project_list


    def print_project_list(self,liste_projets):   
        for projet in liste_projets:
            self.print_project(projet)

            
    def print_project(self,projet):   
            print(projet["NOM"], "réalisé par", ", ".join(projet["ETUDIANTS"]), "en section", projet["SECTION"], "\n" , "NOTE: ", projet["NOTE"],"\nCommentaire: ", projet["COMMENTAIRE"])
            input("")

        
    def eval_project(self,projet):
        
        etudiant_eval = "\nEvaluer le projet de: "
        for etudiants in projet["ETUDIANTS"]:
            etudiant_eval = etudiant_eval + etudiants + ", "
            
        etudiant_eval = etudiant_eval + "en section " + projet["SECTION"] + ":\n"
        print(etudiant_eval)
        projet["NOTE"] = input("    NOTE ? >> ")
        projet["COMMENTAIRE"] = input("    COMMENTAIRE ? >> ")
        projet["CHECKSUM"] = 1
        
        return(projet)


    def eval_all_projects(self,session): # Lance l'evaluation des projets, et retourne la liste mise à jour.
        for projet in session.project_list:
            projet = self.eval_project(projet)
            session.save()
        return session.project_list

        
    def resume_eval(self,session):
        for projet in session.project_list:
            if projet["CHECKSUM"] == 1:
                continue
            else:
                projet = self.eval_project(projet)
                session.save()
        input("\nTout les projets ont étés évalués.")
        return(session.project_list)


    def print_txt_file(self, txt_file, liste_projets):
        with open(txt_file,"w+") as _txt_:
            for projet in liste_projets:
                buffer = "\n\n" + projet["NOM"] + " réalisé par " + ", ".join(projet["ETUDIANTS"]) + ", en section " + projet["SECTION"] +":\n    NOTE: " + projet["NOTE"] + "\n    Commentaire: " + projet["COMMENTAIRE"] + "\n\n"
                _txt_.write(buffer)

        input("\nFichier Evaluations.txt enregistré dans le dossier Sorties.")


    def find_project(self,session):
        while True:
            clear()
            try:
                choix = int(input('''Quelle recherche souhaitez vous effectuer ?
        
        1. Rechercher par nom.
        2. Rechercher par section.
        3. Rechercher par note.
        4. Rechercher par commentaire.
        5. Retour.
        
    >> '''))
                break
            except ValueError:
                input("\nEntrez un numéro de menu valide.")



        if choix == 1:
            clear()
            resultats =[]
            recherche = input("Entrez le nom de l'etudiant que vous recherchez.\n\n>> ")
                
            for projet in session.project_list:
                for etudiant in projet["ETUDIANTS"]:
                    
                    if etudiant != recherche:
                        continue
                    else:
                        if projet in resultats:
                            break
                        else:
                            resultats.append(projet)
                    
            if resultats:
                print(len(resultats) , " projets trouvé(s) :\n")

                for index, resultat in enumerate(resultats):
                    
                    print("    " + str(index+1) + ".", resultat["SECTION"] +":", end=" " )
                    for nom in resultat["ETUDIANTS"]:
                        if nom != recherche:
                            print(nom, end=", ")
                        else:
                            print(bcolors.OKBLUE + nom + bcolors.ENDC, end=", ")

                    print( "\n        NOTE: " + resultat["NOTE"] , "\n        COMMENTAIRE:", resultat["COMMENTAIRE"],"\n")
                print("    0. Retour\n")

                
                while True:
                    try:
                        selection_index = int(input("Entrez votre selection >> "))
                        break
                    except ValueError:
                        input("\nEntrez un numéro de menu valide.")    
                print("")


                if selection_index !=0:
                    self.print_project(resultats[selection_index - 1])
                    
                    while True:
                    
                        try:
                            choix_2 = int(input("\nQue souhaitez vous faire ?\n1. Réevaluer le projet ?\n2. Retour\n>> "))
                            break
                        except ValueError:
                            input("\nEntrez un numéro de menu valide.")
            
                    if choix_2 == 1:
                        self.eval_project(resultats[selection_index - 1])
                        session.save()
                
                else:
                    pass
                
            if resultats == []:                
                input("\nAucun projet trouvé.")
            else:
                input("\nFin de la recherche")

    
    
    
        if choix == 2:
            clear()
            resultats = []
            recherche = input("Entrez le nom de la section que vous recherchez.\n\n>> ")
            
            for projet in session.project_list:
                if projet["SECTION"] != recherche:
                    continue
                else:
                    if projet in resultats:
                        break
                    else:
                        resultats.append(projet)  
            if resultats:
                print(len(resultats) , " projets trouvé(s) :\n")
                for index, resultat in enumerate(resultats):
                    print("    " + str(index+1) + ".", bcolors.OKBLUE + resultat["SECTION"] +bcolors.ENDC +":", ", ".join(resultat["ETUDIANTS"]), "\n        NOTE: " + resultat["NOTE"], "\n        COMMENTAIRE:", resultat["COMMENTAIRE"],"\n")

                while True:
                    try:
                        selection_index = int(input(" Choisissez un projet >> ")) - 1
                        break
                    except ValueError:
                        input("\nEntrez un numéro de menu valide.")

                print("")
                self.print_project(resultats[selection_index])
                choix_2 = int(input("\nQue souhaitez vous faire ?\n1. Réevaluer le projet ?\n2. Retour\n>> "))
        
                if choix_2 == 1:
                    self.eval_project(resultats[selection_index])
                    session.save()


                            
            if resultats == []:                
                input("\nAucun projet trouvé.")
            else:
                input("\nFin de la recherche")


        if choix == 3:
            clear()
            resultats = []
            recherche = input("Entrez la note que vous recherchez.\n\n>> ")
            
            for projet in session.project_list:
                if projet["NOTE"] != recherche:
                    continue
                else:
                    if projet in resultats:
                        break
                    else:
                        resultats.append(projet)


            if resultats:
                print(len(resultats) , " projets trouvé(s) :\n")
                for index, resultat in enumerate(resultats):
                    print("    " + str(index+1) + ".", resultat["SECTION"] +":", ", ".join(resultat["ETUDIANTS"]), "\n        NOTE: " + bcolors.OKBLUE + resultat["NOTE"] + bcolors.ENDC, "\n        COMMENTAIRE:", resultat["COMMENTAIRE"],"\n")

                while True:
                    try:
                        selection_index = int(input(" Choisissez un projet >> ")) - 1
                        break
                    except ValueError:
                        input("\nEntrez un numéro de menu valide.")

                print("")
                self.print_project(resultats[selection_index])
                choix_2 = int(input("\nQue souhaitez vous faire ?\n1. Réevaluer le projet ?\n2. Retour\n>> "))
        
                if choix_2 == 1:
                    self.eval_project(resultats[selection_index])
                    session.save()


            if resultats == []:                
                input("\nAucun projet trouvé.")
            else:
                input("\nFin de la recherche")

        if choix == 4:
            clear()
            resultats = []
            recherche = input("Entrez le commentaire que vous recherchez.\n\n>> ")
            
            for projet in session.project_list:
                for word in projet["COMMENTAIRE"].split():
                    if word != recherche:
                        continue
                    else:
                        if projet in resultats:
                            break
                        else:
                            resultats.append(projet)

            if resultats:
                print(len(resultats) , " projets trouvé(s) :\n")
                for index, resultat in enumerate(resultats):
                    print("    " + str(index+1) + ".", resultat["SECTION"] +":", ", ".join(resultat["ETUDIANTS"]), "\n        COMMENTAIRE:", end=" ")
                    for word in resultat["COMMENTAIRE"].split():
                        if word != recherche:
                            print(word, end=" ")
                        else:
                            print(bcolors.OKBLUE + word + bcolors.ENDC, end=" ")

                    print("\n")

                while True:
                    try:
                        selection_index = int(input(" Choisissez un projet >> ")) - 1
                        break
                    except ValueError:
                        input("\nEntrez un numéro de menu valide.")

                print("")
                while True:

                    try:
                        clear()
                        self.print_project(resultats[selection_index])
                        choix_2 = int(input("\nQue souhaitez vous faire ?\n1. Réevaluer le projet ?\n2. Retour\n>> "))
                        if choix_2 == 1:
                            self.eval_project(resultats[selection_index])
                            session.save()

                        break
                    except ValueError:
                        input("\nEntrez un numéro de menu valide.")
            
            if resultats == []:                
                input("\nAucun projet trouvé.")
            else:
                input("\nFin de la recherche")
