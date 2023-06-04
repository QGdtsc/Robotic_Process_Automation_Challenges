# Importer les dépendances
from selenium import webdriver
import os.path
import time
from RPA.Excel.Files import Files # the Excel.Files library can be used to read and write Excel files without the need to start the actual Excel application



# Définir les chemins des dossiers/fichiers
PATH_TO_EXCEL_WORKBOOK_FOLDER = "*:/*****/*****/*******/***/RPA_challenge_Input_Forms/" # Le chemin doit être modifié pour être fonctionnel
PATH_TO_DOWNLOAD_FOLDER = "*:/*****/*****/Downloads/" # Le chemin doit être modifié pour être fonctionnel
PATH_TO_WEBDRIVERS_FOLDER = "*:/Webdrivers/" # Le chemin doit être modifié pour être fonctionnel



# Définir les fonctions

def verification_file_already_in_folder():
    if os.path.exists(PATH_TO_EXCEL_WORKBOOK_FOLDER + 'challenge.xlsx'):
        file_already_in_folder = True
    else:
        file_already_in_folder = False
    return file_already_in_folder


def wait_for_downloads():
    print("Waiting for downloads")
    while any([filename.endswith(".crdownload") for filename in 
               os.listdir(PATH_TO_DOWNLOAD_FOLDER)]): # variable PATH_TO_DOWNLOAD_FOLDER : Chemin du dossier Téléchargements
        time.sleep(2)
        print(".")
    print("DONE")


def change_file_directory():
    print('Change directory of the downloaded file')
    source = PATH_TO_DOWNLOAD_FOLDER + "challenge.xlsx" # variable PATH_TO_DOWNLOAD_FOLDER : Chemin du dossier Téléchargements
    destination = PATH_TO_EXCEL_WORKBOOK_FOLDER + "challenge.xlsx" # variable "PATH_TO_EXCEL_WORKBOOK_FOLDER" : Chemin du dossier dans lequel le fichier Excel du challenge doit être déplacé
    try:
        if os.path.exists(destination):
            print("There is already a file there")
        else:
            os.replace(source,destination)
            print(source+" was moved")
    except FileNotFoundError:
        print(source+" was not found")    


def extract_Excel_information():
    excel = Files()
    excel.open_workbook(PATH_TO_EXCEL_WORKBOOK_FOLDER + "challenge.xlsx") # variable "PATH_TO_EXCEL_WORKBOOK_FOLDER" : Chemin du dossier dans lequel est stocké le fichier Excel du challenge
    table_excel_information = excel.read_worksheet_as_table(header=True)
    return table_excel_information


def remove_file(file, location):
    path = os.path.join(location, file)  
    os.remove(path)
    print ("The file has been removed")



# Modifications d'options pour supprimer un message à la fin du programme
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])


# Définir le chemin de chromedriver.exe nécessaire à l'automatisation du navigateur Chrome
chrome_driver_path = PATH_TO_WEBDRIVERS_FOLDER + "chromedriver.exe" # le fichier chromedriver.exe doit être téléchargé préalablement



# Démarrage de la procedure de RPA et Ouverture du navigateur
rpa_challenge = webdriver.Chrome(executable_path=chrome_driver_path, options=options)


# Maximiser la fenetre
rpa_challenge.maximize_window()


# Atteindre la page du rpachallenge
rpa_challenge.get('https://www.rpachallenge.com/')


# Verifier si le fichier challenge.xlsx (contenant les informations pour compléter le challenge) est déjà téléchargé et placé dans le dossier.
# Si ce n'est pas le cas, procéder au téléchargement et le placer dans le dossier.
file_already_in_folder = verification_file_already_in_folder()
if file_already_in_folder == True:
    print("Le fichier est déjà téléchargé dans le dossier")
if file_already_in_folder == False:
    print("Le fichier doit être téléchargé placé dans le dossier")
    # Cliquer sur le bouton télécharger
    rpa_challenge.find_element_by_xpath('/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a').click()
    # Lancement de la fonction wait_for_download
    wait_for_downloads()
    time.sleep(1)
    # Lancement de la fonction change_file_directory
    change_file_directory()


# Extraire les informations du fichier Excel
table_excel_information = extract_Excel_information()


# Cliquer sur le bouton pour démarrer le challenge (et démarrer le compte à rebours)
rpa_challenge.find_element_by_class_name('btn-large.uiColorButton').click()


# Démarrage de la boucle pour lire/stocker les information et les restituer dans le navigateur Chrome. Element par élément (ligne par ligne).
for i in table_excel_information:
    # Lire et stocker les informations du fichier Excel
    firstname = list(i.values())[0]
    lastname = list(i.values())[1]
    companyname = list(i.values())[2]
    roleincompany = list(i.values())[3]
    address = list(i.values())[4]
    email = list(i.values())[5]
    phone = list(i.values())[6]
    # Restituer les informations dans le champ correspondant dans la page du navigateur Chrome.
    rpa_challenge.find_element_by_xpath('//*[@ng-reflect-name="labelFirstName"]').send_keys(firstname)
    rpa_challenge.find_element_by_xpath('//*[@ng-reflect-name="labelLastName"]').send_keys(lastname)
    rpa_challenge.find_element_by_xpath('//*[@ng-reflect-name="labelEmail"]').send_keys(email)
    rpa_challenge.find_element_by_xpath('//*[@ng-reflect-name="labelCompanyName"]').send_keys(companyname)
    rpa_challenge.find_element_by_xpath('//*[@ng-reflect-name="labelRole"]').send_keys(roleincompany)
    rpa_challenge.find_element_by_xpath('//*[@ng-reflect-name="labelAddress"]').send_keys(address)
    rpa_challenge.find_element_by_xpath('//*[@ng-reflect-name="labelPhone"]').send_keys(phone)
    # Cliquer sur le bouton SUBMIT pour procédér à l'élément suivant.
    rpa_challenge.find_element_by_class_name('btn.uiColorButton').click()


# Extraire les informations liées au résultat du challenge
print(rpa_challenge.find_element_by_xpath('/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[1]').get_attribute('innerText'))
print(rpa_challenge.find_element_by_xpath('/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[2]').get_attribute('innerText'))


# Suppression du fichier Excel challenge.xlsx
remove_file('challenge.xlsx', PATH_TO_EXCEL_WORKBOOK_FOLDER)


# Attendre 2 secondes et fermer le navigateur internet
time.sleep(2)
rpa_challenge.close()
