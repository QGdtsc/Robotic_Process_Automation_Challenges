# Importer les dépendances
import os
import pandas as pd
import datetime as dt
import time
import re
import urllib.request
from selenium import webdriver
import cv2
import pytesseract
from pytesseract import Output


# Définir les chemins des dossiers/fichiers. Ces chemins doivent être modifiés et adaptés à la machine locale pour être fontionnels.
PATH_TO_TESSERACT_FILE_FOLDER = '*:/*************/Tesseract-OCR/' # Chemin du dossier où est installé Tesseract-OCR. Doit être installé préalablement
PATH_TO_RPA_CHALLENGE_INVOICE_EXTRACTION_FOLDER = '*:/*****/*****/*******/RPA/RPA_challenge_Invoice_Extraction/' # Chemin du projet
PATH_TO_IMG_DOWNLOAD_FOLDER = '*:/*****/*****/*******/RPA/' # Chemin du dossier dans lequel les images des factures sont téléchargées 
PATH_TO_IMG_FOR_OCR_FOLDER = '*:/*****/*****/*******/RPA/RPA_challenge_Invoice_Extraction/img_for_ocr/' # Chemin dans lequel les factures téléchargés sont déplacées pour être traitées
PATH_TO_WEBDRIVERS_FOLDER = '*:/Webdrivers/' # Chemin du dossier dans lequel sont stocké les Webdrivers.


# Renseigner le chemin du fichier executable tesseract.exe
pytesseract.pytesseract.tesseract_cmd = PATH_TO_TESSERACT_FILE_FOLDER + 'tesseract.exe'



# Definir les fonctions

def change_file_directory(file_name):
    source = PATH_TO_IMG_DOWNLOAD_FOLDER + str(file_name)
    destination = PATH_TO_IMG_FOR_OCR_FOLDER + str(file_name)
    try:
        if os.path.exists(destination):
            print("There is already a file there")
        else:
            os.replace(source,destination)
            print(source+" was moved")
    except FileNotFoundError:
        pass


def remove_file(file, location):
    try:
        path = os.path.join(location, file)  
        os.remove(path)
        print (path + " has been removed")
    except FileNotFoundError:
        pass



# Modifications d'options pour supprimer un message à la fin du programme
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])


# Définir le chemin de chromedriver.exe nécessaire à l'automatisation du navigateur Chrome
chrome_driver_path = PATH_TO_WEBDRIVERS_FOLDER + "chromedriver.exe" # le fichier chromedriver.exe doit être téléchargé préalablement



# Démarrage de la procedure de RPA et Ouverture du navigateur
invoice_extraction_challenge = webdriver.Chrome(executable_path=chrome_driver_path, options=options)


# Maximiser la fenetre
invoice_extraction_challenge.maximize_window()


# Atteindre la page du rpachallenge
invoice_extraction_challenge.get('https://www.rpachallenge.com/')


# Atteindre l'onglet Invoice Extraction
invoice_extraction_challenge.find_element_by_xpath('/html/body/app-root/div[1]/nav/div/ul/li[4]/a').click()


# Creation du dataframe pour soumettre le resultat
df_resultat = pd.DataFrame(columns=['ID', 'DueDate', 'InvoiceNo', 'InvoiceDate', 'CompanyName', 'TotalDue'])


# Cliquer sur le boutton Start pour démarrer le challenge et le compte à rebours
invoice_extraction_challenge.find_element_by_xpath('//*[@id="start"]').click()


time.sleep(0.1) 


# Atteindre le haut de la page
element_to_reach = invoice_extraction_challenge.find_element_by_xpath('//*[@id="sandbox"]')
invoice_extraction_challenge.execute_script("arguments[0].scrollIntoView();", element_to_reach)


time.sleep(0.1) 


# Lister les numéros de pages de la table contenant les factures 
list_pages = list()
list_elements = invoice_extraction_challenge.find_elements_by_class_name('paginate_button')
# Extraire les valeurs pertinentes et les ajouter à list_pages
for i in list_elements:
    if i.get_attribute('innerHTML').isnumeric() == True:
        list_pages.append(i.get_attribute('innerHTML'))



# Boucler dans la liste de pages (list_pages)
for i in list_pages:

    # Accéder à la page numéro i
    invoice_extraction_challenge.find_element_by_xpath('//*[@id="tableSandbox_paginate"]/span/a['+ str(i) +']').click()

    # Extraire les données de la table
    table = invoice_extraction_challenge.find_element_by_xpath('//*[@id="tableSandbox"]')



    # Boucler dans les element extraits de la table (ligne par ligne)
    for i in table.find_elements_by_xpath('.//tr'):
        if '</th>' in i.get_attribute('outerHTML'): # Identifier les titre (noms de colonnes) de la table
            # Variable table_bis pour collecter les élements de la table
            table_bis = re.findall(r'<th\b[^>]*>([^<]*)<\/th>',i.get_attribute('outerHTML'))
            # Création d'un dataframe temporaire df_table pour stocker les elements de la table
            df_table = pd.DataFrame(columns = table_bis)
        else: # Identifier les valeurs/données de la table
            # Variable table_bis pour collecter les élements de la table
            table_bis = re.findall(r'<td\b[^>]*>([^<]*)<\/td>',i.get_attribute('outerHTML'))
            # Variable href_invoice pour collecter le lien pour accéder à la facture
            href_invoice = re.findall(r'<a[\s\S]*?href="([^"]+)"[\s\S]*?>',i.get_attribute('outerHTML'))
            # Ajouter href_invoice de la facture à table_bis
            table_bis.append(href_invoice[0])
            # Ajouter l'information table_bis (ligne) au dataframe temporaire df_table
            df_table.loc[len(df_table)] = table_bis



    # Ajouter une colonne 'A traiter' dans le dataframe temporaire df_table
    df_table.insert(4,'A traiter','')
    # Decider si la facture doit etre traitee ou non
    for i in range(0,len(df_table)):
        due_date_invoice = str(df_table.iloc[i,2])
        due_date_invoice = dt.datetime.strptime(due_date_invoice, '%d-%m-%Y')
        if due_date_invoice < dt.datetime.now():
            df_table.iloc[i,4] = 'Oui'
        else:
            df_table.iloc[i,4] = 'Non'



    # Boucler dans le dataframe temporaire df_table
    for i in range(0,len(df_table)):
        # Si la facture est à traiter : 
        # Extraire les informations liées à celle-ci et cliquer sur le lien de la facture
        # Puis télécharger la facture sous forme d'image png et la stocker dans le dossier img_for_ocr
        # Ensuite executer la procédure OCR
        if df_table.iloc[i,4] == 'Oui':
            # Extraire les information et cliquer sur le lien de la facture
            due_date_invoice = df_table.iloc[i,2]
            invoice_id = df_table.iloc[i,1]
            href_link = df_table.iloc[i,3]
            invoice_extraction_challenge.find_element_by_xpath('//a[@href="'+ href_link +'"]').click()
            # Atteindre l'onglet de la facture qui vient de s'ouvrir
            invoice_extraction_challenge.switch_to.window(invoice_extraction_challenge.window_handles[-1])
            # Télécharger l'image et la stocker dans le dossier img_for_ocr
            name_image = df_table.iloc[i,1]
            image_selector = invoice_extraction_challenge.find_element_by_xpath('/html/body/img').get_attribute('src')
            file_name = str(name_image) + '.png'
            urllib.request.urlretrieve(image_selector, file_name)
            change_file_directory(file_name)



            # Procédure OCR
            # Définir le chemin de l'image correspondant à la facture
            img_path = PATH_TO_IMG_FOR_OCR_FOLDER + file_name

            # Lecture de l'image avec cv2
            img = cv2.imread(img_path)
            hImg, wImg, _ = img.shape

            # Transformation de l'image en data
            d = pytesseract.image_to_data(img, output_type= Output.DICT)
            data_text = d['text']
            data_text = list(filter(None, data_text))
            number_of_elements_in_data_text = len(data_text)


            # Boucler dans les data extraites de l'image et identifier/extraire les informations necessaires
            for i in range(int(number_of_elements_in_data_text)):
                if data_text[i] == 'INVOICE':
                    company_name = ' '.join([data_text[n] for n in range(0,i)])
                if data_text[i] == 'Invoice' and ('#') in data_text[i+1]:
                    invoiceNo = data_text[i+1]
                    invoiceNo = str(invoiceNo).replace('#','')
                if data_text[i] == 'Total' and str(data_text[i+1]).replace('.','').isnumeric() == True:
                    total = data_text[i+1]
                try:
                    invoiceDate = dt.datetime.strptime(str(data_text[i]),'%Y-%m-%d').strftime('%d-%m-%Y') 
                except:
                    pass



            # Enregister les donnees extraites de l'image (la facture) dans le dataframe df_resultat
            df_resultat.loc[len(df_resultat)] = [invoice_id,due_date_invoice,invoiceNo,invoiceDate,company_name,total]

            # Supprimer le fichier/l'image
            remove_file(file_name, PATH_TO_IMG_FOR_OCR_FOLDER)



            # Fermer la fenetre / l'onglet actif contenant l'image
            invoice_extraction_challenge.close()


            # Atteindre la fenetre / l'onglet principal du rpachallenge
            invoice_extraction_challenge.switch_to.window(invoice_extraction_challenge.window_handles[0])



# Creer le fichier csv resultat.csv
df_resultat.to_csv(PATH_TO_RPA_CHALLENGE_INVOICE_EXTRACTION_FOLDER + 'resultat.csv',sep=',', index=False)


# Soumettre le resultat (resultat.csv) par injection html
invoice_extraction_challenge.find_element_by_xpath('//*[@id="submit"]/div/div/div/form/input[1]').send_keys(PATH_TO_RPA_CHALLENGE_INVOICE_EXTRACTION_FOLDER + 'resultat.csv')


# Attendre 1 secondes
time.sleep(1)


# Supprimer le fichier resultat.csv
remove_file('resultat.csv', PATH_TO_RPA_CHALLENGE_INVOICE_EXTRACTION_FOLDER)


# Attendre 2 secondes et fermer le navigateur Chrome
time.sleep(2)
invoice_extraction_challenge.close()
