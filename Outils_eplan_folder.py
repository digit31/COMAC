import fitz,pyautogui,os ,PyPDF2,subprocess
from tkinter import Tk, filedialog,simpledialog
from tkinter.filedialog import askdirectory
from zipfile import ZipFile
from PIL import Image


def reduire_taille_pdf(fichier_pdf,taille_reduction):
    """Assurez-vous que Ghostscript est installé sur votre système. 
    Vous pouvez le télécharger depuis le site officiel : 
    https://www.ghostscript.com/download/gsdnld.html
    changer la variable chemin_ghostscript 
    """
    if fichier_pdf:
        try:
            # Créez un nouvel objet PDF réduit
            nouveau_nom = f"reduced_{taille_reduction}pct_{os.path.basename(fichier_pdf)}"
            new_pdf_path = os.path.join(os.path.dirname(fichier_pdf), nouveau_nom)

            # Spécifiez correctement le chemin vers l'exécutable Ghostscript
            chemin_ghostscript = r"C:\Program Files (x86)\gs\gs10.02.1\bin\gswin32c.exe"

            # Utilisez Ghostscript pour compresser le fichier PDF avec la taille spécifiée
            subprocess.run([chemin_ghostscript, "-sDEVICE=pdfwrite", f"-dCompatibilityLevel=1.4", "-dNOPAUSE", "-dBATCH", f"-dPDFSETTINGS=/screen", f"-dAutoRotatePages=/None", f"-dDownsampleColorImages=true", f"-dColorImageDownsampleThreshold={taille_reduction}", f"-sOutputFile={new_pdf_path}", fichier_pdf])

            print(f"La taille du fichier a été réduite de {taille_reduction}%. Nouveau fichier : {new_pdf_path}")

        except Exception as e:
            print(f"Une erreur s'est produite : {e}")

def supprimer_derniere_page(fichier_source):
    # Ouvrir le fichier source PDF en mode binaire
    with open(fichier_source, "rb") as source_file:
        # Créer un objet PdfReader pour lire le fichier source
        pdf_reader = PyPDF2.PdfReader(source_file)

        # Créer un objet PdfWriter pour écrire dans le fichier destination
        destination_pdf_writer = PyPDF2.PdfWriter()

        # Ajouter toutes les pages sauf la dernière
        for page_num in range(len(pdf_reader.pages) - 1):
            page = pdf_reader.pages[page_num]
            destination_pdf_writer.add_page(page)

    # Écrire les modifications dans le fichier destination
    with open(fichier_source, "wb") as destination_file:
        destination_pdf_writer.write(destination_file)
    print(f"La dernière page a été supprimée avec succès pour le fichier {os.path.basename(fichier_source)}.")

def create_zip(folder_path, files_to_include):
    name_folder = os.path.basename(folder_path)
    ref_eplan , ref_folder =name_folder.split('-EXE-')
    zip_name = os.path.join(folder_path, f'{ref_eplan}.zip')
    with ZipFile(zip_name, 'w') as zipf:
        for file_name in files_to_include:
            file_name = F"{ref_folder}{file_name}"
            file_path = os.path.join(folder_path,file_name)
            if os.path.isfile(file_path):
                zipf.write(file_path, arcname=file_name)
            elif os.path.isdir(file_path):
                for root, _, files in os.walk(file_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, folder_path)
                        zipf.write(file_path, arcname=arcname)

def fusionner_pages(fichier_destination, fichier_sources, position_insertion):
    # Ouvrir le fichier source PDF en mode binaire
    with open(fichier_destination, "rb") as source_file:
        # Créer des objets PdfReader pour lire le fichier source
        pdf_reader_source = PyPDF2.PdfReader(source_file)

        # Créer un objet PdfWriter pour écrire dans le fichier source
        destination_pdf_writer = PyPDF2.PdfWriter()

        # Ajouter les pages du fichier source jusqu'à la position désirée
        for page_num in range(position_insertion):
            page = pdf_reader_source.pages[page_num]
            destination_pdf_writer.add_page(page)

        # Ajouter les pages des fichiers sources à partir de la position désirée
        for fichier_source in fichier_sources:
            with open(fichier_source, "rb") as additional_file:
                pdf_reader_additional = PyPDF2.PdfReader(additional_file)
                for page_num in range(1, len(pdf_reader_additional.pages)):
                    if page_num == 1:
                        page = pdf_reader_additional.pages[page_num]
                        destination_pdf_writer.add_page(page)
            print(f"Pages de {fichier_source} ajoutées avec succès")
        # Ajouter les pages restantes du fichier source
        for page_num in range(position_insertion, len(pdf_reader_source.pages)):
            page = pdf_reader_source.pages[page_num]
            destination_pdf_writer.add_page(page)

    # Écrire les modifications dans le fichier destination
    with open(fichier_destination, "wb") as destination_file:
        destination_pdf_writer.write(destination_file)

def optimize_images_in_place(folder_path, target_size=(800, 800), quality=85):
    for filename in os.listdir(folder_path):
        if filename.endswith(('.jpg', '.jpeg', '.png')):
            file_path = os.path.join(folder_path, filename)

            with Image.open(file_path) as img:
                # Redimensionner l'image
                img.thumbnail(target_size)

                # Enregistrer l'image avec la qualité spécifiée (écraser l'originale)
                img.save(file_path, 'JPEG', quality=quality)


    
flag_main = pyautogui.confirm('Veuillez sélectionner une option \n MERCI DE LIRE LES MESSAGES (les prints) SUR LA CONSOLE', 
                            buttons=["faire la fusion des pdf",
                                    "Supprimée La dernière page du document pdf",
                                    "montage du dossier zip",
                                    "optimize les images",
                                    "optimize les PDF",
                                    "annuler"])

if flag_main == "Supprimée La dernière page du document pdf":
    path_of_pdf = filedialog.askopenfilenames(title="Sélectionner les fichiers pdf", filetypes=[('pdf', "PDF*")])
    for pdf_to_del_page in path_of_pdf:
        supprimer_derniere_page(pdf_to_del_page)

if flag_main ==  "annuler":
    print ("BAY BAY BAY")

if flag_main == "faire la fusion des pdf":
    # Sélectionner les fichiers PDF
    path_of_all_pdf = filedialog.askopenfilenames(title="Sélectionner les fichiers pdf", filetypes=[('pdf', "PDF*")])
    all_path_list = list()
    for path in  path_of_all_pdf:
        all_path_list.append(path)
    dir_name = os.path.dirname(all_path_list[0])
    dir_name = os.path.realpath(dir_name)
    path_files_output = []
    path_file_export = None
    path_new_output = None
    position = 2
    path_of_output_1 = None
    nom_fichier_pdf = None

    # Séparer les fichiers de sortie et d'export

    for pdf in all_path_list:
        if 'output' in os.path.basename(pdf).lower():
            path_files_output.append(pdf)
            if 'output_1' in os.path.basename(pdf).lower():
                path_of_output_1 = pdf 
        elif 'exportpdf' in os.path.basename(pdf).lower():
            path_file_export = pdf
            
    nom_fichier_pdf = os.path.basename(path_file_export)
    nom_fichier_pdf_output= nom_fichier_pdf.split('_')[0]
    all_path_list.remove(path_of_output_1)
    all_path_list.remove(path_file_export)

    # Vérifier s'il y a plusieurs fichiers export
    if len (path_file_export) !=1 :
        print("Il y a plusieurs fichiers export")

    # Fusionner les fichiers output et export

    path_plan_expot = os.path.join(dir_name, f"{nom_fichier_pdf_output}_Plan.pdf")
    
    new_pdf_merger = fitz.open()
    new_pdf_merger.insert_pdf(fitz.open(path_of_output_1))
    new_pdf_merger.insert_pdf(fitz.open(path_file_export))
    new_pdf_merger.save(path_plan_expot)
    new_pdf_merger.close()
    
    
    # Vérifier s'il y a plusieurs fichiers output
    if all_path_list:
        print("Il y a plusieurs fichiers output")
        fichier_sources = all_path_list
        fichier_destination = path_plan_expot
        fusionner_pages(fichier_destination, fichier_sources, position_insertion=2)
    
    print(f"le fichie : {nom_fichier_pdf_output}_Plan.pdf a été créée")

if flag_main == "montage du dossier zip":
    folder_path = askdirectory()
    folder_path = os.path.realpath(folder_path)

    files_to_include = ['_ExportComac.xlsx', '_Plan.pdf ','.pcm', '_photos']

    # Appelez la fonction pour créer le zip
    create_zip(folder_path, files_to_include)
    
if flag_main == "optimize les images":
    folder_path = askdirectory(title="Sélectionner un dossier")
    optimize_images_in_place(folder_path)
    print("done")
    
if flag_main == "optimize les PDF":
    path_of_pdf = filedialog.askopenfilenames(title="Sélectionner les fichiers pdf", filetypes=[('pdf', "PDF*")])
    # Obtenir la taille de réduction souhaitée de l'utilisateur
    taille_reduction = simpledialog.askinteger("Réduction de taille", "Entrez la taille de réduction en pourcentage:", minvalue=1, maxvalue=100)

    for pdf_to_optimize in path_of_pdf:
        reduire_taille_pdf(pdf_to_optimize,taille_reduction)