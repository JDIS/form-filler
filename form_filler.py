import datetime
import pymupdf
from pymupdf import Pixmap
import os
from os.path import isfile, join
import config as cfg


class Form_Filler:
    form = None
    page = None
    base_filename = ""
    pagefill_count = 0
    expected_file_extensions = [".pdf", ".jpg", ".png"]
    files = [
        f
        for f in os.listdir(os.getcwd())
        if (
            isfile(join(os.getcwd(), f))
            and os.path.splitext(f)[1] in expected_file_extensions
        )
    ]

    def check(self, def_X_pos, alt_X_pos, prompt, default="n", yes_char="y"):
        response = input(prompt).lower().strip()
        if (response and response[0] == yes_char) or default == yes_char:
            self.text(def_X_pos, "X")
        else:
            self.text(alt_X_pos, "X")

    def remboursement(self):
        self.text((170, 97), cfg.nom_groupe)
        self.group_type_check()
        name = input("Nom pour remboursement: ")
        self.text((235, 183), name)
        self.text((235, 219), cfg.user_name)
        self.text((110, 245), cfg.user_email)
        self.text((415, 245), cfg.user_phone)
        self.text((200, 285), input("Montant à rembourser: "))
        self.check((329, 273), (329, 299), "Inscrit Accès D? (Y/[N]): ")
        self.check((227, 347), (287, 347), "Pièce justificative originale? ([Y]/N): ", "y")
        self.check((227, 389), (287, 389), "Activité de financement? (Y/[N]): ")
        self.check((227, 416), (287, 416), "Revenu lié à la dépense? (Y/[N]): ")
        project = input("Projet/activité si connu: ")
        self.text((210, 455), project)
        self.textfield((210, 470, 540, 550), "Description de la dépense: ")
        time = datetime.datetime.now()
        self.text((105, 570), time.strftime("%d-%b-%Y"))

        self.signForm((355, 533, 400, 583))
        self.text((220, 630), cfg.nom_presidence)
        self.text((355, 630), cfg.nom_tresorie)
        self.add()
        self.save(f"remboursement {project} {name} {time.strftime('%Hh%M')}.pdf")

    def deposit(self):
        self.text((175, 97), cfg.nom_groupe)
        self.group_type_check()
        name = input("Provenance du dépot: ")
        self.text((250, 175), name)
        self.text((250, 205), cfg.user_name)
        self.text((110, 235), cfg.user_email + "@usherbrooke.ca")
        self.text((420, 235), cfg.user_phone)
        self.text((175, 270), input("Montant déposé: ") + "$")
        method = input(
            "Méthode de dépot? (Argent/Chèque/[Dépot direct]/Square)"
        ).lower()
        method = method[0] if method else "d"
        match method:
            case "a":
                self.text((78, 299), "X")
            case "c":
                self.text((78, 330), "X")
            case "d":
                self.text((225, 299), "X")
            case "self":
                self.text((225, 330), "X")
            case _:
                self.text((320, 330), "X")
        self.check((227, 373), (287, 373), "Activité de financement? (Y/[N]): ")
        self.check((227, 400), (287, 400), "Commandite? (Y/[N]): ")
        self.textfield((210, 440, 540, 520), "Description de l'activité: ")
        time = datetime.datetime.now()
        self.text((105, 570), time.strftime("%d-%b-%Y"))

        self.signForm((425, 575, 470, 625))
        self.text((255, 575), cfg.nom_presidence)
        self.text((425, 575), cfg.nom_tresorie)
        self.save(f"dépot {name} {time.strftime('%Hh%M')}.pdf")

    def gas(self):
        self.text((170, 97), cfg.nom_groupe)
        self.group_type_check()
        name = input("Nom de la personne à rembourser: ")
        self.text((255, 183), name)
        self.text((255, 212), cfg.user_name)
        self.text((110, 240), cfg.user_email + "@usherbrooke.ca")
        self.text((415, 240), cfg.user_phone)
        self.text((200, 280), input("Adresse de destination: "))
        while True:
            try:
                distance = float(input("Distance: "))
                break
            except ValueError:
                print("Veuillez entrer une valeur numérique valide pour la distance.")
        taux_essence = input(f"Taux essence [{cfg.default_gas_rate}]: ")
        taux_essence = cfg.default_gas_rate if not taux_essence else float(taux_essence)
        self.text((128, 335), f"{distance}km")
        self.text((370, 337), f"{taux_essence}$")
        self.text((200, 362), f"{distance * taux_essence:.2f}$")
        self.check((227, 409), (287, 406), "Activité de financement? (Y/[N]): ")
        self.check((227, 436), (287, 433), "Revenu lié à la dépense? (Y/[N]): ")
        project = input("Projet/activité si connu: ")
        self.text((210, 475), project)
        self.textfield((210, 505, 540, 585), "Raison du déplacement: ")
        time = datetime.datetime.now()
        self.text((105, 600), time.strftime("%d-%b-%Y"))
        self.signForm((435, 585, 480, 660))
        self.save(f"essence {project} {name} {time.strftime('%Hh%M')}.pdf")

    def commandite(self):
        self.text((175, 110), cfg.nom_groupe)
        self.group_type_check()
        nom_compagnie = input("Nom de Cie: ")
        self.text((140, 240), nom_compagnie)
        self.text((275, 325), input("Nom de la personne à contacter (Cie): "))
        self.text((275, 350), input("Email (Cie): "))
        self.text((275, 375), input("Téléphone (Cie): "))
        self.text((260, 175), cfg.user_name)
        self.text((115, 200), cfg.user_email + "@usherbrooke.ca")
        self.text((360, 200), cfg.user_phone)
        self.textfield((140, 250, 420, 330), "Adresse: ")
        self.text((240, 415), input("Montant de la commandite: "))
        self.check((277, 455), (337, 455), "Paiment de la commandite? (Chèque/[Dépot direct]): ", "d", "c")
        self.check((277, 500), (337, 500), "Facture requise à émettre pour la Cie? (Y/[N]): ")
        self.check((277, 535), (337, 535), "Reçu demandé par la Cie? (Y/[N]): ")
        time = datetime.datetime.now()
        self.text((105, 590), time.strftime("%d-%b-%Y"))
        self.signForm((425, 640, 470, 680))
        self.text((255, 615), cfg.nom_presidence)
        self.text((425, 615), cfg.nom_tresorie)
        self.save(f"commandite {nom_compagnie} {time.strftime('%Hh%M')}.pdf")

    def group_type_check(self):
        match cfg.type_groupe:
            case "Groupe Technique":
                self.text((78, 140), "X")
            case "Groupe de l'AGEG":
                self.text((228, 140), "X")
            case "Promo":
                self.text((365, 140), "X")
            case "CSG":
                self.text((445, 140), "X")
            case _:
                print("Type de groupe non reconnu, veuillez vérifier la configuration.")

    def text(self, pos, text):
        self.page.insert_text(pos, text)
    
    def textfield(self, rect, prompt):
        while 0 >= self.page.insert_textbox(rect, input(prompt), lineheight=2.25):
            print("Textfield too long")

    def signForm(self, rect):
        if cfg.sign_tres:
            self.page.insert_image(rect, filename=cfg.signature_tres_filename)

    def open(self):
        self.form = pymupdf.open(self.base_filename)
        self.page = self.form[0]

    def add(self):
        self.print_cwd_files()
        image_list, pdf_file_list = self.get_image_file_list()
        image_ratios = self.get_image_ratios(image_list, pdf_file_list)
        all_files = image_list + pdf_file_list
        # Sort: files with None ratio (raw PDFs) go last
        all_files.sort(key=lambda i: (image_ratios[str(i)] is None, image_ratios[str(i)] if image_ratios[str(i)] is not None else 0), reverse=True)
        self.add_images(image_list, pdf_file_list, image_ratios)

    def add_pdf_files(self, pdf_file_list):
        for p in pdf_file_list:
            self.add_file(self.files[p])

    def get_file_index_list(self):
        while True:
            user_input = input("List files to add (ie: 1,3,5,6): ").strip()
            if not user_input:
                return []
            try:
                indices = list(
                    map(lambda i: int(i), filter(lambda x: x.strip() != "", user_input.split(",")))
                )
                return indices
            except ValueError:
                print("Invalid input. Please enter only comma-separated numbers (e.g., 0,1,2). Try again.")

    def get_image_ratios(self, image_list, pdf_file_list):
        def get_ratio(i, is_pdf=False):
            path = os.path.join(self.input_dir, self.files[i])
            if is_pdf:
                doc = pymupdf.open(path)
                page = doc[0]
                rect = page.rect
                aspect = rect.width / rect.height
                # Only add raw if large and aspect ratio is typical for documents
                if rect.width > 500 and rect.height > 700 and 0.7 <= aspect <= 0.8:
                    return None  # Signal to add raw
                return aspect
            else:
                pix = Pixmap(path)
                return pix.width / pix.height
        ratios = {}
        for i in image_list:
            ratios[str(i)] = get_ratio(i)
        for i in pdf_file_list:
            ratios[str(i)] = get_ratio(i, is_pdf=True)
        return ratios

    def add_images(self, image_list, pdf_file_list, image_ratios):
        for i in image_list + pdf_file_list:
            ratio = image_ratios[str(i)]
            if ratio is None:
                self.add_file(self.files[i])  # Add PDF raw
            elif ratio <= 0.333:
                self.add_halfpage_image(self.files[i])
            else:
                self.add_quarterpage_image(self.files[i])

    def print_cwd_files(self):
        input_dir = os.path.join(os.getcwd(), "input")
        if not os.path.exists(input_dir):
            os.makedirs(input_dir)
        files = [f for f in os.listdir(input_dir) if isfile(join(input_dir, f))]
        i = 0
        print("\n './input' directory files")
        for file in files:
            file_extension = os.path.splitext(file)[1]
            if file_extension in Form_Filler.expected_file_extensions:
                print(f"[{i}] {file}")
                i += 1
        self.files = files
        self.input_dir = input_dir

    def get_image_file_list(self):
        image_list = []
        pdf_file_list = []
        for i in self.get_file_index_list():
            filename = self.files[i]
            file_extension = os.path.splitext(filename)[1]
            if file_extension == ".pdf":
                pdf_file_list.append(i)
            else:
                image_list.append(i)
        return (image_list, pdf_file_list)

    def add_file(self, file_name=""):
        if not file_name:
            file_name = input("File name: ")
        file_path = os.path.join(self.input_dir, file_name)
        self.form.insert_file(file_path)

    def add_fullpage_image(self, file_name=""):
        if not file_name:
            file_name = input("File name: ")
        file_path = os.path.join(self.input_dir, file_name)
        self.page = self.form.new_page(-1, *pymupdf.paper_size("letter"))
        self.page.insert_image((0, 0, 595, 842), filename=file_path)
        self.pagefill_count = 0

    def add_halfpage_image(self, file_name=""):
        if not file_name:
            file_name = input("File name: ")
        file_path = os.path.join(self.input_dir, file_name)
        ext = os.path.splitext(file_name)[1]
        if ext == ".pdf":
            # Render first page of PDF to image
            doc = pymupdf.open(file_path)
            pix = doc[0].get_pixmap()
            temp_img = file_path + "_temp.png"
            pix.save(temp_img)
            img_path = temp_img
        else:
            img_path = file_path
        match self.pagefill_count:
            case 1 | 2:
                self.page.insert_image((298, 0, 595, 842), filename=img_path)
                self.pagefill_count = 0
            case _:
                self.page = self.form.new_page(-1, *pymupdf.paper_size("letter"))
                self.page.insert_image((0, 0, 297, 842), filename=img_path)
                self.pagefill_count = 2
        if ext == ".pdf":
            os.remove(temp_img)

    def add_quarterpage_image(self, file_name=""):
        if not file_name:
            file_name = input("File name: ")
        file_path = os.path.join(self.input_dir, file_name)
        ext = os.path.splitext(file_name)[1]
        if ext == ".pdf":
            doc = pymupdf.open(file_path)
            pix = doc[0].get_pixmap()
            temp_img = file_path + "_temp.png"
            pix.save(temp_img)
            img_path = temp_img
        else:
            img_path = file_path
        match self.pagefill_count:
            case 1:
                self.page.insert_image((0, 421, 287, 842), filename=img_path)
            case 2:
                self.page.insert_image((298, 0, 595, 421), filename=img_path)
            case 3:
                self.page.insert_image((298, 421, 595, 842), filename=img_path)
            case _:
                self.page = self.form.new_page(-1, *pymupdf.paper_size("letter"))
                self.page.insert_image((0, 0, 297, 421), filename=img_path)
                self.pagefill_count = 0
        self.pagefill_count += 1
        if ext == ".pdf":
            os.remove(temp_img)

    def open_and_confirm(self, output_path):
        filename = os.path.basename(output_path)
        # Open the file using the default application
        if os.name == "nt":
            os.startfile(output_path)
        elif os.name == "posix":
            os.system(f'xdg-open "{output_path}"')
        else:
            print("Automatic file opening not supported on this OS.")
        # Ask for confirmation
        response = input(f"Please review '{filename}'. Is it correct? ([Y]/n): ").strip().lower()
        if response == "n":
            os.remove(output_path)
            print(f"Deleted '{filename}'. Trying again.")
            self.select()

    def save(self, output_name):
        output_dir = os.path.join(os.getcwd(), "output")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        output_path = os.path.join(output_dir, output_name)
        self.form.save(output_path)
        print(f"Successfully saved {output_name}")
        self.open_and_confirm(output_path)

    def select(self):
        match input("Type de formulaire? [Remboursement|Essence|Dépot|Commandite]: ").lower()[0]:
            case "r":
                self.base_filename = "E24-Remboursement-Achat.pdf"
                self.open()
                self.remboursement()
            case "e":
                self.base_filename = "A22-Formulaire de frais km.pdf"
                self.open()
                self.gas()
            case "d":
                self.base_filename = "A22-Formulaire-dépôt-dargent.pdf"
                self.open()
                self.deposit()
            case "c":
                self.base_filename = "Formulaire de dépôt de commandite.pdf"
                self.open()
                self.commandite()
            case _:
                print("Mauvaise option")
                self.select()


Form_Filler().select()
