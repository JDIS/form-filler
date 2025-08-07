import datetime
import pymupdf
from pymupdf import Pixmap
import os
from os.path import isfile, join


class Form_Filler:
    from config import (
        user_name,
        user_email,
        user_phone,
        nom_tresorie,
        nom_presidence,
        default_gas_rate,
        sign,
        signature_filename,
        purchase_form_filename,
        gas_form_filename,
        deposit_form_filename,
    )

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
            and os.path.splitext(f)[1] in [".pdf", ".jpg", ".png"]
        )
    ]

    def check(s, yes_point, no_point, prompt, default="n"):
        target = no_point
        if input(prompt).lower() == "y" or default != "n":
            target = yes_point
        s.page.insert_text(target, "X")

    def remboursement(s):
        s.page.insert_text((170, 97), "JDIS")
        s.page.insert_text((78, 140), "X")
        name = input("Nom pour remboursement: ")
        s.page.insert_text((235, 183), name)
        s.page.insert_text((235, 219), s.user_name)
        s.page.insert_text((165, 245), s.user_email)
        s.page.insert_text((415, 245), s.user_phone)
        s.page.insert_text((200, 285), input("Montant à rembourser: ") + "$")
        s.check((329, 273), (329, 299), "Inscrit Accès D? (Y/[N]): ")
        s.check((227, 347), (287, 347), "Pièce justificative originale? ([Y]/N): ", "y")
        s.check((227, 389), (287, 389), "Activité de financement? (Y/[N]): ")
        s.check((227, 416), (287, 416), "Revenu lié à la dépense? (Y/[N]): ")
        project = input("Projet/activité si connu: ")
        s.page.insert_text((210, 455), project)
        s.textfield((210, 470, 540, 550), "Description de la dépense: ")
        time = datetime.datetime.now()
        s.page.insert_text((105, 570), time.strftime("%d-%b-%Y"))

        s.signForm((355, 533, 400, 583))
        s.page.insert_text((255, 630), s.nom_presidence)
        s.page.insert_text((355, 630), s.nom_tresorie)
        s.add()
        s.save(f"remboursement {project} {name} {time.strftime('%Hh%M')}.pdf")

    def deposit(s):
        s.page.insert_text((175, 97), "JDIS")
        s.page.insert_text((78, 140), "X")
        name = input("Provenance du dépot: ")
        s.page.insert_text((250, 175), name)
        s.page.insert_text((250, 205), s.user_name)
        s.page.insert_text((110, 235), s.user_email + "@usherbrooke.ca")
        s.page.insert_text((420, 235), s.user_phone)
        s.page.insert_text((175, 270), input("Montant déposé: ") + "$")
        method = input(
            "Méthode de dépot? (Argent/Chèque/[Dépot direct]/Square)"
        ).lower()
        method = method[0] if method else "d"
        match method:
            case "a":
                s.page.insert_text((78, 299), "X")
            case "c":
                s.page.insert_text((78, 330), "X")
            case "d":
                s.page.insert_text((225, 299), "X")
            case "s":
                s.page.insert_text((225, 330), "X")
            case _:
                s.page.insert_text((320, 330), "X")
        s.check((227, 373), (287, 373), "Activité de financement? (Y/[N]): ")
        s.check((227, 400), (287, 400), "Commandite? (Y/[N]): ")
        s.textfield((210, 440, 540, 520), "Description de l'activité: ")
        time = datetime.datetime.now()
        s.page.insert_text((105, 570), time.strftime("%d-%b-%Y"))

        s.signForm((425, 575, 470, 625))
        s.page.insert_text((255, 575), s.nom_presidence)
        s.page.insert_text((425, 575), s.nom_tresorie)
        s.save(f"dépot {name} {time.strftime('%Hh%M')}.pdf")

    def gas(s):
        s.page.insert_text((170, 97), "JDIS")
        s.page.insert_text((78, 140), "X")
        name = input("Nom de la personne à rembourser: ")
        s.page.insert_text((255, 183), name)
        s.page.insert_text((255, 212), s.user_name)
        s.page.insert_text((110, 240), s.user_email + "@usherbrooke.ca")
        s.page.insert_text((415, 240), s.user_phone)
        s.page.insert_text((200, 280), input("Adresse de destination: "))
        distance = float(input("Distance: "))

        taux_essence = input("Taux essence [0.20]: ")
        taux_essence = s.default_gas_rate if not taux_essence else float(taux_essence)
        # taux_essence = float(input("Taux essence: "))
        s.page.insert_text((128, 335), f"{distance}km")
        s.page.insert_text((370, 337), f"{taux_essence}$")
        s.page.insert_text((200, 362), f"{distance * taux_essence:.2f}$")
        s.check((227, 409), (287, 406), "Activité de financement? (Y/[N]): ")
        s.check((227, 436), (287, 433), "Revenu lié à la dépense? (Y/[N]): ")
        project = input("Projet/activité si connu: ")
        s.page.insert_text((210, 475), project)
        s.textfield((210, 505, 540, 585), "Raison du déplacement: ")
        time = datetime.datetime.now()
        s.page.insert_text((105, 600), time.strftime("%d-%b-%Y"))
        s.signForm((435, 585, 480, 660))
        s.save(f"essence {project} {name} {time.strftime('%Hh%M')}.pdf")

    def textfield(s, rect, prompt):
        while 0 >= s.page.insert_textbox(rect, input(prompt), lineheight=2.25):
            print("Textfield too long")

    def signForm(s, rect):
        if s.sign:
            s.page.insert_image(rect, filename=s.signature_filename)

    def open(s):
        s.form = pymupdf.open(s.base_filename)
        s.page = s.form[0]

    def add(s):
        s.print_cwd_files()
        image_list, pdf_file_list = s.get_image_file_list()
        image_ratios = s.get_image_ratios(image_list, pdf_file_list)
        all_files = image_list + pdf_file_list
        # Sort: files with None ratio (raw PDFs) go last
        all_files.sort(key=lambda i: (image_ratios[str(i)] is None, image_ratios[str(i)] if image_ratios[str(i)] is not None else 0), reverse=True)
        s.add_images(image_list, pdf_file_list, image_ratios)

    def add_pdf_files(s, pdf_file_list):
        for p in pdf_file_list:
            s.add_file(s.files[p])

    def get_file_index_list(s):
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

    def get_image_ratios(s, image_list, pdf_file_list):
        def get_ratio(i, is_pdf=False):
            path = os.path.join(s.input_dir, s.files[i])
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

    def add_images(s, image_list, pdf_file_list, image_ratios):
        for i in image_list + pdf_file_list:
            ratio = image_ratios[str(i)]
            if ratio is None:
                s.add_file(s.files[i])  # Add PDF raw
            elif ratio <= 0.333:
                s.add_halfpage_image(s.files[i])
            else:
                s.add_quarterpage_image(s.files[i])

    def print_cwd_files(s):
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
        s.files = files
        s.input_dir = input_dir

    def get_image_file_list(s):
        image_list = []
        pdf_file_list = []
        for i in s.get_file_index_list():
            filename = s.files[i]
            file_extension = os.path.splitext(filename)[1]
            if file_extension == ".pdf":
                pdf_file_list.append(i)
            else:
                image_list.append(i)
        return (image_list, pdf_file_list)

    def add_file(s, file_name=""):
        if not file_name:
            file_name = input("File name: ")
        file_path = os.path.join(s.input_dir, file_name)
        s.form.insert_file(file_path)

    def add_fullpage_image(s, file_name=""):
        if not file_name:
            file_name = input("File name: ")
        file_path = os.path.join(s.input_dir, file_name)
        s.page = s.form.new_page(-1, *pymupdf.paper_size("letter"))
        s.page.insert_image((0, 0, 595, 842), filename=file_path)
        s.pagefill_count = 0

    def add_halfpage_image(s, file_name=""):
        if not file_name:
            file_name = input("File name: ")
        file_path = os.path.join(s.input_dir, file_name)
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
        match s.pagefill_count:
            case 1 | 2:
                s.page.insert_image((298, 0, 595, 842), filename=img_path)
                s.pagefill_count = 0
            case _:
                s.page = s.form.new_page(-1, *pymupdf.paper_size("letter"))
                s.page.insert_image((0, 0, 297, 842), filename=img_path)
                s.pagefill_count = 2
        if ext == ".pdf":
            os.remove(temp_img)

    def add_quarterpage_image(s, file_name=""):
        if not file_name:
            file_name = input("File name: ")
        file_path = os.path.join(s.input_dir, file_name)
        ext = os.path.splitext(file_name)[1]
        if ext == ".pdf":
            doc = pymupdf.open(file_path)
            pix = doc[0].get_pixmap()
            temp_img = file_path + "_temp.png"
            pix.save(temp_img)
            img_path = temp_img
        else:
            img_path = file_path
        match s.pagefill_count:
            case 1:
                s.page.insert_image((0, 421, 287, 842), filename=img_path)
            case 2:
                s.page.insert_image((298, 0, 595, 421), filename=img_path)
            case 3:
                s.page.insert_image((298, 421, 595, 842), filename=img_path)
            case _:
                s.page = s.form.new_page(-1, *pymupdf.paper_size("letter"))
                s.page.insert_image((0, 0, 297, 421), filename=img_path)
                s.pagefill_count = 0
        s.pagefill_count += 1
        if ext == ".pdf":
            os.remove(temp_img)

    def open_and_confirm(s, output_path):
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
            s.select()

    def save(s, output_name):
        output_dir = os.path.join(os.getcwd(), "output")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        output_path = os.path.join(output_dir, output_name)
        s.form.save(output_path)
        print(f"Successfully saved {output_name}")
        s.open_and_confirm(output_path)

    def test(s):
        s.form = pymupdf.open(s.purchase_form_filename)
        s.add_fullpage_image("parkscreenshotsg.png")
        s.save("test.pdf")

    def select(s):
        match input("Type de formulaire? [Remboursement|Essence|Dépot]: ").lower()[0]:
            case "r":
                s.base_filename = s.purchase_form_filename
                s.open()
                s.remboursement()
            case "e":
                s.base_filename = s.gas_form_filename
                s.open()
                s.gas()
            case "d":
                s.base_filename = s.deposit_form_filename
                s.open()
                s.deposit()
            case "t":
                s.test()
            case _:
                print("Mauvaise option")
                s.select()


Form_Filler().select()
