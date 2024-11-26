import datetime
import pymupdf
from config import *


def check(page, yes_point, no_point, prompt):
    if input(prompt).lower() == "y":
        page.insert_text(yes_point, "X")
    else:
        page.insert_text(no_point, "X")


def remboursement():
    form = pymupdf.open(purchase_form_filename)
    page = form[0]

    page.insert_text((170, 97), "JDIS")
    page.insert_text((78, 140), "X")
    name = input("Nom pour remboursement: ")
    page.insert_text((235, 183), name)
    page.insert_text((235, 219), user_name)
    page.insert_text((165, 245), user_email)
    page.insert_text((415, 245), user_phone)
    page.insert_text((200, 285), input("Montant à rembourser: ") + "$")
    check(page, (329, 273), (329, 299), "Inscrit Accès D? (Y/N): ")
    check(page, (227, 347), (287, 347), "Pièce justificative originale? (Y/N): ")
    check(page, (227, 389), (287, 389), "Activité de financement? (Y/N): ")
    check(page, (227, 416), (287, 416), "Revenu lié à la dépense? (Y/N): ")
    project = input("Projet/activité si connu: ")
    page.insert_text((210, 455), project)
    while 0 >= page.insert_textbox(
        (210, 470, 540, 550),
        input("Description de la dépense: "),
        lineheight=2.25,
    ):
        print("Despcription too long")
    time = datetime.datetime.now()
    page.insert_text((105, 570), time.strftime("%d-%b-%Y"))
    if sign:
        page.insert_image((355, 533, 400, 583), filename=signature_filename)
    page.insert_text((255, 630), nom_presidence)
    page.insert_text((355, 630), nom_tresorie)
    while True:
        filename = input("Nom du fichier à ajouter? (Laisser vide si aucun): ")
        if not filename:
            break
        form.insert_file(filename)

    form.save(f"remboursement {project} {name} {time.strftime('%Hh%M')}.pdf")


def gas():
    form = pymupdf.open(gas_form_filename)
    page = form[0]

    page.insert_text((170, 97), "JDIS")
    page.insert_text((78, 140), "X")
    name = input("Nom de la personne à rembourser: ")
    page.insert_text((255, 183), name)
    page.insert_text((255, 212), user_name)
    page.insert_text((110, 240), user_email + "@usherbrooke.ca")
    page.insert_text((415, 240), user_phone)
    page.insert_text((200, 280), input("Adresse de destination: "))
    distance = float(input("Distance: "))
    taux_essence = default_gas_rate
    # taux_essence = float(input("Taux essence: "))
    page.insert_text((128, 335), f"{distance}km")
    page.insert_text((370, 337), f"{taux_essence}$")
    page.insert_text((200, 362), f"{distance * taux_essence:.2f}$")
    check(page, (227, 409), (287, 406), "Activité de financement? (Y/N): ")
    check(page, (227, 436), (287, 433), "Revenu lié à la dépense? (Y/N): ")
    project = input("Projet/activité si connu: ")
    page.insert_text((210, 475), project)
    while 0 >= page.insert_textbox(
        (210, 505, 540, 585),
        input("Raison du déplacement: "),
        lineheight=2.25,
    ):
        print("Despcription too long")
    time = datetime.datetime.now()
    page.insert_text((105, 600), time.strftime("%d-%b-%Y"))
    if sign:
        page.insert_image((435, 585, 480, 660), filename=signature_filename)

    form.save(f"remboursement essence {project} {name} {time.strftime('%Hh%M')}.pdf")


def select():
    match input("Type de formulaire? [Remboursement|Essence]:").lower()[0]:
        case "r":
            remboursement()
        case "e":
            gas()
        case _:
            print("Mauvaise option")
            select()


select()
