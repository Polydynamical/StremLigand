#!/usr/bin/python3
import os
import tqdm
import pandas
import shutil
import requests
import openpyxl
import time
from pathlib import Path
import urllib.parse
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By

today = str(date.today())
#try:
#    Path(today).mkdir(parents=True,exist_ok=False)
#except:
#    shutil.rmtree(today)
#    Path(today).mkdir(parents=True,exist_ok=False)

os.chdir(today)

def get_search_results():

    families = ["Aurolite™", "BIMAH", "BINAP", "BINOL", "Biocatalyst", "BIPHEN", "BIPHEP", "BPE", "Buchwald Precatalysts & Ligands", "cataCXium", "catASium", "CATHy Catalyst", "CatKit", "catMETium", "Corey Catalyst", "DUPHOS", "Escat", "FibreCat™", "Iridicycle", "Jacobsen Ligand", "MARUOKA CAT", "Metal Chloride", "Metallocenes, Derivatives & Cp Precursors", "Metal Oxidation Catalyst", "METAMORPhos", "Metathesis Catalyst", "N-Heterocyclic Carbenes (NHCs)", "Nanomaterials", "NORPHOS", "Organocatalyst", "Palladacycle", "PHANEPHOS", "Photocatalyst", "Photochemical Equipment", "Pincer Ligands and Complexes", "Royer Pd Catalyst", "Schrock's Catalyst", "Schrock-Hoveyda Catalyst", "SEGPHOS", "TADDOL", "Thiourea Catalysts", "ThrePHOX", "UREAPhos"]
    reaction_ids =   [19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 43, 45]
    reaction_names = ["Amination", "Aziridination", "Carbon-carbon bond formation - General", "Carbon-carbon bond Formation - Cross Coupling", "Carbon-carbon bond Formation - Heck Reaction", "Carbon-heteratom bond formation", "Cyclization", "Cyclopropanation", "Decarbonylation", "Decarboxylation", "Expoxidation", "Hydroboration", "Hydroformylation", "Hydrogenation", "Hydrosilyation", "Kinetic Resolution", "Metathesis", "Oxidation", "Hydrolysis", "Isomerization", "Dehydrogenation", "Ring Opening"]


    family_array = []
    for family in families:
        filename = family.replace("&", "and").replace(" ", "_").replace("(","").replace(")","").replace("'", "")
        if not os.path.isfile(filename):
            url = "https://www.strem.com/catalog/family/" + urllib.parse.quote(f"{family}", safe="")
            os.system(f'wget {url} -O {filename}')

        family_file = open(filename, "r").readlines()
        # https://stackoverflow.com/questions/64127075/how-to-retrieve-partial-matches-from-a-list-of-strings
        if family == "Nanomaterials":
            start = family_file.index([v for v in family_file if "nanomaterials_list" in v][0]) + 1
            family_file = family_file[start:]

            end = family_file.index([v for v in family_file if "fix_float" in v][0])
            family_file = family_file[:end]

            family_array.append(''.join(family_file))
            continue

        start = family_file.index([v for v in family_file if "product_section" in v][0]) + 1
        family_file = family_file[start:]

        end = family_file.index([v for v in family_file if "document_section" in v][0])
        family_file = family_file[:end]

        family_array.append(''.join(family_file))


    reaction_array = []
    for family_file_id in reaction_ids:
        filename = reaction_names[reaction_ids.index(family_file_id)].replace(" ", "_")

        if not os.path.isfile(filename):
            url = "https://www.strem.com/catalog/catalysts.php?page_function=search" + urllib.parse.quote(f"&reaction_id={family_file_id}", safe="")
            os.system(f"wget {url} -O {filename}")

        reaction_file = open(filename, "r").readlines()

        # https://stackoverflow.com/questions/64127075/how-to-retrieve-partial-matches-from-a-list-of-strings
        start = reaction_file.index([v for v in reaction_file if "product_section" in v][0]) + 1
        reaction_file = reaction_file[start:]

        end = reaction_file.index([v for v in reaction_file if "document_section" in v][0])
        reaction_file = reaction_file[:end]

        reaction_array.append(''.join(reaction_file))

      
    assert len(families) == len(family_array)
    assert len(reaction_ids) == len(reaction_array)
    return (family_array, reaction_array)

def get_product_links(families, reactions):
    out = open("out.txt", "a")
    families_links  = []
    reactions_links = []

    for family in families:
        links = []
        family = family.split('class="catalog_number"><a href="')
        for page in family:
            links.append("https://www.strem.com" + page.split('"')[0])
            out.write("https://www.strem.com" + page.split('"')[0] + "\n")
        links = links[1:]
        families_links.append(links)

    for reaction in reactions:
        links = []
        reaction = reaction.split('class="catalog_number"><a href="')
        for page in reaction:
            links.append("https://www.strem.com" + page.split('"')[0])
            out.write("https://www.strem.com" + page.split('"')[0] + "\n")
        links = links[1:]
        reactions_links.append(links)


    exit(0)
    return families_links, reactions_links

def retrieve_product_pages(product_links):
    Path("./products").mkdir(parents=True,exist_ok=True) 

    os.chdir("products")
    browser = webdriver.Firefox()
    for i, link in tqdm.tqdm(enumerate(product_links), desc="getting product pages"):
        filename = link.split('catalog/v/')[1].replace('/', '_')

        ## Risky Approach
        # if i==0:
        #     browser.get(link)
        #     browser.find_element(By.CLASS_NAME, 'country_select_button').click()
        #     document_html = browser.page_source
        #     open(filename, "w", encoding="utf-8").write(document_html)
        #     continue

        if not Path(filename).is_file():
            ## Risky approach
            # browser.execute_script(f"window.location.href = '{link}';") 
            # try:
            #     browser.find_element(By.CLASS_NAME, 'country_select_button').click()
            # except:
                # cookies preserved
            #     time.sleep(3)

            browser.get(link)
            document_html = browser.page_source
            open(filename, "w", encoding="utf-8").write(document_html)
    browser.quit()
    
    os.chdir("..")
    return

def get_structure_images(product_links):
    Path("./images").mkdir(parents=True,exist_ok=True) 
    os.chdir("images")

    image_names = []
    for link in tqdm.tqdm(product_links, desc="getting images"):
        filename = link.split('catalog/v/')[1].replace('/', '_') + ".gif"
        image_names.append(filename)

        if not Path(filename).is_file():
            url = "https://www.strem.com/uploads/web_structures/" + link.split('catalog/v/')[1].split("/")[0] + ".gif"
            os.system(f"wget {url} -O {filename}")

    os.chdir("..")
    return image_names
    

def clean_pages():
    os.chdir("products")
    for filename in tqdm.tqdm(os.listdir(), desc="cleaning pages"):
        try:
            out = open(filename, encoding="utf-8").read().split("                    <!-- Email a friend form -->")[0].split("        <div class=\"section body\">")[1]
        except:
            continue
        open(filename, "w", encoding="utf-8").write(out)
    os.chdir("..")



def process_data():
    os.chdir("products")
    unicode_subscript_dict = {
        "<sub>0": "₀",
        "<sub>1": "₁",
        "<sub>2": "₂",
        "<sub>3": "₃",
        "<sub>4": "₄",
        "<sub>5": "₅",
        "<sub>6": "₆",
        "<sub>7": "₇",
        "<sub>8": "₈",
        "<sub>9": "₉"
    }

    data = [["Structure", "Catalog #", "Size Price Availability", "Name", "Color & Form", "Note", "Formula Weight", "Chemical Formula", "Molecular Formula", "CAS #", "MDL #"]]
    for filename in tqdm.tqdm(os.listdir(), desc="processing data"):
        _file = open(filename, encoding="utf-8").read()
        catalog_number = _file.split("<div class=\"catalog_number\">")[1].split("</div><span class=\"category\">")[0]
        category = _file.split("<span class=\"category\">")[1].split("</span>")[0]
        name = _file.split("<span id=\"header_description\">")[1].split("</span>")[0]
        cas_number = _file.split("<td class=\"title top\">CAS Number:</td>")[1].split("</tr>")[0].split("data top\">")[1].split("</td")[0]
        mdl_number = _file.split("<td class=\"title\">MDL Number:</td>")[1].split("</tr>")[0].split("data\">")[1].split("</td")[0]

        molecular_formula = _file.split("<td class=\"title\">Molecular Formula:</td>")[1].split("</tr>")[0].split("data\">")[1].split("</td")[0].replace("</sub>", "")
        for raw, unicode in unicode_subscript_dict.items():
            molecular_formula = molecular_formula.replace(raw, unicode)

        formula_weight = _file.split("<td class=\"title\">Formula Weight:</td>")[1].split("</tr>")[0].split("data\">")[1].split("</td")[0]

        chemical_formula = _file.split("<td class=\"title\">Chemical Formula:</td>")[1].split("</tr>")[0].split("data\">")[1].split("</td")[0].replace("</sub>", "")
        for raw, unicode in unicode_subscript_dict.items():
            chemical_formula = chemical_formula.replace(raw, unicode)

        color_and_form = _file.split("<td class=\"title\">Color and Form:</td>")[1].split("</tr>")[0].split("data\">")[1].split("</td")[0]
        note = _file.split("<td class=\"title\">Note:</td>")[1].split("</tr>")[0].split("data\">")[1].split("</td")[0]

        try:
            note = note.replace("&nbsp;", "")
        except:
            pass

        tmp = _file.split("                <th>Quantity</th>")[1].split("                        </tbody>")[0].split("class=\"size\">")[1:]
        tmp2 = []
        for i, thing in enumerate(tmp):
            tmp2.append(thing.split("<td")[0:3])
            tmp2[i][0] = tmp2[i][0].split("</td>")[0]
            tmp2[i][1] = "$" + tmp2[i][1].split("</span>")[1].split(" ")[0]
            tmp2[i][2] = tmp2[i][2].split("</div>\n")[0].split("\"summary\">")[1]

        tmp3 = []
        for thing in tmp2:
            tmp3.append(' '.join(thing))

        size_price_availability = ""
        for thing in tmp3:
            size_price_availability += thing + " | "
            


        data.append(["",catalog_number, size_price_availability, name, color_and_form, note, formula_weight, chemical_formula, molecular_formula, cas_number, mdl_number])

    os.chdir("..")
    return data

def make_xlsx(data, image_names):
    assert len(data) == len(image_names) + 1

    filepath = os.getcwd() + "\\" + str(date.today()) + "_StremCatalog.xlsx"
    
    df = pandas.DataFrame(data)
    df.to_excel(excel_writer=filepath, index=False, header=False)

    wb = openpyxl.load_workbook(filepath)
    ws = wb.worksheets[0]
    for i, line in tqdm.tqdm(enumerate(data), desc="making excel file", total=len(data)-1):
        try:
            img = openpyxl.drawing.image.Image(os.getcwd() + "\\images\\" + image_names[i])
            ws.add_image(img, f"A{i+2}")
        except:
            ws[f"A{i+2}"] = "NO STRUCTURE IMAGE"

    for i in range(1, 26):
        ws.column_dimensions[chr(ord('@')+i)].width = 55.0 

    for i in range(1, len(data) + 20):
        ws.row_dimensions[i].height = 215.0 

    wb.save(filepath)


def main():
#        print("Catalog already exists for today. Bye :)")
#        exit(0)

    families, reactions = get_search_results()
    product_links = get_product_links(families, reactions)
    retrieve_product_pages(product_links)
    image_names = get_structure_images(product_links)
    clean_pages()
    data = process_data()
    make_xlsx(data, image_names)

if __name__ == "__main__":
    main()
