#!/usr/bin/python3
import os
import tqdm
import pandas
import requests
import openpyxl
from pathlib import Path
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By
today = str(date.today())

def get_ligand_webpage():
    os.chdir(today)
    if not os.path.isfile("./ligands"):
        os.system("wget https://www.strem.com/catalog/ligands.php -O ligands")

    ligands_file_lines = open("ligands", "r").readlines()
    # https://stackoverflow.com/questions/64127075/how-to-retrieve-partial-matches-from-a-list-of-strings
    start = ligands_file_lines.index([v for v in ligands_file_lines if "Phosphorus" in v][0]) + 17
    ligands_file_lines = ligands_file_lines[start:]
    end = ligands_file_lines.index([v for v in ligands_file_lines if "</tbody>" in v][0])

    ligands_file_lines = ligands_file_lines[:end]
    ligands_file_lines = ''.join(ligands_file_lines)
      
    return ligands_file_lines

def get_product_pages(source):
    product_links = []
    source = source.split('class="catalog_number"><a href="')
        
    for product in tqdm.tqdm(source, desc="getting product links"):
        product_links.append("https://www.strem.com" + product.split('"')[0])

    return product_links[1:]

def retrieve_product_pages(product_links):
    Path("./products").mkdir(parents=True,exist_ok=True) 

    os.chdir("products")
    browser = webdriver.Edge()
    for link in tqdm.tqdm(product_links, desc="getting product pages"):
        filename = link.split('catalog/v/')[1].replace('/', '_')

        if not Path(filename).is_file():
            browser.get(link)
            browser.find_element(By.CLASS_NAME, 'country_select_button').click()
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

    filepath = os.getcwd() + "\\" + str(date.today()) + "_StremPhosphorusLigands.xlsx"
    
    df = pandas.DataFrame(data)
    df.to_excel(excel_writer=filepath, index=False, header=False)

    wb = openpyxl.load_workbook(filepath)
    ws = wb.worksheets[0]
    for i, line in tqdm.tqdm(enumerate(data), desc="making excel file", total=926):
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
    try:
        Path(today).mkdir(parents=True,exist_ok=False)
    except FileExistsError:
        pass
#        print("Catalog already exists for today. Bye :)")
#        exit(0)

    source = get_ligand_webpage()
    product_links = get_product_pages(source)
    retrieve_product_pages(product_links)
    image_names = get_structure_images(product_links)
    clean_pages()
    data = process_data()
    make_xlsx(data, image_names)

if __name__ == "__main__":
    main()
