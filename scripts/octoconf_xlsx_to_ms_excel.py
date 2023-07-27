# @copyright Copyright (c) 2023 Nicolas GRELLETY
# @license https://opensource.org/licenses/GPL-3.0 GNU GPLv3
# @link https://gitlab.internal.lan/octo-project/octokonverter
# @link https://github.com/nillyr/octokonverter
# @since 0.1.0

import argparse
from itertools import chain
from pathlib import Path
import re
import shutil
import zipfile


def extract_files_from_xlsx(xlsx_file, xlsx_folder):
    with zipfile.ZipFile(xlsx_file, 'r') as zip_ref:
        file_list = zip_ref.namelist()
        for file_name in file_list:
            zip_ref.extract(file_name, path=xlsx_folder)


def create_xlsx_from_folder(xlsx_folder, output_file):
    with zipfile.ZipFile(Path(output_file), 'w') as zip_ref:
        xlsx_folder_path = Path(xlsx_folder)
        xml_or_rels_files = chain(xlsx_folder_path.rglob('*.xml'), xlsx_folder_path.rglob('*.rels'))
        for xml_or_rels_file_path in xml_or_rels_files:
            rel_path = xml_or_rels_file_path.relative_to(xlsx_folder_path)
            zip_ref.write(xml_or_rels_file_path, arcname=rel_path)


def remove_folder(folder_path):
    folder_path = Path(folder_path)

    if folder_path.is_file():
        folder_path.unlink()
    elif folder_path.is_dir():
        shutil.rmtree(folder_path)
    else:
        print(f"[x] Error: '{folder_path}' is neither a file nor a directory.")


def replace_chars(input_str):
    input_str = input_str.replace(';', ',')
    input_str = input_str.replace("'", '&apos;')
    input_str = input_str.replace('"', '&quot;')
    input_str = input_str.replace(', &apos;', ',&apos;')
    input_str = input_str.replace(', &quot;', ',&quot;')
    input_str = input_str.replace('<v></v>', '<v>0</v>')
    return input_str


def format_formulae_for_ms_excel(xlsx_folder):
    regex = r"<f>[a-z0-9\.]+\(\s?('|\")[a-zàâçéèêëîïôûù0-9\s\-\=\(\)]*('|\")![A-Z]+[0-9]*:[A-Z]+[0-9]*;\s?('|\")[a-zàâçéèêëîïôûù0-9\s\-\=\(\)]*\s?('|\");\s?('|\")[a-zàâçéèêëîïôûù0-9\s\-\=\(\)]*('|\")![A-Z]+[0-9]*:[A-Z]+[0-9]*;\s?('|\")[a-zàâçéèêëîïôûù0-9\s\-\=\(\)]*\s?('|\")\)</f><v></v>"

    # sheet2 = synthesis sheet with all the formulae
    with open(f"{xlsx_folder}/xl/worksheets/sheet2.xml", "r") as sheet2:
        og_sheet2_content = sheet2.read()

    start_index = 0
    formatted_sheet2_content = ""
    matches = re.finditer(regex, og_sheet2_content, re.MULTILINE | re.IGNORECASE)
    for _, match in enumerate(matches, start=1):
        formatted_sheet2_content += og_sheet2_content[start_index:match.start()]
        formatted_sheet2_content += replace_chars(match.group())
        start_index = match.end()

    formatted_sheet2_content += og_sheet2_content[start_index:]

    formatted_sheet2_path = Path(f"{xlsx_folder}/xl/worksheets/sheet2.xml")
    formatted_sheet2_path.write_text(formatted_sheet2_content)


def convert(input_file: Path):
    extract_dir = input_file.parent / f"{input_file.stem}_extract"
    output_file = input_file.parent / f"{input_file.stem}-ms-excel.xlsx"
    extract_files_from_xlsx(input_file, extract_dir)
    format_formulae_for_ms_excel(extract_dir)
    create_xlsx_from_folder(extract_dir, output_file)
    remove_folder(extract_dir)
    print(f"[+] non-Excel applications XLSX file converted to Microsoft Excel application: {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", required=True, help="non-Excel appplications XLSX input file")
    args = parser.parse_args()

    convert(Path(args.input))
