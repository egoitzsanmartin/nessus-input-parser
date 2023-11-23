import pandas as pd
from lxml import objectify
import re
import argparse

def parse_document(document):
    result = {}
    current_key = None

    lines = document.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]

        # Ignore empty lines
        if not line.strip():
            i += 1
            continue

        # Check if the line is a key-value pair
        match = re.match(r'\s*([^:]+)\s*:\s*(.*)', line)
        if not match:
            # Check if the line has two spaces before the key
            match = re.match(r'\s{2}([^:]+)\s*:\s*(.*)', line)
            if not match:
                i += 1
                continue

        key = match.group(1).strip()
        value = match.group(2).strip()

        # Check if the key is 'type' to start a new section
        if key.lower() == 'type':
            current_key = value
            if current_key not in result:
                result[current_key] = {}
        elif current_key is not None:
            if key.lower() in ['rationale', 'impact', 'default value', 'additional information', 'modify user parameters for all users with a password set to match']:
                # Capture the values for 'Rationale', 'Impact', 'Default Value', 'Additional information' and 'Modify user parameters for all users with a password set to match'
                i += 2
                value = lines[i].strip()
                result[current_key][key.lower()] = value
            else:
                # Check if the value is multiline and needs to be concatenated
                while i + 1 < len(lines) and not lines[i + 1].strip():
                    i += 1
                    value += ' ' + lines[i].strip()

                # Replace new lines in the value with spaces
                value = value.replace('\n', ' ')

                # Exclude 'type' values from keys
                if key.lower() != 'type':
                    result[current_key][key] = value

        i += 1

    return result

def parse_all_variables(root):
    result = []
    for item in root.custom_item:
        result.append(parse_document(item.text))
    return result

def save_to_excel(data, output_path):
    # Reorganize data for DataFrame
    rows = []
    for item in data:
        row = {'type': list(item.keys())[0]}  # Adding 'type' as a separate column
        row.update(item[row['type']])
        rows.append(row)

    # Create a DataFrame with the reorganized data
    df = pd.DataFrame(rows)

    # Save DataFrame to Excel
    df.to_excel(output_path, index=False)

# Cambia la ruta de salida a una ubicación donde tengas permisos
parser = argparse.ArgumentParser()
parser.add_argument("-f", "--file", help="Nombre (o ruta ) de archivo a parsear")
parser.add_argument("-o", "--output", help="Nombre (o ruta) de archivo excel a generar")
args = parser.parse_args()

if args.file and args.output :
    if args.output.endswith('xlsx') :
        tree = objectify.parse(args.file)
        root = tree.getroot()
        parsed_data = parse_all_variables(root)
        save_to_excel(parsed_data, args.output)
        print("File succesfully created in: "+ args.output)
    else :
        print("err: The extension must be 'xlsx'")
elif args.file:
    tree = objectify.parse(args.file)
    root = tree.getroot()
    parsed_data = parse_all_variables(root)
    save_to_excel(parsed_data, "output.xlsx")
    print("File succesfully created in: output.xlsx")
else :
    print("err: You need to indicate the output file")