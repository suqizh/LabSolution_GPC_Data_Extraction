import openpyxl

def extract_content(file_path):
    keyword_start = "Dis"
    keyword_end = "[LC"

    start_reading = False
    extracted_content = []

    with open(file_path, 'r') as file:
        for line in file:
            line = line.strip()

            if start_reading and keyword_end in line:
                break

            if start_reading:
                columns = line.split()
                if len(columns) >= 4:
                    extracted_content.append(columns[2:4])

            if keyword_start in line:
                start_reading = True

    return extracted_content


def save_extracted_content_to_excel(file_contents):
    workbook = openpyxl.Workbook()

    for file_name, content in file_contents.items():
        sheet_name = file_name[:-4]  # remove .txt
        sheet = workbook.create_sheet(title=sheet_name)
        for row in content:
            sheet.append(row)

    default_sheet = workbook["Sheet"]
    workbook.remove(default_sheet)

    output_file_path = "output.xlsx"
    workbook.save(output_file_path)
    print(f"Extracted contents has been saved to {output_file_path}")


def process_files():
    file_contents = {}

    file_list = input("Please enter the names of the data files you want to process. Split with comma. ").split(',')

    for file_name in file_list:
        file_name = file_name.strip()
        if not file_name.endswith(".txt"):
            file_name += ".txt"
        file_path = file_name

        
        extracted_content = extract_content(file_path)
        file_contents[file_name] = extracted_content

  
    save_extracted_content_to_excel(file_contents)


process_files()
