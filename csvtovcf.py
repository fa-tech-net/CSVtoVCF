import PySimpleGUI as pSg
import csv
import codecs


# compilation
# pip install auto-py-to-exe
# pyinstaller -y -F -w  "PATH csvToVCF.py"


def build_vcard(contact: dict):
    """
    Apple VCard
    :param contact:
    :return:
    """
    begin = "BEGIN:VCARD\n"
    version = "VERSION:3.0\n"
    name = f"N:{contact['Last Name']};{contact['First Name']};;;\n"
    title = f"FN:{contact['Last Name']} {contact['First Name']}\n"
    org = f"ORG:{contact['Organization']}\n"
    phone = f"TEL;TYPE=CELL;TYPE=VOICE:{contact['Mobile Phone']}\n"
    phone2 = f"TEL;TYPE=WORK;TYPE=VOICE:{contact['Business Phone']}\n"
    phone3 = f"TEL;TYPE=HOME;TYPE=VOICE:{contact['Home Phone']}\n"
    email = f"EMAIL;TYPE=WORK:{contact['E-mail Address']}\n"
    email2 = f"EMAIL;TYPE=WORK:{contact['E-mail 2 Address']}\n"
    email3 = f"EMAIL;TYPE=WORK:{contact['E-mail 3 Address']}\n"
    address = f"item1.ADR;TYPE=WORK:{contact['Business Address']};{contact['Business City']};" \
              f"{contact['Business Postal Code']};{contact['Business Country']}\n"
    address2 = f"item2.ADR;TYPE=HOME:{contact['Home Street']};{contact['Home Address 2']};" \
               f"{contact['Home City']};{contact['Home Postal Code']};{contact['Home Country']}\n"
    end = "END:VCARD\n"

    result = begin + version + name + title + org + phone + \
        phone2 + phone3 + email + email2 + email3 + address + address2 + end
    return result


def parse_csv(filename, delimeter):
    """
    Return Header + rows
    :param filename:
    :param delimeter:
    :return:
    """
    try:
        with codecs.open(f"{filename}", "r", "utf-8-sig") as f:
            content = csv.reader(f, delimiter=delimeter)
            header = next(content)  # saves header
            rows = [dict(zip(header, row)) for row in content]
            return rows
    except IOError as i:
        print(i)
        return []


window_layout = [
    [pSg.Text('Please select the csv file')],
    [pSg.Text('Choose CSV File :'), pSg.InputText('CSV File.csv'), pSg.FileBrowse(file_types=(("CSV", "*.csv"),))],
    [pSg.Text('Save VCF File as :'), pSg.InputText('VCF Card.vcf'), pSg.FileSaveAs(file_types=(("VCARD", ".vcf"),))],
    [pSg.Button("Convert"), pSg.Exit()]
]

default_input = "CSV File.csv"
default_output = "VCF Card.vcf"

main_window = pSg.Window('CSV to VCF Converter').Layout(window_layout)
while True:
    btn, value = main_window.Read()
    print(btn, value)
    if btn in (None, 'Exit'):
        break
    if default_input == value[0] or default_output == value[1]:
        pSg.Popup("Please select input source/ output destination ")
        continue
    input_file = value[0]
    output_file = value[1]
    contacts = parse_csv(input_file, ',')
    output = ""
    for c in contacts:
        output += build_vcard(c)
    with open(output_file, 'w') as outfile:
        outfile.write(output)
    pSg.Popup("Conversion successfully done", "%d contacts exported successfully" % (len(contacts)))

print(value)
