# -*- coding: utf-8 -*-

import sys
import datetime
import os
import tkinter.filedialog as tk
import re

def openFile(filename):
    infile = open(filename, "r", encoding='utf-8')
    content = infile.readlines()
    infile.close()
    return content


def check_if_date(date_text):
    try:
        datetime.datetime.strptime(date_text, '%d.%m.%y')
        return True
    except ValueError:
        return False


def get_month(month_number):
    month_vector = ["Januar", "Febuar", "Mars", "April", "Mai", "Juni",
                    "Juli", "August", "September", "Oktober", "November", "Desember"]
    month_name = month_vector[month_number-1]
    return month_name


def parseContent(content):
    mnd = {
        "Januar": {"mnd_nr": 1},
        "Febuar": {"mnd_nr": 2},
        "Mars": {"mnd_nr": 3},
        "April": {"mnd_nr": 4},
        "Mai": {"mnd_nr": 5},
        "Juni": {"mnd_nr": 6},
        "Juli": {"mnd_nr": 7},
        "August": {"mnd_nr": 8},
        "September": {"mnd_nr": 9},
        "Oktober": {"mnd_nr": 10},
        "November": {"mnd_nr": 11},
        "Desember": {"mnd_nr": 12}
    }

    content_dict = {"Bank og forsikring": {},
                    "Sparing": {},
                    "Annet forbruk": {},
                    "Bil og transport": {},
                    "Bolig": {},
                    "Ferie og fritid": {},
                    "Dagligvarer": {},
                    "Klær og sko": {},
                    "Helse og velvære": {},
                    "Utdanning": {},
                    "Kafe og restaurant": {},
                    "Innskudd": {}}

    Main_category = ""
    current_month = get_month(datetime.datetime.now().month)
    extract_stats(content)
    year = 0

    for idx in range(len(content)):
        line = content[idx][:-1]
        line = line.replace(",", ".")
        try:
            next_line = content[idx+1][:-1]
        except IndexError:
            break
        if check_if_date(line[0:8]):
            month = line[3:5]
            current_month = get_month(int(month))

        if (line in content_dict.keys()):
            Main_category = line

        if not Main_category in mnd[current_month].keys():
            mnd[current_month][Main_category] = {}
            mnd[current_month][Main_category]["Butikker"] = {}

        if "Velg år" in line:
            year = next_line
            idx += 1

        elif "Velg tidsrom" in line:
            if "Hele året" in next_line:
                current_month = get_month(datetime.datetime.now().month)
            else:
                current_month = next_line
            idx += 1

        else:
            try:
                sum = float(line.replace(" ", ""))
                if "." in line:
                    transaction_type = next_line
                    butikk = content[idx-1][:-1]
                    sub_category = content[idx+2][:-1]
                    if sub_category in mnd[current_month][Main_category].keys():
                        mnd[current_month][Main_category][sub_category] += sum
                    else:
                        mnd[current_month][Main_category][sub_category] = sum

                    if not sub_category in content_dict[Main_category].keys():
                        content_dict[Main_category][sub_category] = 0

                    if not butikk in mnd[current_month][Main_category]["Butikker"].keys():
                        mnd[current_month][Main_category]["Butikker"][butikk] = sum
                    else:
                        mnd[current_month][Main_category]["Butikker"][butikk] += sum

            except ValueError:
                pass
            except KeyError as e:
                pass

    return mnd, content_dict, year

def extract_stats(content):
    return_dict= {}
    category = ""
    sum = 0
    content_string = "".join(content).replace("\n","\n")
    # stats = re.compile(r'''
    #     ^(?P<Transaction_name>(\w*\s*[/&]*){0,5})
    #     (?P<sum>((-?)\d*\s*\d*,\d*))
    #     (?P<method>(\w*\s*){0,5})
    #     (?P<category>(\w*\s*){0,5})$
    # ''')
    # print(content_string)

    stats = re.compile(r"""(?P<Transaction_name>(\w+[/&,]*[^\S\r\n]*){1,5})
(?P<sum>(-?\d*\s*\d+,\d+))
(?P<method>(\w+[/&,]*[^\S\r\n]*){1,5})
(?P<category>(\w+[/&,]*[^\S\r\n]*){1,5})""")
    # print(stats)

    matches = re.finditer(stats,content_string)
    # print(matches)

    for match in matches:
        Transaction_name,sum, method, category = match.group("Transaction_name","sum", "method", "category")
        if not category in return_dict.keys():
            return_dict[category] = {}
        if not Transaction_name in return_dict[category].keys():
            return_dict[category][Transaction_name] = {}
        if not "sum" in return_dict[category][Transaction_name].keys():
            return_dict[category][Transaction_name]["sum"] = float(sum.replace(",",".").replace(" ",""))
        else:
            return_dict[category][Transaction_name]["sum"] += float(sum.replace(",",".").replace(" ",""))
        if not "method" in return_dict[category][Transaction_name].keys():
            return_dict[category][Transaction_name]["method"] = [method]
        else:
            return_dict[category][Transaction_name]["method"].append(method)

    for key in return_dict:
        print(key)
        for subkey in return_dict[key]:
            print("  %50s:%10.2f  "%(subkey, return_dict[key][subkey]["sum"]), return_dict[key][subkey]["method"])
    return return_dict

def write_CSV_file(mnd_content, content_dict, csv_filename):
    outfile_csv = open(csv_filename, "w", encoding='utf-8')
    sum_idx = []

    start_idx = 2
    idx = 1

    outfile_csv.write("Hoved Kategori\tUnderkategori\t")
    for month_key in mnd_content:
        outfile_csv.write(month_key + "\t")
    outfile_csv.write("Total\tSnitt Per Måned\n")

    for cont_key in content_dict:
        for sub_category in content_dict[cont_key]:
            idx += 1
            outfile_csv.write(cont_key+"\t"+sub_category+"\t")
            month_collumn = 98 #Start on collum B since the first month is in collum C
            for month_key in mnd_content:
                month_collumn += 1
                try:
                    outfile_csv.write(
                        str(mnd_content[month_key][cont_key][sub_category]).replace(".", ","))
                except KeyError:
                    outfile_csv.write("0")
                outfile_csv.write("\t")
            outfile_csv.write(
                "=SUM(C{line}:N{line})\t=AVERAGE(C{line}:{month_collumn}{line})\n".format(line=idx,month_collumn=chr(month_collumn).upper()))

        end_idx = idx
        idx += 1
        category_sum_string = "\t"
        for i in range(99, 99+len(mnd_content)+2):
            category_sum_string += "\t=sum({collum}{start}:{collum}{end})".format(
                collum=chr(i), start=start_idx, end=end_idx)

        if content_dict[cont_key]:
            outfile_csv.write(cont_key+category_sum_string)
            outfile_csv.write("\n")
            start_idx = end_idx+2
            sum_idx.append(idx)
        else:
            idx -= 1

    outfile_csv.write("SUM\t\t")
    for i in range(99, 99+len(mnd_content)+2):
        sum_string = ""
        for idx in sum_idx:
            sum_string += "{collum}{line};".format(collum=chr(i), line=idx)
        outfile_csv.write("=sum({sum_string})\t".format(
            sum_string=sum_string[:-1]))

    outfile_csv.close()
    return 0


def write_txt_file(mnd_content, content_dict, txt_filename):

    outfile_txt = open(txt_filename, 'w', encoding="utf-8")

    main_category_padding = find_length_of_longest_word(content_dict)
    sub_category_padding = 0
    for key in content_dict:
        sub_key_length = find_length_of_longest_word(content_dict[key])
        if sub_key_length > sub_category_padding:
            sub_category_padding = sub_key_length

    outfile_txt.write("{name:>{width}}  ".format(
        name="Hovedkategori", width=main_category_padding))
    outfile_txt.write("{name:>{width}}  ".format(
        name="Underkategori", width=sub_category_padding))
    for mnd in mnd_content:
        outfile_txt.write("{:^11} ".format(mnd))
    outfile_txt.write("\n")

    for key in content_dict:
        if content_dict[key]:
            outfile_txt.write("\n")
        for sub_key in content_dict[key]:
            outfile_txt.write("{name:>{width}}  ".format(
                name=key, width=main_category_padding))
            outfile_txt.write("{name:>{width}}: ".format(
                name=sub_key, width=sub_category_padding))

            for mnd in mnd_content:
                try:
                    outfile_txt.write("{:11.2f} ".format(
                        mnd_content[mnd][key][sub_key]))
                except KeyError:
                    outfile_txt.write("{:11.2f} ".format(0.0))

            outfile_txt.write("\n")

    outfile_txt.close()
    return 0


def find_length_of_longest_word(word_list):
    longest_length = 0
    for item in word_list:
        if len(item) > longest_length:
            longest_length = len(item)
    return longest_length


def split_path_and_filename(path_to_filename):
    path, filename = os.path.split(os.path.abspath(path_to_filename))

    return path, filename


def main():
    if len(sys.argv) > 1:
        work_file = sys.argv[1]
    else:
        work_file = tk.askopenfilename()
        print("Opening following file: "+work_file)

    path, filename = split_path_and_filename(work_file)
    content = openFile(work_file)

    mnd_content, content_dict, year = parseContent(content)

    csv_filename = path + "\\" + year + "_" + filename[:-4] + ".csv"
    txt_filename = path + "\\" + year + "_" + filename[:-4] + ".txt"
    write_CSV_file(mnd_content, content_dict, csv_filename)
    write_txt_file(mnd_content, content_dict, txt_filename)


if __name__ == "__main__":
    main()