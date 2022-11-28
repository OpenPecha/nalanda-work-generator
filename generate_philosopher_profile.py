import csv
import random
from pathlib import Path
import openpyxl
from openpyxl.styles import Font
from openpecha.buda.api import get_buda_scan_info, get_image_list

def get_bdrc_philo_profile(philo_id):
    philo_profiles = []
    with open(f'./data/bdrc_philo_profiles/{philo_id}.csv', mode ='r')as file:
        philo_profile_rows = csv.reader(file)
        for philo_profile_row in philo_profile_rows:
            philo_profiles.append(philo_profile_row)
    return philo_profiles[1:]

def get_random_img_file_name(img_list):
    img_list = img_list[3:]
    random_img_info = random.choice(img_list)
    random_img_file_name = random_img_info['filename']
    return random_img_file_name


def get_cover_image_link(instance_id, buda_scan_info):
    cover_image_link = ""
    random_image_link = ""
    if img_grps := buda_scan_info.get("image_groups", {}):
        first_img_grp = list(img_grps.keys())[0]
        img_list = get_image_list(instance_id, first_img_grp)
        if len(img_list) > 3:
            first_img_file_name = img_list[2]['filename']
            random_img_file_name = get_random_img_file_name(img_list)
        else:
            return cover_image_link, random_image_link
        cover_image_link = f"https://iiif.bdrc.io/bdr:{first_img_grp}::{first_img_file_name}/full/150,/0/default.jpg"
        random_image_link = f"https://iiif.bdrc.io/bdr:{first_img_grp}::{random_img_file_name}/full/150,/0/default.jpg"
    return cover_image_link, random_image_link

def get_pecha_images(instance_id):
    cover_image = ""
    random_image = ""
    if "_" in instance_id:
        return cover_image, random_image
    instance_id = instance_id[1:]
    try:
        buda_scan_info = get_buda_scan_info(instance_id)
    except:
        return cover_image, random_image
    if not buda_scan_info:
        return cover_image, random_image
    cover_image_link, random_image_link = get_cover_image_link(instance_id, buda_scan_info)
    if cover_image_link:
        cover_image = f"=HYPERLINK(\"https://library.bdrc.io/show/bdr:{instance_id}\",IMAGE(\"{cover_image_link}\"))"
        random_image = f"=HYPERLINK(\"https://library.bdrc.io/show/bdr:{instance_id}\",IMAGE(\"{random_image_link}\"))"
    return cover_image, random_image

def parse_philo_instance(philo_id, philo_instance):
    instance_info = []
    cover_image = ""
    random_image = ""
    

    work_id = philo_instance[0][4:]
    instance_id = philo_instance[1][4:]
    if "bdr:M" in philo_instance[1]:
        instance = f"=HYPERLINK(\"https://library.bdrc.io/show/bdr:{instance_id}?uilang=bo\",\"{instance_id}\")"
        cover_image, random_image = get_pecha_images(instance_id)
    elif "bdr:I" in philo_instance[1]:
        instance = f"=HYPERLINK(\"https://library.bdrc.io/show/bdr:{instance_id}?uilang=bo\",\"{instance_id}\")"

    else:
        instance = philo_instance[1]

    title = philo_instance[2]
    extra_works_link = f"https://library.bdrc.io/search?lg=bo&t=Work&pg=1&f=author,exc,bdr:{philo_id}&uilang=bo&q={title}~1"
    extra_etexts_link = f"https://library.bdrc.io/search?lg=bo&t=Etext&pg=1&f=author,exc,bdr:{philo_id}&uilang=bo&q={title}~1"
    extra_works = f"=HYPERLINK(\"{extra_works_link}\", \"བརྩམས་ཆོས་གཞན།\")"
    extra_etexts = f"=HYPERLINK(\"{extra_etexts_link}\", \"ཡིག་རྐྱང་གཞན།\")"
    
    instance_info.append("")
    instance_info.append(work_id)
    instance_info.append(title)
    instance_info.append(instance)
    instance_info.append(cover_image)
    instance_info.append(random_image)
    instance_info.append(extra_works)
    instance_info.append(extra_etexts)

    return instance_info



def get_philosopher_profile(philo_id, philo_name):
    # philo_profile = []
    profile_headers = ["OCRability info","Work","Title","Instance","Cover Page", "Random Page", "Extra work", "Extra Etext"]
    # philo_profile = [["Work","Instance","Cover","Title"]]
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(profile_headers)
    

    bdrc_philo_profile = get_bdrc_philo_profile(philo_id)
    for instance_walker, philo_instance in enumerate(bdrc_philo_profile, 2):
        philo_instance_info = parse_philo_instance(philo_id, philo_instance)
        sheet.append(philo_instance_info)
        if philo_instance_info[3]:
            sheet.row_dimensions[instance_walker].height = 70
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 30
        sheet.column_dimensions['D'].width = 40
        sheet.column_dimensions['E'].width = 40
        sheet.column_dimensions['F'].width = 35
        sheet.column_dimensions['G'].width = 15
        sheet.column_dimensions['H'].width = 15

        # philo_profile.append(parse_philo_instance(philo_instance))

    # philo_profile_filename = f"./data/philo_profiles/{philo_name}_{philo_id}.csv"
    philo_profile_filename = f"./data/philo_profiles/{philo_id}-{philo_name}.xlsx"
    bold_font = Font(bold=True)
    for cell in sheet["1:1"]:
        cell.font = bold_font
    K = sheet['K2']
    sheet.freeze_panes = K
    workbook.save(philo_profile_filename)
    # with open(filename, 'w') as csvfile: 
    #     csvwriter = csv.writer(csvfile)
    #     csvwriter.writerow(profile_headers) 
    #     csvwriter.writerows(philo_profile)
    # philo_profile_file = pd.read_csv(f"./data/philo_profiles/{philo_name}_{philo_id}.csv")
    # philo_profile_file.to_excel(f"./data/philo_profiles/{philo_name}_{philo_id}.xlsx")
    # return philo_profile


if __name__ == "__main__":
    # philo_id = "test"
    # philo_name = "hd"
    # get_philosopher_profile(philo_id, philo_name)
    philos_info = Path('./data/person_id_mapping.txt').read_text(encoding='utf-8').splitlines()
    for philo_info in philos_info:
        philo_info = philo_info.split(',')
        philo_name = philo_info[0]
        philo_id = philo_info[1]
        get_philosopher_profile(philo_id, philo_name)
        print(f'INFO: {philo_name} completed..')
