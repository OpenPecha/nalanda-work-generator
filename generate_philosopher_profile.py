import csv
from pathlib import Path
from openpecha.buda.api import get_buda_scan_info, get_image_list

def get_bdrc_philo_profile(philo_id):
    philo_profiles = []
    with open(f'./data/bdrc_philo_profiles/{philo_id}.csv', mode ='r')as file:
        philo_profile_rows = csv.reader(file)
        for philo_profile_row in philo_profile_rows:
            philo_profiles.append(philo_profile_row)
    return philo_profiles[1:]

def get_cover_image_link(instance_id, buda_scan_info):
    cover_image_link = ""
    if img_grps := buda_scan_info.get("image_groups", {}):
        first_img_grp = list(img_grps.keys())[0]
        img_list = get_image_list(instance_id, first_img_grp)
        first_img_file_name = img_list[0]['filename']
        cover_image_link = f"https://iiif.bdrc.io/bdr:{first_img_grp}::{first_img_file_name}/full/max/0/default.jpg"
    return cover_image_link

def get_cover_image(instance_id):
    instance_id = instance_id[1:]
    try:
        buda_scan_info = get_buda_scan_info(instance_id)
    except:
        return ""
    if not buda_scan_info:
        return ""
    cover_image_link = get_cover_image_link(instance_id, buda_scan_info)
    cover_image = f'=IMAGE("{cover_image_link}")'
    return cover_image

def parse_philo_instance(philo_instance):
    instance_info = []
    

    work_id = philo_instance[0]
    if "bdr:M" in philo_instance[1]:
        instance_id = philo_instance[1][4:]
        instance = f"[{philo_instance[1]}](https://library.bdrc.io/show/bdr:{instance_id})"
        cover_image = get_cover_image(instance_id)
    else:
        instance = philo_instance[1]
        cover_image = ""
    title = philo_instance[2]
    

    instance_info.append(work_id)
    instance_info.append(instance)
    instance_info.append(title)
    instance_info.append(cover_image)

    return instance_info



def get_philosopher_profile(philo_id, philo_name):
    philo_profile = []
    profile_headers = ["Work","Instance","Title","Cover"]


    bdrc_philo_profile = get_bdrc_philo_profile(philo_id)
    for philo_instance in bdrc_philo_profile:
        philo_profile.append(parse_philo_instance(philo_instance))

    filename = f"./data/philo_profiles/{philo_name}_{philo_id}.xls"
    with open(filename, 'w') as csvfile: 
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(profile_headers) 
        csvwriter.writerows(philo_profile)


if __name__ == "__main__":
    philo_id = "test"
    philo_name = "test_philo"
    get_philosopher_profile(philo_id, philo_name)

