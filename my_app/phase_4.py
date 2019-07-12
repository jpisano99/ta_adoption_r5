import os
import json
import datetime
from my_app.settings import app_cfg


def phase_4(run_dir=app_cfg['UPDATES_DIR']):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    archives_dir = app_cfg['ARCHIVES_DIR']

    path_to_run_dir = (os.path.join(home, working_dir, run_dir))
    path_to_main_dir = (os.path.join(home, working_dir))
    path_to_archives = (os.path.join(home, working_dir, archives_dir))

    # bookings_path = os.path.join(path_to_run_dir, app_cfg['XLS_BOOKINGS'])
    # renewals_path = os.path.join(path_to_run_dir, app_cfg['XLS_RENEWALS'])

    # Read the config_dict.json file
    with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE'])) as json_input:
        config_dict = json.load(json_input)
    data_time_stamp = datetime.datetime.strptime(config_dict['data_time_stamp'], '%m-%d-%y')
    last_run_dir = config_dict['last_run_dir']

    str_data_time_stamp = datetime.datetime.strftime(data_time_stamp, '%m-%d-%y')

    # Make an archive directory where we need to place these update files
    path_to_archive_fldr = os.path.join(path_to_archives, str_data_time_stamp + " Updates")
    if os.path.exists(path_to_archive_fldr):
        print (path_to_archive_fldr, ' already exists')
        exit()
    else:
        os.mkdir(os.path.join(path_to_archives, str_data_time_stamp + " Updates"))

    # Move a copy of all new files to the working directory also
    main_files = os.listdir(path_to_run_dir)
    for file in main_files:
        os.rename(os.path.join(path_to_run_dir, file), os.path.join(path_to_archive_fldr, file))

    print(path_to_run_dir)
    print(path_to_main_dir)
    print(path_to_archives)
    print(path_to_archive_fldr)
    exit()

    # Delete all current working files from the working directory stamped with del_date
    files = os.listdir(path_to_main_dir)
    del_date = ''
    for file in files:
        if file.find('Master Bookings') != -1:
            del_date = file[-13:-13 + 8]
            break

    for file in files:
        if file[-13:-13 + 8] == del_date:
            print('Deleting file', file)
            os.remove(os.path.join(path_to_main_dir, file))

    # Move a copy of all new files to the working directory also
    main_files = os.listdir(path_to_run_dir)
    for file in main_files:
        copyfile(os.path.join(path_to_updates, file), os.path.join(path_to_main_dir, file))

    # Move all updates to the archive directory
    update_files = os.listdir(path_to_run_dir)
    for file in update_files:
        print(file)
        os.rename(os.path.join(path_to_updates, file), os.path.join(archive_folder_path, file))

    print('All data files have been refreshed and archived !')
    print('Before init', app_cfg['XLS_BOOKINGS'])
    print('after init', app_cfg['XLS_BOOKINGS'])
    return

phase_4()