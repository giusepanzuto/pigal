import os
import exifread
import datetime
import shutil
import filecmp
import win32file
import pywintypes
import logging
import time

EXIF_DATE_FORMAT = "%Y:%m:%d %H:%M:%S"
EXIF_DATE_TAG = 'EXIF DateTimeOriginal'
EXIF_DATE_TAG_2 = 'Image DateTime'
EXIF_IMAGE_MAKE = 'Image Make'

class FileAlreadyExistException(Exception):
    pass

def try_parse_date(value):
    try:
        return datetime.datetime.strptime(str(value), EXIF_DATE_FORMAT)
    except ValueError:
        return None

def get_date_from_tags(tags):
    if EXIF_DATE_TAG in tags:
        return try_parse_date(tags[EXIF_DATE_TAG])
    if EXIF_DATE_TAG_2 in tags:
        return try_parse_date(tags[EXIF_DATE_TAG_2])

    return None

def get_date_from_os(file_path):
    return datetime.datetime.fromtimestamp(os.path.getmtime(file_path))

def copy(source, dest):
    #print('copy ' + source + ' => ' + dest)
    shutil.copy2(source, dest)
    handle = win32file.CreateFile(dest, win32file.GENERIC_WRITE, 0, None, win32file.OPEN_EXISTING, 0, 0)
    creation_time = pywintypes.Time(os.path.getctime(source))
    win32file.SetFileTime(handle, creation_time)
    handle.close()

def copy_to_library(path_name, file_date, file_maker):
    dest_folder = os.path.join('D:/Img/', str(file_date.year), file_maker)
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
    copy_count = 0
    while True:
        file = os.path.basename(path_name)
        dest_name=file_date.strftime("%Y%m%d-%H%M%S")  
        if(copy_count>0):
            dest_name+='_'+str(copy_count)
        dest_name+= '_(' + os.path.splitext(file)[0] + ')' + os.path.splitext(file)[1]
        dest_path=os.path.join(dest_folder, dest_name)
        if os.path.exists(dest_path):
            if filecmp.cmp(path_name, dest_path, shallow=False):
                raise FileAlreadyExistException(path_name)
            else:
                copy_count+=1
        else:
            break
    copy(path_name, dest_path)

def add_to_library(file_path):
    # if not file.endswith('.jpg'):
    #     continue
    with open(file_path, 'rb') as f:
        tags = exifread.process_file(f, details=False)
    image_date = get_date_from_tags(tags) or get_date_from_os(file_path)
    image_make = 'Unknown'
    if EXIF_IMAGE_MAKE in tags:
        image_make = str(tags[EXIF_IMAGE_MAKE])
    if image_date:
        copy_to_library(file_path, image_date, image_make)
    else:
        raise ValueError('Date not found for file: ' + file_path)

def walk_images():
    count = 0
    duplicated_count = 0
    exception_count=0
    copied_count=0
    for (root, _, files) in os.walk('D:/OneDrive/Immagini', topdown=True):
        files = [file for file in files if file != 'desktop.ini']
        for file in files:
            file_path=os.path.join(root, file)
            try:
                add_to_library(file_path)
                copied_count+=1
            except FileAlreadyExistException as ex:
                duplicated_count+=1
            except Exception as ex:
                exception_count+=1
                logging.exception(ex)
            count+=1
            print(f'\rProcessed: {count}\tCopied: {copied_count}\tDuplicated: {duplicated_count}\tException: {exception_count}', end='\r')

if __name__ == "__main__":
    logging.basicConfig(
        filename=f'execution-{time.strftime("%Y%m%d-%H%M%S")}.log',
        level=logging.ERROR,
        format='%(asctime)s %(levelname)-8s %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    walk_images()