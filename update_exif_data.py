import os
import json
import datetime
import win32file
import win32api
import re
import pywintypes

DIRECTORY_PATH = r"D:\takeout-20230205T072154Z-001\Takeout\Google Photos\Photos from 2005"
# DIRECTORY_PATH = r"."
TARGET_FOLDER_NAME = "sorted_photos"
DATE_FORMAT = "%Y-%m-%d"

# Supported file extensions
SUPPORTED_EXTENSIONS = [".jpg", ".jpeg", ".png", ".mp4", ".mov"]


def is_supported_file(filename):
    """
    Check if the given filename has a supported file extension.
    """
    file_extension = os.path.splitext(filename)[-1].lower()
    return file_extension in SUPPORTED_EXTENSIONS


def timestamp_to_filetime(timestamp):
    """
    Convert a timestamp to a FILETIME object.
    """
    file_time = datetime.datetime.utcfromtimestamp(timestamp // 10 ** 7 - 11644473600)
    return pywintypes.Time(file_time)


def get_photo_metadata(photo_file_path):
    """
    Given a filepath to a photo, extracts and returns the metadata from the corresponding JSON file.
    Returns None if the JSON file is missing required fields.
    """
    json_file_path = photo_file_path + ".json"
    # print({json_file_path})
    if not os.path.exists(json_file_path):
        return None

    with open(json_file_path, encoding='utf-8') as f:
        metadata = json.load(f)
        # print(metadata)

    required_fields = ["title", "photoTakenTime", "url"] if is_supported_file(photo_file_path) else ["title",
                                                                                                     "videoTakenTime",
                                                                                                     "url"]
    # if not all(field in metadata for field in required_fields):
    #     print(f"Skipping {photo_file_path} due to missing required fields in JSON file.")
    #     return None
    for field in required_fields:
        if field not in metadata:
            print(f"Skipping {photo_file_path} due to missing {field} field in JSON file: {metadata}")
            return None
    return metadata


def get_photo_date(metadata):
    """
    Given photo metadata, extracts and returns the date the photo was taken.
    """
    date_string = metadata["photoTakenTime"]["formatted"] if "photoTakenTime" in metadata else \
        metadata["videoTakenTime"]["formatted"]

    date_string = date_string.encode('utf-8').decode('unicode_escape')  # Decode the string to remove Unicode characters
    date_string = re.sub(r'[^\x00-\x7F]', lambda x: ' ',
                         date_string)  # Replace non-ASCII characters with a regular space
    return datetime.datetime.strptime(date_string, "%b %d, %Y, %I:%M:%S %p UTC").date()


def get_target_folder_path(photo_date):
    """
    Given a date, returns the target folder path for the photo.
    """
    target_folder_name = photo_date.strftime(DATE_FORMAT)
    return os.path.join(DIRECTORY_PATH, TARGET_FOLDER_NAME, target_folder_name)


def move_photo_to_folder(photo_file_path, metadata):
    """
    Moves the photo file and JSON metadata file to the sorted_photos folder and JSON folder, respectively.
    """
    sorted_photos_folder_path = DIRECTORY_PATH
    # sorted_photos_folder_path = os.path.join(DIRECTORY_PATH, TARGET_FOLDER_NAME)
    if not os.path.exists(sorted_photos_folder_path):
        os.makedirs(sorted_photos_folder_path)

    json_folder_path = os.path.join(DIRECTORY_PATH, "JSON")
    if not os.path.exists(json_folder_path):
        os.makedirs(json_folder_path)

    photo_file_name = os.path.basename(photo_file_path)
    target_photo_file_path = os.path.join(sorted_photos_folder_path, photo_file_name)
    target_json_file_path = os.path.join(json_folder_path, photo_file_name + ".json")

    # Move photo file to sorted_photos folder
    os.rename(photo_file_path, target_photo_file_path)

    # Move JSON metadata file to JSON folder
    json_file_path = photo_file_path + ".json"
    os.rename(json_file_path, target_json_file_path)

    # Set the photo's date created attribute to the photo taken date
    photo_date = get_photo_date(metadata)
    date_time = datetime.datetime.combine(photo_date, datetime.time.min).replace(microsecond=0)
    timestamp = int((date_time - datetime.datetime(1970, 1, 1)).total_seconds()) * 10 ** 7 + 116444736000000000

    file_time = timestamp_to_filetime(timestamp)

    h_file = win32file.CreateFile(target_photo_file_path, win32file.GENERIC_WRITE, 0, None, win32file.OPEN_EXISTING, 0,
                                  None)
    win32file.SetFileTime(h_file, file_time, None, None)  # Set the creation time
    h_file.Close()

    # print(f"Processed {photo_file_path} -> sorted_photos and JSON")


def process_photo(photo_file_path):
    """
    Given a filepath to a photo, processes the photo by moving it to the sorted_photos folder and its metadata
    to the JSON folder.
    """
    if os.path.isfile(photo_file_path) and is_supported_file(photo_file_path):
        metadata = get_photo_metadata(photo_file_path)
        if metadata is None:
            print(f"Skipping {photo_file_path} due to missing required fields in JSON file.")
            return
        move_photo_to_folder(photo_file_path, metadata)
        # print(f"Processed {photo_file_path} -> sorted_photos")


def main():
    # Process all photos in directory
    for filename in os.listdir(DIRECTORY_PATH):
        photo_file_path = os.path.join(DIRECTORY_PATH, filename)
        if os.path.isfile(photo_file_path):
            process_photo(photo_file_path)
    print("Files in ", {DIRECTORY_PATH}, " processed")


if __name__ == "__main__":
    main()
