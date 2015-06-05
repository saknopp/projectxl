import os, fnmatch
### this file defines the global variables needed to initialize an instance of the msreader project dictionary
## they should be defined in the program dashboard settings file, not in this placeholder template file

####Used to search for .mpp files to read
FILE_KEYS = ("")

####Path to search for MS Project Files
LOCAL_MPP_PATH = ''
LOCAL_DASH_PATH = ''
NETWORK_MPP_PATH = ''
NETWORK_DASH_PATH = ''
NETWORK_ARCHIVE_PATH = ''
TEMP_PATH = LOCAL_MPP_PATH + 'temp\\'

def project_find(pattern, path):
    result = []
    for root, dirs, files in os.walk(path):
        for name in files:
            if fnmatch.fnmatch(name, pattern):
                result.append(name)
    return result

project_list = []


__all__ = ['FILE_KEYS', 'LOCAL_MPP_PATH', 'LOCAL_DASH_PATH', 'NETWORK_MPP_PATH', 'NETWORK_DASH_PATH', 'NETWORK_ARCHIVE_PATH', \
'TEMP_PATH', 'project_list', 'webdav', 'project_find']

