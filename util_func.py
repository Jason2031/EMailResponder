import os


def securely_check_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)
