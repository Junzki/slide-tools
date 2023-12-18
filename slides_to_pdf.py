# -*- coding: utf-8 -*-
import os
import argparse

import win32api
import win32com.client

FILE_TYPE_SLIDE = 1
FILE_TYPE_NORMAL = 2

parser = argparse.ArgumentParser()
parser.add_argument('--copy-all-files', action='store_true', help='Copy all files to target directory')
parser.add_argument('--dest', type=str, help='Destination directory', default=os.getcwd())
parser.add_argument('source', nargs=1, help='Source directory')


def iterate_file(base_dir: str, prefix: str = None):
    if not prefix:
        prefix = base_dir

    collected = list()

    for dirname, dirs, files in os.walk(base_dir):
        for file in files:
            print(dirname, file)
            name_ = os.path.join(dirname, file)
            rel_ = os.path.relpath(name_, prefix)
            if file.endswith('.pptx'):
                collected.append((rel_, FILE_TYPE_SLIDE))
            else:
                collected.append((rel_, FILE_TYPE_NORMAL))

        for d in dirs:
            name_ = os.path.join(dirname, d)
            deep_files = iterate_file(name_, prefix=prefix)

            collected.extend(deep_files)

    return collected


def convert_ppt_to_pdf(ppt_file: str, pdf_file: str):
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Presentations.Open(ppt_file, True, False, False).SaveAs(pdf_file, 32)
    powerpoint.Quit()


def main(source: str, target: str):
    collected = iterate_file(source)
    for name_, type_ in collected:
        src_full_path = os.path.join(source, name_)
        dest_full_path = os.path.join(target, name_)
        dirname_, _ = os.path.split(dest_full_path)

        if type_ == FILE_TYPE_SLIDE:
            os.makedirs(dirname_, exist_ok=True)
            dest_full_path, _ = dest_full_path.rsplit('.', 1)
            dest_full_path += '.pdf'
            print("Converting: ", src_full_path, '-->', dest_full_path)
            convert_ppt_to_pdf(src_full_path, dest_full_path)

        elif copy_all_files:
            os.makedirs(dirname_, exist_ok=True)
            print("Copying: ", src_full_path, '-->', dest_full_path)
            win32api.CopyFile(src_full_path, dest_full_path, False)


if __name__ == '__main__':
    args, _ = parser.parse_known_args()

    src_ = args.source[0]
    dest_ = args.dest
    copy_all_files = args.copy_all_files

    src_ = os.path.abspath(src_)
    dest_ = os.path.abspath(dest_)

    main(src_, dest_)
