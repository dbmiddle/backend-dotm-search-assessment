#!/usr/bin/env python
# -*- coding: utf-8 -*-


"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "dbmiddle"


import os
import zipfile
import sys
import argparse
# import docx


dotm_files_path = os.path.join(os.getcwd(), 'dotm_files')
file_list = os.listdir(dotm_files_path)
file_list = [f for f in file_list if f.endswith('.dotm')]  # store path names of files that match given string


def zip_search():
    file_matches = 0  # keep count of number of files that match given string
    files_searched = 0  # keep count of total number of files searched
    for item in file_list:
        full_path = os.path.join(dotm_files_path, item)
        with zipfile.ZipFile(full_path, 'r') as zip_ref:
            names = zip_ref.namelist()
            for name in names:
                if name.endswith('document.xml'):
                    files_searched += 1
                    with zip_ref.open('word/document.xml') as f:
                        lines = f.readlines()
                        for line in lines:
                            if '$' in line:
                                print('Match found in file:', full_path)
                                dollar_index = line.index('$')
                                print('...'+line[dollar_index - 40: dollar_index + 40]+'...')
                                file_matches += 1
                                break
    print('Total number of files searched: ', files_searched)
    print('File matches: ', file_matches)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--text', help='search text', required=True)
    args = parser.parse_args()
    dollar_search = args.text
    print('The args are:', args)
    if dollar_search:
        zip_search()
    else:
        print('Unknown option')
        sys.exit(1)
    # # raise NotImplementedError("Your awesome code begins here!")


if __name__ == '__main__':
    main()
