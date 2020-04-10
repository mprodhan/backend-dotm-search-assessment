#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "mprodhan/yabamov/madarp/mobcoding/denas"

import os
import sys
import argparse
from zipfile import ZipFile

def create_parser():
    parser = argparse.ArgumentParser(description="search for a particular substring within dotm files")
    parser.add_argument("--dir", default=".", help="specify the directory to search into ")
    parser.add_argument('text', help="specify the text that is being searched for ")
    return parser

files_matched = 0
def scan_directory(dir_name, search_text):
    files_searched = 0
    global files_matched
    for root, _, files in os.walk(dir_name):
        for file in files:
            files_searched += 1
            if file.endswith(".dotm"):
                dot_m = ZipFile(os.path.join(root, file))
                content = dot_m.read('word/document.xml').decode('utf-8')
                if search_file(content, search_text):
                    # print("Searching directory {} for dotm files with text '{}'...".format(search_text))
                    print('Match found in file./dotm_file/' + file)
                    print(search_text)
    print('files matched {}'.format(files_matched))
    print('files searched {}'.format(files_searched))




def search_file(text, search_text):
    global files_matched
    for line in text.split('\n'):
        index = line.find(search_text)
        if index >= 0:
            files_matched += 1
            print(line[index-40:index+40])
            return True
    return False
    return files_matched


def main():
    parser = create_parser()
    args = parser.parse_args()
    print(args)
    scan_directory(args.dir, args.text)



if __name__ == '__main__':
    main()
