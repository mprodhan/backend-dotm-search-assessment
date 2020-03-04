#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "mprodhan/yabamov/madarp"

import os
import argparse
from zipfile import ZipFile

def create_parser():
    parser = argparse.ArgumentParser(description="search for a particular substring within dotm files")
    parser.add_argument("--dir", default=".", help="specify the directory to search into ")
    parser.add_argument('text', help="specify the text that is being searched for ")
    return parser

def scan_directory(dir_name, search_text):
    for root, _, files in os.walk(dir_name):
        for file in files:
            if file.endswith(".dotm"):
                dot_m = ZipFile(os.path.join(root, file))
                content = dot_m.read('word/document.xml').decode('utf-8')
                if search_file(content, search_text):
                    print('Match found in file./dotm_file/' + file)
                    print(search_text)




def search_file(text, search_text):
    for line in text.split('\n'):
        index = line.find(search_text)
        if index >= 0:
            print(line[index-40:index+40])
        return True
    return False


def main():
    parser = create_parser()
    args = parser.parse_args()
    print(args)
    scan_directory(args.dir, args.text)



if __name__ == '__main__':
    main()
