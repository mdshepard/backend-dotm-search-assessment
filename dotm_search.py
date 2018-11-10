#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text
string within the 'word/document.xml' section of a MSWord .dotm file.
"""
import os
import sys
import glob
import zipfile
import argparse

cwd = os.getcwd() + '/dotm_files/'


def dotm_decode_search(text, folder):
    searched_f = 0
    match_f = 0
    print("Searching CWD: {}".format(folder))
    for filename in glob.iglob(folder + '/*.dotm'):
        zip_file = zipfile.ZipFile(filename, 'r')
        data = zip_file.read('word/document.xml')
        searched_f += 1
        if text in data:
            match_f += 1
            text_index = data.index(text)
            match_display = 'Match found in file ' + filename + '\n'
            match_display += (
                '\t...' + data[text_index - 40:text_index + 40] + '...\n'
                )
            print(match_display)
    print(
        "Files searced:{0} Files matched:{1}".format(searched_f, match_f) +
        '\n')
    return


def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('text', help='text to search within each dotm file')
    parser.add_argument('--dir', help='dir of dotm files to search. def = cwd')
    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()
    text = args.text
    folder = args.dir
    if folder:
        dotm_decode_search(text, folder)
    elif not folder:
        dotm_decode_search(text)
    else:
        print 'unknown option: '
        sys.exit(1)


if __name__ == "__main__":
    main()
