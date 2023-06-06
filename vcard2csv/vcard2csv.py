#!/usr/bin/env python
import vobject
import glob
import csv
import argparse
import os.path
import sys
import logging
import collections

column_order = [
    'Name',
    'Full name',
    'Cell phone',
    'Work phone',
    'Home phone',
    'Email',
    'Note',
]

def get_phone_numbers(vCard):
    cell = home = work = waid = None
    for tel in vCard.tel_list:
        if vCard.version.value == '2.1':
            if 'CELL' in tel.singletonparams:
                cell = str(tel.value).strip()
            elif 'WORK' in tel.singletonparams:
                work = str(tel.value).strip()
            elif 'HOME' in tel.singletonparams:
                home = str(tel.value).strip()
            if 'WAID' in tel.singletonparams:
                waid = tel.singletonparams['WAID'][0]

            else:
                logging.warning("Warning: Unrecognized phone number category in `{}'".format(vCard))
                tel.prettyPrint()

        elif vCard.version.value == '3.0':
            key = 0
            # print(tel.params)
            if 'TYPE' in tel.params:
                key = 1
                if 'CELL' in tel.params['TYPE']:
                    cell = str(tel.value).strip()
                elif 'WORK' in tel.params['TYPE']:
                    work = str(tel.value).strip()
                elif 'HOME' in tel.params['TYPE']:
                    home = str(tel.value).strip()

            if 'WAID' in tel.params:
                key = 1
                waid = tel.params['WAID'][0]

            if key == 0:
                # print("======")
                # print(cell, work, home, waid)
                # print("======")
                logging.warning("Unrecognized phone number category in `{}'".format(vCard))
                tel.prettyPrint()

        else:
            raise NotImplementedError("Version not implemented: {}".format(vCard.version.value))


    return cell, home, work, waid

def get_info_list(vCard, vcard_filepath):
    vcard = collections.OrderedDict()
    for column in column_order:
        vcard[column] = None
    name = cell = work = home = email = note = None
    vCard.validate()
    for key, val in list(vCard.contents.items()):
        if key == 'fn':
            vcard['Full name'] = vCard.fn.value
        elif key == 'n':
            name = str(vCard.n.valueRepr()).replace('  ', ' ').strip()
            vcard['Name'] = name
        elif key == 'tel':
            cell, home, work, waid = get_phone_numbers(vCard)
            vcard['Cell phone'] = cell
            vcard['Home phone'] = home
            vcard['Work phone'] = work
            if vcard['Cell phone'] is None:
                vcard['Cell phone'] = waid

            elif vcard['Home phone'] is None:
                vcard['Home phone'] = waid

            elif vcard['Work phone'] is None:
                vcard['Work phone'] = waid

        elif key == 'email':
            email = str(vCard.email.value).strip()
            vcard['Email'] = email
        elif key == 'note':
            note = str(vCard.note.value)
            vcard['Note'] = note
        else:
            # An unused key, like `adr`, `title`, `url`, etc.
            pass
    if name is None:
        logging.warning("no name for vCard in file `{}'".format(vcard_filepath))
    if all(telephone_number is None for telephone_number in [cell, work, home]):
        logging.warning("no telephone numbers for file `{}' with name `{}'".format(vcard_filepath, name))

    return vcard

def get_vcards(vcard_filepath):
    with open(vcard_filepath) as fp:
        all_text = fp.read()
    for vCard in vobject.readComponents(all_text):
        yield vCard


def readable_directory(path):
    if not os.path.isdir(path):
        raise argparse.ArgumentTypeError(
            'not an existing directory: {}'.format(path))
    if not os.access(path, os.R_OK):
        raise argparse.ArgumentTypeError(
            'not a readable directory: {}'.format(path))
    return path

def writable_file(path):
    if os.path.exists(path):
        if not os.access(path, os.W_OK):
            raise argparse.ArgumentTypeError(
                'not a writable file: {}'.format(path))
    else:
        # If the file doesn't already exist,
        # the most direct way to tell if it's writable
        # is to try writing to it.
        with open(path, 'w') as fp:
            pass
    return path

def main():
    parser = argparse.ArgumentParser(
        description='Convert a bunch of vCard (.vcf) files to a single TSV file.'
    )
    parser.add_argument(
        'read_dir',
        type=readable_directory,
        help='Directory to read vCard files from.'
    )
    parser.add_argument(
        'tsv_file',
        type=writable_file,
        help='Output file',
    )
    parser.add_argument(
        '-v',
        '--verbose',
        help='More verbose logging',
        dest="loglevel",
        default=logging.WARNING,
        action="store_const",
        const=logging.INFO,
    )
    parser.add_argument(
        '-d',
        '--debug',
        help='Enable debugging logs',
        action="store_const",
        dest="loglevel",
        const=logging.DEBUG,
    )
    args = parser.parse_args()
    logging.basicConfig(level=args.loglevel)

    vcard_pattern = os.path.join(args.read_dir, "*.vcf")
    vcard_paths = sorted(glob.glob(vcard_pattern))
    if len(vcard_paths) == 0:
        logging.error("no files ending with `.vcf` in directory `{}'".format(args.read_dir))
        sys.exit(2)

    # Tab separated values are less annoying than comma-separated values.
    with open(args.tsv_file, 'w', encoding="utf-8", newline='') as tsv_fp:
        writer = csv.writer(tsv_fp, delimiter='\t')
        writer.writerow(column_order)

        for vcard_path in vcard_paths:
            for vcard in get_vcards(vcard_path):
                vcard_info = get_info_list(vcard, vcard_path)
                writer.writerow(list(vcard_info.values()))

if __name__ == "__main__":
    main()
