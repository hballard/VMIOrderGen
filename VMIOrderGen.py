#!/usr/bin/env python

import json
import os

from gooey import Gooey, GooeyParser
import pandas as pd

CONFIG_FILE = 'config.json'
INPUT_COUNT_FILE = 'counts.xlsx'
INPUT_BACKORDER_FILE = 'backorders.xlsx'
OUTPUT_QUOTE_FILE = 'quote.xlsx'
OUTPUT_OEUPLOAD_FILE = 'oe_upload.xlsx'
OUTPUT_PATH = 'output'


@Gooey(program_name='VMI Order Generator', default_size=(810, 530))
def get_args():
    parser = GooeyParser(description='Process VMI counts and'
                                     ' return quote and OE Upload files')
    parser.add_argument(
        '--count_file',
        '-C',
        default=INPUT_COUNT_FILE,
        widget="FileChooser",
        help='Provide a path to a count file to import')
    parser.add_argument(
        '--backorder_file',
        '-B',
        default=INPUT_BACKORDER_FILE,
        widget="FileChooser",
        help='Provide a path to a backorder file to import')
    parser.add_argument(
        '--config',
        '-c',
        dest='config_file',
        default=CONFIG_FILE,
        widget="FileChooser",
        help='Provide a config file if desired in JSON format; see'
        #  ' example; can be used for remapping "shipto" names for example')
        ' example')
    parser.add_argument(
        '--path',
        '-P',
        dest='output_path',
        default=OUTPUT_PATH,
        widget="DirChooser",
        help='Provide a path for the ouput files')
    parser.add_argument(
        '--quote',
        '-Q',
        dest='quote_name',
        default=OUTPUT_QUOTE_FILE,
        help='Provide a filename for output of Excel quotation file')
    parser.add_argument(
        '--OEUpload',
        '-O',
        dest='OEUpload_name',
        default=OUTPUT_OEUPLOAD_FILE,
        help='Provide a filename for output of Excel OE upload template file')

    return parser.parse_args()


def read_config_file(config_file):
    try:
        with open(config_file) as f:
            config = json.load(f)
    except FileNotFoundError:
        config = None
    except json.decoder.JSONDecodeError:
        print("Error in config file; please correct and re-run")
        return
    return config


def make_output_dir(path):
    try:
        os.makedirs(path)
    except FileExistsError:
        return


def process_counts(count_file, backorder_file, config):
    input_count = pd.read_excel(count_file)

    input_count['bin'], input_count['shipto'], input_count[
        'prod'] = input_count['barcode'].str.split('-', 2).str

    if config:
        input_count.replace(
            to_replace={'shipto': config['shiptos']}, value=None, inplace=True)

    input_backorder = pd.read_excel(backorder_file)[[
        'prod', 'shipto', 'backorder'
    ]].copy()

    orders = input_count.merge(
        input_backorder, on=['prod', 'shipto'], how='left')

    orders['order_amt'] = orders['qty'] - orders['backorder']
    return orders


def write_quote_template(orders, quote_file_path):
    with pd.ExcelWriter(quote_file_path, engine='xlsxwriter') as writer:
        orders.to_excel(writer)


def write_oe_template(orders, oe_file):
    pass


if __name__ == "__main__":
    # Get args and configs
    args = get_args()
    config = read_config_file(args.config_file)

    # Process orders using count and backorders
    orders = process_counts(args.count_file, args.backorder_file, config)

    # Write out quote and / or OE upload template
    make_output_dir(args.output_path)
    quote_file_path = f'{args.output_path}/{args.OEUpload_name}'
    write_quote_template(orders, quote_file_path)
