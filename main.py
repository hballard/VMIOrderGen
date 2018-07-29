#!/usr/bin/env python

import argparse
import json
import os

import pandas as pd

CONFIG_FILE = './config.json'
INPUT_COUNTS = './counts.xlsx'
INPUT_BACKORDERS = './backorders.xlsx'
OUTPUT_QUOTE = './output/quote.xlsx'
OUTPUT_OE_UPLOAD = './output/oe_upload.xlsx'


def make_output_dir():
    try:
        os.mkdir('./output')
    except FileExistsError:
        return


def get_args():
    parser = argparse.ArgumentParser(description='Process VMI counts and'
                                     ' return quote and OE Upload files')
    parser.add_argument(
        '--count_file',
        '-C',
        default=INPUT_COUNTS,
        help='Provide a path to a count file to import')
    parser.add_argument(
        '--backorder_file',
        '-B',
        default=INPUT_BACKORDERS,
        help='Provide a path to a backorder file to import')
    parser.add_argument(
        '--quote',
        '-Q',
        dest='quote_path',
        default=OUTPUT_QUOTE,
        help='Provide a file path to ouput Excel quotation file')
    parser.add_argument(
        '--OEUpload',
        '-O',
        dest='OEUpload_path',
        default=OUTPUT_OE_UPLOAD,
        help='Provide a file path for Excel OE upload template file')
    parser.add_argument(
        '--config',
        '-c',
        dest='config_file',
        default=CONFIG_FILE,
        help='Provide a config file if desired in JSON; see'
        ' example; can be used for remapping "shipto" names for example')

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


def write_quote_template(orders, quote_file):
    with pd.ExcelWriter(quote_file, engine='xlsxwriter') as writer:
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
    make_output_dir()
    write_quote_template(orders, args.OEUpload_path)
