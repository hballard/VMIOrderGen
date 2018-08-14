#!/usr/bin/env python

import json
import os
from argparse import Namespace

import pandas as pd
from gooey import Gooey, GooeyParser

CONFIG_FILE = 'config.json'
DATA_FILE = 'product_data.csv'
INPUT_COUNT_FILE = ''
INPUT_BACKORDER_FILE = ''
OUTPUT_QUOTE_FILE = 'quote.xlsx'
OUTPUT_OEUPLOAD_FILE = 'oe_upload.xlsx'
OUTPUT_PATH = os.path.expanduser('~/Desktop')


@Gooey(program_name='VMI Quote Generator', default_size=(810, 600))
def get_args() -> Namespace:
    parser = GooeyParser(description='Process VMI counts and'
                         ' return quote and OE Upload files')
    parser.add_argument(
        'count_file',
        default=INPUT_COUNT_FILE,
        widget="FileChooser",
        help='Provide a path to a count file to import')
    parser.add_argument(
        'backorder_file',
        default=INPUT_BACKORDER_FILE,
        widget="FileChooser",
        help='Provide a path to a backorder file to import')
    parser.add_argument(
        '--config',
        dest='config_file',
        default=CONFIG_FILE,
        widget="FileChooser",
        help='Provide a config file if desired in JSON format; see'
        ' example')
    parser.add_argument(
        '--product_data',
        dest='product_data_file',
        default=DATA_FILE,
        widget="FileChooser",
        help='Provide a product data file if desired in CSV format;'
        ' will be used for product descriptions and price')
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
    parser.add_argument(
        '--path',
        '-P',
        dest='output_path',
        default=OUTPUT_PATH,
        widget="DirChooser",
        help='Provide a path for the ouput files')

    return parser.parse_args()


def read_config_file(config_file_path: str):
    try:
        with open(config_file_path) as f:
            config = json.load(f)
            return config
    except FileNotFoundError:
        return None
    except json.decoder.JSONDecodeError:
        print("Error in config file; please correct and re-run")
        return None


def make_output_dir(path: str) -> None:
    try:
        os.makedirs(path)
    except FileExistsError:
        return


def process_counts(count_file: str, backorder_file: str,
                   product_data_file: str, config) -> pd.DataFrame:

    # TODO: add try / except and include CSV as an option
    # Read in count file to dataframe
    input_count = pd.read_excel(count_file)

    input_count['bin'], input_count['shipto'], input_count[
        'prod'] = input_count['Barcode'].str.split('-', 2).str

    input_count['prod'] = input_count['prod'].str.rstrip()

    if config:
        input_count.replace(
            to_replace={'shipto': config['shiptos']}, value=None, inplace=True)

    # TODO: add try / except and include CSV as an option
    # Read in backorder file to dataframe
    input_backorder = pd.read_excel(backorder_file)[[
        'prod', 'shipto', 'backorder'
    ]].copy()

    orders = input_count.merge(
        input_backorder, on=['prod', 'shipto'], how='left')
    orders.fillna(0, inplace=True)
    orders['order_amt'] = orders.apply(
        lambda x: x['Count Qty'] - x['backorder'] if x['Count Qty'] >= x['backorder'] else 0,
        axis=1)

    # TODO: add try / except and include Excel as an option
    # Add product description and price
    product_descriptions = pd.read_csv(product_data_file)

    orders_with_descr = orders.merge(
        product_descriptions, on=['prod'], how='left')

    orders_with_descr['price'] = orders_with_descr['price'].replace(
        '[\$,]', '', regex=True).astype(float)

    orders_with_descr['total price'] = (
        orders_with_descr['price'] * orders_with_descr['order_amt'])

    return orders_with_descr


def write_quote_template(orders: pd.DataFrame, quote_file_path: str,
                         config) -> None:
    with pd.ExcelWriter(quote_file_path, engine='xlsxwriter') as writer:
        for i in orders.shipto.unique():
            # Filter orders dataframe by shipto
            orders_by_shipto = orders[orders['shipto'] == i]
            orders_by_shipto.to_excel(
                writer, sheet_name=f'{i}', startrow=1, index=False)

            # Get workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets[i]

            # Specify column widths
            worksheet.set_column('A:A', 28)
            worksheet.set_column('D:D', 12)
            worksheet.set_column('H:H', 19)
            worksheet.set_column('K:K', 50)

            # Specify price column format
            price_format = workbook.add_format()
            price_format.set_num_format(0x08)
            worksheet.set_column('L:L', None, price_format)
            worksheet.set_column('M:M', None, price_format)

            # Set logo on each tab
            worksheet.set_row(0, 43)
            worksheet.insert_image(
                0, 0, os.path.join(os.getcwd(), 'logos', 'PSSI Horz Logo.png'))

            # Add title to worksheet tab
            merge_format = workbook.add_format({
                'bold': True,
                'font_size': 28,
                'font_color': 'red',
                'align': 'center',
                'valign': 'vcenter'
            })

            worksheet.merge_range('F1:J1', 'Quote', merge_format)


def write_oe_template(orders: pd.DataFrame, oe_file_path: str, config) -> None:
    with pd.ExcelWriter(oe_file_path, engine='xlsxwriter') as writer:
        for i in orders.shipto.unique():

            # write orders data
            orders_by_shipto = orders[orders['shipto'] == i]
            orders_by_shipto.to_excel(
                writer,
                sheet_name=f'{i}',
                columns=['prod'],
                header=False,
                index=False,
                startrow=8,
                startcol=0)
            orders_by_shipto.to_excel(
                writer,
                sheet_name=f'{i}',
                columns=['order_amt'],
                header=False,
                index=False,
                startrow=8,
                startcol=2)

            # write oe upload template headers
            worksheet = writer.sheets[i]
            worksheet.write('A1', config.get('customerNo'))
            worksheet.write('A2', config.get('warehouse'))
            worksheet.write('A3', config.get('PO')[i])
            worksheet.write('A4', config.get('shipVia'))
            worksheet.write('B1', i)
            worksheet.write('B2', 'QU')

            # write data headers for product rows
            data_header = ('Product', 'Description', 'Quantity', 'Unit',
                           'Price', 'Discount', 'Disc Type', 'Vendor',
                           'Prod Line', 'Prod Cat', 'Prod Cost', 'Tie Type',
                           'Tie Whse', 'Drop Ship Option', 'Print Option')

            worksheet.write_row('A8', data_header)

            # format the width of several columns
            worksheet.set_column('A:A', 25)
            worksheet.set_column('B:B', 15)
            worksheet.set_column('N:N', 15)
            worksheet.set_column('O:O', 15)


if __name__ == "__main__":
    # Get args and configs
    args = get_args()
    config = read_config_file(args.config_file)

    # Process orders using count and backorders
    orders = process_counts(args.count_file, args.backorder_file,
                            args.product_data_file, config)

    # Write out OE upload template
    make_output_dir(args.output_path)
    oe_file_path = os.path.join(args.output_path, args.OEUpload_name)
    write_oe_template(orders, oe_file_path, config)

    # Write out quote
    quote_file_path = os.path.join(args.output_path, args.quote_name)
    write_quote_template(orders, quote_file_path, config)
