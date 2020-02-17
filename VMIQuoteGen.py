#!/usr/bin/env python

# TODO: handle case wehere count is for an invalid barcode format (crashes
# right now)

import json
import os
from argparse import Namespace
from typing import Any, Dict

import pandas as pd
import numpy as np
from gooey import Gooey, GooeyParser

CONFIG_FOLDER = 'config'
CONFIG_FILE = 'config.json'
DATA_FOLDER = 'data'
DATA_FILE = 'product_data.csv'
IMAGE_FOLDER = 'images'
ICON_FOLDER = 'icon'
QUOTE_LOGO_FILE = 'company_logo.png'
INPUT_COUNT_FILE = ''
INPUT_BACKORDER_FILE = ''
OUTPUT_QUOTE_FILE = 'quote'
OUTPUT_OEUPLOAD_FILE = 'oe_upload'
OUTPUT_PATH = os.path.join(os.path.expanduser('~'), 'Desktop')
BASE_PATH = ''

# Create JSON type alias for type hinting config file
JsonType = Dict[str, Any]


@Gooey(
    program_name='VMI Quote Generator',
    image_dir=os.path.join(BASE_PATH, ICON_FOLDER),
    default_size=(810, 600),
)
def get_args() -> Namespace:
    parser = GooeyParser(
        description='Process VMI counts and ' 'create quote and OE Upload files'
    )
    parser.add_argument(
        'count_file',
        default=INPUT_COUNT_FILE,
        widget='FileChooser',
        help='Provide a path to a count file to import (Excel or CSV)',
    )
    parser.add_argument(
        'backorder_file',
        default=INPUT_BACKORDER_FILE,
        widget='FileChooser',
        help='Provide a path to a backorder file to import (Excel or CSV)',
    )
    parser.add_argument(
        '--config',
        dest='config_file',
        default=os.path.join(BASE_PATH, CONFIG_FOLDER, CONFIG_FILE),
        widget='FileChooser',
        help='Provide a config file in JSON format; see example',
    )
    parser.add_argument(
        '--product_data',
        dest='product_data_file',
        default=os.path.join(BASE_PATH, DATA_FOLDER, DATA_FILE),
        widget='FileChooser',
        help='Provide a product data file in CSV or Excel format; see example',
    )
    parser.add_argument(
        '--path',
        '-P',
        dest='output_path',
        default=OUTPUT_PATH,
        widget='DirChooser',
        help='Provide a folder path for the ouput files',
    )
    parser.add_argument(
        '--add_prices',
        '-A',
        dest='add_prices',
        action='store_true',
        widget='CheckBox',
        help='Toggle if you want to add prices to upload file',
    )
    parser.add_argument(
        '--quote',
        '-Q',
        dest='quote_name',
        default=OUTPUT_QUOTE_FILE,
        help='Provide a filename prefix for output of Excel quotation file(s)',
    )
    parser.add_argument(
        '--OEUpload',
        '-O',
        dest='OEUpload_name',
        default=OUTPUT_OEUPLOAD_FILE,
        help='Provide a filename for output of Excel OE upload template file',
    )

    return parser.parse_args()


def read_config_file(config_file_path: str) -> JsonType:
    try:
        with open(config_file_path) as f:
            config = json.load(f)
            return config
    except FileNotFoundError:
        print(
            'You do not have a config file at the location selected',
            'Creating a sample config file for you...',
            'Open it with a text editor and modify the values',
            sep='\n\n',
            end='\n\n',
        )

        config_file_template = {
            'customerNo': '',
            'warehouse': '',
            'shipVia': '',
            'shiptos': {'shipto_alias_1': 'shipto_1', 'shipto_alias_2': 'shipto_2'},
            'PO': {'shipto_1': '0000000', 'shipto_2': '1111111',},
        }

        os.makedirs(os.path.dirname(config_file_path), exist_ok=True)
        with open(config_file_path, 'w') as f:
            json.dump(config_file_template, f, indent=2)
            return config_file_template
    except json.decoder.JSONDecodeError:
        print(
            'Error in config file, please correct and re-run; see exact ' 'cause below',
            end='\n\n',
        )
        raise


def make_output_dir(path: str) -> None:
    try:
        os.makedirs(path)
    except FileExistsError:
        return


def process_counts(
    count_file: str, backorder_file: str, product_data_file: str, config: JsonType
) -> pd.DataFrame:

    # Read in count file to dataframe (default is xlsx, but accepts csv too),
    # format, and add "ship_alias" column
    count_column_names = ['barcode', 'count', 'new_prod', 'additional_qty', 'comments']
    try:
        input_count = pd.read_excel(count_file, names=count_column_names)
    except FileNotFoundError:
        try:
            input_count = pd.read_csv(
                count_file.replace('xlsx', 'csv'), names=count_column_names, header=0
            )
        except FileNotFoundError:
            print('No count file found. Try again with a count file.', end='\n\n')

    input_count['bin'], input_count['shipto'], input_count['prod'] = (
        input_count['barcode'].str.split('-', 2).str
    )

    input_count['prod'] = input_count['prod'].str.rstrip().str.upper()

    input_count['shipto_alias'] = input_count['shipto']

    input_count.replace(
        to_replace={'shipto': config.get('shiptos')}, value=None, inplace=True
    )

    # Read product data file (default is csv, but accepts xlsx too), merge
    # with input_count dataframe
    product_column_names = ['prod', 'description', 'price', 'alt_prod']
    try:
        product_data = pd.read_csv(
            product_data_file, names=product_column_names, header=0
        )
    except FileNotFoundError:
        try:
            product_data = pd.read_excel(
                product_data_file.replace('csv', 'xlsx'), names=product_column_names
            )
        except FileNotFoundError:
            print(
                'You do not have a product data file at the location selected',
                'Creating a sample product data file for you...',
                'Open it with a text editor and modify the values',
                sep='\n\n',
                end='\n\n',
            )
            product_data = pd.DataFrame(columns=product_column_names)
            os.makedirs(os.path.dirname(product_data_file), exist_ok=True)
            product_data.to_csv(product_data_file, index=False)

    product_data['prod'] = product_data['prod'].str.rstrip().str.upper()
    product_data['alt_prod'] = product_data['alt_prod'].str.rstrip().str.upper()
    product_data['description'] = product_data['description'].str.upper()

    input_count = input_count.merge(product_data, on=['prod'], how='left')

    input_count['prod'] = np.where(
        input_count['alt_prod'].isna(), input_count['prod'], input_count['alt_prod']
    )

    # Read in backorder file to dataframe (default is xlsx, but accepts csv
    # too), merge with counts dataframe, fill NAs, and add "order_amt" column
    backorder_column_names = ['enter_date', 'prod', 'backorder', 'custno', 'shipto']
    try:
        input_backorder = pd.read_excel(
            backorder_file,
            usecols='D,F,W,AB,AD',
            names=backorder_column_names,
            skip_rows=[0],
        )
    except FileNotFoundError:
        try:
            input_backorder = pd.read_csv(
                backorder_file.replace('xlsx', 'csv'),
                usecols=[4, 6, 23, 27, 29],
                names=backorder_column_names,
                header=1,
                skip_rows=[0],
            )
        except FileNotFoundError:
            print('No backorder file found. Try again with backorder file.', end='\n\n')

    input_backorder = input_backorder[
        input_backorder.custno.astype(str) == config.get('customerNo')
    ]

    input_backorder.drop_duplicates(inplace=True)

    input_backorder['prod'] = input_backorder['prod'].str.upper()
    input_backorder['shipto'] = input_backorder['shipto'].str.upper()

    input_backorder = (
        input_backorder.groupby(['prod', 'shipto'])['backorder'].sum().reset_index()
    )

    orders = pd.merge(
        input_count.assign(shipto=input_count['shipto'].astype(str)),
        input_backorder.assign(shipto=input_backorder['shipto'].astype(str)),
        on=['prod', 'shipto'],
        how='left',
    )

    orders.fillna(0, inplace=True)

    orders['order_amt'] = orders.apply(
        lambda x: (x['count'] - x['backorder'] if x['count'] >= x['backorder'] else 0),
        axis=1,
    )

    orders['order_amt'] = orders['order_amt'] + orders['additional_qty']

    orders['price'] = orders['price'].replace('[\$,]', '', regex=True).astype(float)

    orders['total_price'] = orders['price'] * orders['order_amt']

    # Re-arrange column order
    new_column_order = [
        'barcode',
        'count',
        'new_prod',
        'additional_qty',
        'comments',
        'bin',
        'shipto',
        'shipto_alias',
        'prod',
        'description',
        'backorder',
        'order_amt',
        'price',
        'total_price',
    ]
    orders = orders[new_column_order]

    return orders


def write_quote_template(orders: pd.DataFrame, quote_file_path: str) -> None:
    image_path = os.path.join(BASE_PATH, IMAGE_FOLDER, QUOTE_LOGO_FILE)
    if not os.path.exists(image_path):
        print(
            'You are missing a "company_logo.png" file in the "images"'
            ' folder; your quotes will not have an image in the'
            ' header unless you provide one',
            end='\n\n',
        )

    for shipto_alias in orders.shipto_alias.unique():
        with pd.ExcelWriter(
            f'{quote_file_path}-{shipto_alias}.xlsx', engine='xlsxwriter'
        ) as writer:

            # Filter orders dataframe by shipto_alias and print data
            orders_by_shipto = orders[orders['shipto_alias'] == shipto_alias]
            print_columns = [
                'bin',
                'shipto',
                'shipto_alias',
                'prod',
                'description',
                'new_prod',
                'count',
                'additional_qty',
                'backorder',
                'order_amt',
                'price',
                'total_price',
            ]
            print_headers = [
                'Bin',
                'Shipto',
                'Shipto Alias',
                'Prod',
                'Description',
                'New Prod',
                'Count',
                'Additional Qty',
                'Backorder',
                'Order Amt',
                'Unit Price',
                'Total Price',
            ]
            orders_by_shipto.to_excel(
                writer,
                sheet_name=f'{shipto_alias}',
                columns=print_columns,
                header=print_headers,
                startrow=1,
                index=False,
            )

            # Get workbook and worksheet objects; format worksheet print
            # properties
            workbook = writer.book
            workbook.set_size(1800, 1200)
            worksheet = writer.sheets[shipto_alias]
            worksheet.set_landscape()
            worksheet.hide_gridlines(1)
            worksheet.fit_to_pages(1, 0)

            # Specify column widths
            worksheet.set_column('C:C', 12)
            worksheet.set_column('D:D', 20)
            worksheet.set_column('E:E', 48)
            worksheet.set_column('H:H', 13)
            worksheet.set_column('K:K', 12)
            worksheet.set_column('L:L', 11)

            # Specify price column format
            price_format = workbook.add_format()
            price_format.set_num_format(0x08)
            worksheet.set_column('M:M', 13, price_format)
            worksheet.set_column('N:N', 11, price_format)

            # Set logo on each tab
            worksheet.set_row(0, 45)
            # TODO: also add optional command arguement to allow user to
            # overide
            if os.path.exists(image_path):
                worksheet.insert_image(0, 0, image_path)

            # Add title to worksheet tab
            merge_format = workbook.add_format(
                {
                    'bold': True,
                    'font_size': 28,
                    'font_color': 'red',
                    'align': 'center',
                    'valign': 'vcenter',
                }
            )

            total_price_format = workbook.add_format(
                {'bold': True, 'font_size': 14, 'font_color': 'red',}
            )
            total_price_format.set_num_format(0x08)

            # add Total Price Excel "sum" function to header
            worksheet.merge_range('E1:I1', 'Quote', merge_format)
            worksheet.write('K1', 'Total Price', total_price_format)
            worksheet.write_formula(
                'L1', f'=sum(L3:L{2+len(orders_by_shipto.index)})', total_price_format
            )


def write_oe_template(
    orders: pd.DataFrame, oe_file_path: str, add_prices: bool, config: JsonType
) -> None:

    with pd.ExcelWriter(f'{oe_file_path}.xlsx', engine='xlsxwriter') as writer:
        for shipto in orders.shipto.unique():

            # write orders data
            orders_by_shipto = orders[orders['shipto'] == shipto]
            orders_by_shipto.to_excel(
                writer,
                sheet_name=f'{shipto}',
                columns=['prod'],
                header=False,
                index=False,
                startrow=8,
                startcol=0,
            )
            orders_by_shipto.to_excel(
                writer,
                sheet_name=f'{shipto}',
                columns=['order_amt'],
                header=False,
                index=False,
                startrow=8,
                startcol=2,
            )
            if add_prices:
                orders_by_shipto.to_excel(
                    writer,
                    sheet_name=f'{shipto}',
                    columns=['price'],
                    header=False,
                    index=False,
                    startrow=8,
                    startcol=4,
                )

            # write oe upload template headers
            worksheet = writer.sheets[shipto]
            worksheet.write('A1', config.get('customerNo'))
            worksheet.write('A2', config.get('warehouse'))
            worksheet.write('A3', config.get('PO', {}).get(shipto))
            worksheet.write('A4', config.get('shipVia'))
            worksheet.write('B1', shipto)
            worksheet.write('B2', 'QU')

            # write data headers for product rows
            data_header = (
                'Product',
                'Description',
                'Quantity',
                'Unit',
                'Price',
                'Discount',
                'Disc Type',
                'Vendor',
                'Prod Line',
                'Prod Cat',
                'Prod Cost',
                'Tie Type',
                'Tie Whse',
                'Drop Ship Option',
                'Print Option',
            )

            worksheet.write_row('A8', data_header)

            # format the width of several columns
            worksheet.set_column('A:A', 25)
            worksheet.set_column('B:B', 15)
            worksheet.set_column('N:N', 15)
            worksheet.set_column('O:O', 15)


if __name__ == '__main__':

    # Get args and configs
    args = get_args()
    config = read_config_file(args.config_file)

    # Process orders using VMI count and current SXe backorders
    orders = process_counts(
        args.count_file, args.backorder_file, args.product_data_file, config
    )

    # Write out OE upload template file (one tab for each shipto)
    make_output_dir(args.output_path)
    oe_file_path = os.path.join(args.output_path, args.OEUpload_name)
    write_oe_template(orders, oe_file_path, args.add_prices, config)

    # Write out quote files (one file for each shipto)
    quote_file_path = os.path.join(args.output_path, args.quote_name)
    write_quote_template(orders, quote_file_path)
