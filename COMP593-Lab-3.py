import sys
import os
import pandas as pd
import xlsxwriter
from datetime import date

def create_directory(directory_path):
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)

def create_order_file(data, order_id, order_directory):
    order_file_path = os.path.join(order_directory, f"order_{order_id}.xlsx")
    writer = pd.ExcelWriter(order_file_path, engine='xlsxwriter')
    order_data = data[data['Order ID'] == order_id]
    order_data = order_data.sort_values(by='Item Number')
    order_data['Total Price'] = order_data['Item Quantity'] * order_data['Item Price']
    order_data.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    money_format = workbook.add_format({'num_format': '$#,##0.00'})
    worksheet.set_column('A:A', 11)
    worksheet.set_column('B:B', 13)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 15)
    worksheet.set_column('E:E', 15)
    worksheet.set_column('F:F', 13)
    worksheet.set_coloum('G:G', 13)
    worksheet.set_coloum('G:G', 10)
    worksheet.set_coloum('G:G', 30)
    worksheet.write(len(order_data) + 1, 4, 'Grand Total:', money_format)
    worksheet.write(len(order_data) + 1, 5, order_data['Total Price'].sum(), money_format)
    writer.save()

def main(argv):
    if len(argv) < 2:
        print("Error: Please provide the path of the sales data CSV file.")
        sys.exit(1)

    csv_file_path = argv[1]
    if not os.path.isfile(csv_file_path):
        print("Error: The provided file path does not exist.")
        sys.exit(1)

    today = date.today().isoformat()
    order_directory = os.path.join(os.path.dirname(csv_file_path), f"Orders_{today}")
    create_directory(order_directory)
    data = pd.read_csv(csv_file_path)
    order_ids = data['Order ID'].unique()
    for order_id in order_ids:
        create_order_file(data, order_id, order_directory)

if __name__ == '__main__':
    main(sys.argv)
