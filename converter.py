import xlrd
import csv
import os

out_directory = os.path.dirname(__file__) + "/output"  # Define the output directory path
resources_directory = os.path.dirname(__file__) + "/resources"  # Define the source files directory path
template_file = resources_directory + "/Item.csv"  # Define the template file located in /resources
# Define the filename without extension, it will help create output file with the same name
# as input file but with another extension
filename = "/test"  # This will be the name of output file without extension
destination_file = out_directory + ''.join([filename, '.csv'])  # Append file format for output csv file
excel_file = resources_directory + filename + '.xlsx'  # Append file format for input xlsx file

# Create /output directory if it's not already exist
if not os.path.exists(out_directory):
    os.makedirs(out_directory)


# This function will create copy of template file and will append rows from excel file
def excel_to_csv_template(excel_file, template_file, destination_file):
    f_reader = None
    f_writer = None

    try:
        # Read the template and write a copy of data to new file in /resources
        f_reader = open(template_file, 'rb')  # Open template for reading
        # Open new file for writing (the file will be overridden each time when script runs again)
        f_writer = open(destination_file, 'wb')
        for line in f_reader:
            f_writer.write(line)
        f_writer.write('\n')  # The template doesn't have a new_line character at the end of file, so attach it manually

        # Open excel file and get all sheet names
        workbook = xlrd.open_workbook(excel_file)
        all_worksheets = workbook.sheet_names()

        # Reopen file with 'ab' parameter, that will allow to append new data below the existing data
        f_writer = open(destination_file, 'ab')

        # Open destination file for writing csv format
        csv_writer = csv.writer(f_writer, quoting=csv.QUOTE_ALL)
        # Run over all sheets
        for sheet_name in all_worksheets:
            sheet = workbook.sheet_by_name(sheet_name)
            if sheet.nrows == 0:
                continue

            # Run over all entries en each sheet
            for row_num in xrange(sheet.nrows):
                csv_writer.writerow([unicode(entry).encode("utf-8") for entry in sheet.row_values(row_num)])

        print "Template successfully created! Template location: "
        print destination_file

    except IOError as exc:
        print "I/O Error occurred ({0}): {1}".format(exc.errno, exc.strerror)

    finally:
        # If exception will occur the streams will be closed anyway
        if f_reader is not None:
            f_reader.close()
        if f_writer is not None:
            f_writer.close()


excel_to_csv_template(excel_file, template_file, destination_file)
