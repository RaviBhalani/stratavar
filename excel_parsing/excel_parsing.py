from datetime import datetime
from json import dumps
from sys import argv

from xlrd import open_workbook, xldate_as_tuple


def check_for_empty_list(array: list) -> bool:
    for elem in array:
        if elem:
            return True
    else:
        return False


class ParseSpreadSheet:
    def __init__(self):
        self.break_loop = False
        self.file_end = False

        # Sheet variables
        self.sheet_obj = None
        self.sheet = 'excel_files/Python Skill Test.xlsx'
        self.sheet_data = dict()
        self.sheet_no = 0
        self.item_table_header_row = None

        # Sheet data constants variables
        self.date_key = 'Date'
        self.name_key = 'Name'
        self.line_number_key = 'LineNumber'
        self.line_number_key_index = None
        self.file_end_separator = '----------'
        self.header_keys = ['Quote Number', self.date_key, 'Ship To', 'Ship From']
        self.item_table_headers = [self.line_number_key, 'PartNumber', 'Description', 'Price']
        self.item_table_header_indices = dict()

        self.error_list = list()

    def get_spread_sheet(self) -> None:
        """
        :return:
        This function reads a given spreadsheet file and accesses its sheet for the given index.
        """
        self.sheet_obj = open_workbook(self.sheet).sheet_by_index(self.sheet_no)

    def convert_to_json(self) -> str:
        """
        :return: string
        This function returns sheet data dictionary in JSON format.
        """
        return dumps(self.sheet_data)

    def parse_spread_sheet(self):
        """
        :return:
        This function parses the spreadsheet in two parts:
            1. Header Data
            2. Item Table Data - This data is parsed only if items data table is present in the spreadsheet.
        """
        self.parse_header_data()
        if self.file_end is False:
            self.parse_item_data()

    def print_missing_data_info(self):
        """
        :return:
        This function prints all the missing data information collected during parsing of the spreadsheet.
        """
        for error in self.error_list:
            print(self.error_list)

    def parse_header_data(self) -> None:
        for row_no in range(self.sheet_obj.nrows):
            cell_skip_list = []
            row = self.sheet_obj.row_values(row_no)

            for cell_no, cell_value in enumerate(row):

                # Skip iteration if cell is empty or cell is already processed in one of the earlier iterations.
                if not cell_value or cell_no in cell_skip_list:
                    continue

                # Remove leading and trailing whitespaces.
                if type(cell_value) == str:
                    cell_value = cell_value.strip()

                    # Break all loops and stop header data parsing if file end separator is found.
                    if self.file_end_separator in cell_value:
                        self.file_end = True
                        self.break_loop = True
                        break

                # Break loop and stop header data parsing if all headers have been parsed or the iteration has reached
                # the item data table.
                if (not self.header_keys and not self.name_key) or cell_value == self.line_number_key:
                    self.break_loop = True

                    # Store item data table row number for item data parsing.
                    self.item_table_header_row = row_no
                    break

                # Check for required headers only.
                if self.header_keys and cell_value in self.header_keys:

                    # Skip iteration corresponding to header's value.
                    cell_skip_list.append(cell_no + 1)
                    if type(row[cell_no + 1]) == str:
                        self.sheet_data[cell_value] = row[cell_no + 1].strip()
                    self.sheet_data[cell_value] = row[cell_no + 1]

                    # Remove parsed header from header list to facilitate loop break
                    self.header_keys.remove(cell_value)

                # Check for name header as the header and its value is stored in the same cell.
                elif self.name_key in cell_value:
                    self.sheet_data[self.name_key] = cell_value.split(':')[1].strip()

                    # Set name header to None to facilitate loop break.
                    self.name_key = None

            # If all headers have been parsed, then break the loop.
            if self.break_loop:
                break

        if self.date_key in self.sheet_data:
            # Convert float date received from spreadsheet due to xlrd library to yyyy-mm-dd format.
            self.sheet_data[self.date_key] = datetime(*xldate_as_tuple(self.sheet_data[self.date_key], 1)).date().__str__()

    def parse_item_data(self) -> None:
        self.break_loop = False
        self.sheet_data['Items'] = list()

        # Parse item table headers
        for cell_no, cell_value in enumerate(self.sheet_obj.row_values(self.item_table_header_row)):
            cell_value = cell_value.strip()

            # Store LineNumber column index for error logging purpose.
            if cell_value == self.line_number_key:
                self.line_number_key_index = cell_no

            # Parse required headers only.
            if cell_value in self.item_table_headers:
                self.item_table_header_indices[cell_no] = cell_value

        # Start item table data parsing from the row after item table headers.
        for row_no in range(self.item_table_header_row + 1, self.sheet_obj.nrows):
            row = self.sheet_obj.row_values(row_no)
            item_dict = dict()
            row_error_list = list()

            for cell_no, cell_value in enumerate(row):

                # Parse only required data corresponding to required headers.
                if cell_no not in self.item_table_header_indices:
                    continue

                if type(cell_value) == str:
                    cell_value = cell_value.strip()

                    # Break the loop in case parsing reached the end of the file.
                    if self.file_end_separator in cell_value:
                        self.break_loop = True
                        break

                # Collect all missing data information for logging purpose.
                if not cell_value:
                    row_error_list.append(
                        '{} is missing for line number {}'
                        .format(self.item_table_header_indices[cell_no], row[self.line_number_key_index])
                    )

                # Store entire row of item data.
                item_dict[self.item_table_header_indices[cell_no]] = cell_value

            # Break the loop if parsing reached the end of the file.
            if self.break_loop:
                break

            # Check if the entire row is empty. This case handles the case where there are empty rows between items
            # data table & end of file separator. If we have reached end of item data table, then it breaks the loop.
            if check_for_empty_list(list(item_dict.values())):
                self.sheet_data['Items'].append(item_dict)
                self.error_list.extend(row_error_list)
            else:
                break


if __name__ == '__main__':
    sheet_parser = ParseSpreadSheet()
    sheet_parser.get_spread_sheet()
    sheet_parser.parse_spread_sheet()
    print(sheet_parser.convert_to_json())
