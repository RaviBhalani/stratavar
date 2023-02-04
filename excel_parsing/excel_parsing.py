from datetime import datetime
from json import dumps

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

        self.sheet_obj = None
        self.sheet_data = dict()
        self.sheet = 'Python Skill Test.xlsx'
        self.sheet_no = 0
        self.item_table_header_row = None

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
        self.sheet_obj = open_workbook(self.sheet).sheet_by_index(self.sheet_no)

    def convert_to_json(self) -> str:
        return dumps(self.sheet_data)

    def parse_spread_sheet(self):
        self.parse_header_data()
        if self.file_end is False:
            self.parse_item_data()

    def print_missing_data_info(self):
        for error in self.error_list:
            print(self.error_list)

    def parse_header_data(self) -> None:
        for row_no in range(self.sheet_obj.nrows):
            cell_skip_list = []
            row = self.sheet_obj.row_values(row_no)

            for cell_no, cell_value in enumerate(row):

                if not cell_value or cell_no in cell_skip_list:
                    continue

                if type(cell_value) == str:
                    cell_value = cell_value.strip()
                    if self.file_end_separator in cell_value:
                        self.file_end = True
                        self.break_loop = True
                        break

                if (not self.header_keys and not self.name_key) or cell_value == self.line_number_key:
                    self.break_loop = True
                    self.item_table_header_row = row_no
                    break

                if self.header_keys and cell_value in self.header_keys:
                    cell_skip_list.append(cell_no + 1)
                    if type(row[cell_no + 1]) == str:
                        self.sheet_data[cell_value] = row[cell_no + 1].strip()
                    self.sheet_data[cell_value] = row[cell_no + 1]
                    self.header_keys.remove(cell_value)
                elif self.name_key in cell_value:
                    self.sheet_data[self.name_key] = cell_value.split(':')[1].strip()
                    self.name_key = None

            if self.break_loop:
                break
        self.sheet_data[self.date_key] = datetime(*xldate_as_tuple(self.sheet_data[self.date_key], 1)).date().__str__()

    def parse_item_data(self) -> None:
        self.break_loop = False
        self.sheet_data['Items'] = list()
        for cell_no, cell_value in enumerate(self.sheet_obj.row_values(self.item_table_header_row)):
            if not cell_value:
                continue
            cell_value = cell_value.strip()

            if cell_value == self.line_number_key:
                self.line_number_key_index = cell_no
            if cell_value in self.item_table_headers:
                self.item_table_header_indices[cell_no] = cell_value

        for row_no in range(self.item_table_header_row + 1, self.sheet_obj.nrows):
            row = self.sheet_obj.row_values(row_no)
            item_dict = dict()
            row_error_list = list()

            for cell_no, cell_value in enumerate(row):
                if cell_no not in self.item_table_header_indices:
                    continue

                if type(cell_value) == str:
                    cell_value = cell_value.strip()
                    if self.file_end_separator in cell_value:
                        self.break_loop = True
                        break

                if not cell_value:
                    row_error_list.append(
                        '{} is missing for line number {}'
                        .format(self.item_table_header_indices[cell_no], row[self.line_number_key_index])
                    )
                item_dict[self.item_table_header_indices[cell_no]] = cell_value

            if self.break_loop:
                break

            if check_for_empty_list(list(item_dict.values())):
                self.sheet_data['Items'].append(item_dict)
                self.error_list.extend(row_error_list)


if __name__ == '__main__':
    sheet_parser = ParseSpreadSheet()
    sheet_parser.get_spread_sheet()
    sheet_parser.parse_spread_sheet()
    print(sheet_parser.convert_to_json())
