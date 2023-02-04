from unittest import main, TestCase

from excel_parsing import ParseSpreadSheet


class TestParseSpreadSheet(TestCase):

    def test_error_messages_for_empty_values_in_sheet(self):
        sheet_parser = ParseSpreadSheet()
        sheet_parser.sheet = 'excel_files/Test File 1.xlsx'
        sheet_parser.get_spread_sheet()
        sheet_parser.parse_spread_sheet()
        self.assertEqual(len(sheet_parser.error_list), 2)

    def test_end_of_file_separator_case(self):
        sheet_parser = ParseSpreadSheet()
        sheet_parser.sheet = 'excel_files/Test File 2.xlsx'
        sheet_parser.get_spread_sheet()
        sheet_parser.parse_spread_sheet()
        self.assertNotIn('Items', sheet_parser.sheet_data)

    def test_excel_file_with_more_data(self):
        sheet_parser = ParseSpreadSheet()
        sheet_parser.sheet = 'excel_files/Test File 3.xlsx'
        sheet_parser.get_spread_sheet()
        sheet_parser.parse_spread_sheet()
        self.assertEqual(len(sheet_parser.sheet_data['Items']), 12)
        self.assertEqual(sheet_parser.sheet_data['Ship From'], 'India')

    def test_excel_file_with_non_required_fields(self):
        sheet_parser = ParseSpreadSheet()
        sheet_parser.sheet = 'excel_files/Test File 3.xlsx'
        sheet_parser.get_spread_sheet()
        sheet_parser.parse_spread_sheet()
        self.assertNotIn('Quantity', sheet_parser.sheet_data['Items'][0])


if __name__ == '__main__':
    main()
