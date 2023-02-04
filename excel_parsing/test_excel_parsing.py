from unittest import main, TestCase

from excel_parsing import ParseSpreadSheet


class TestParseSpreadSheet(TestCase):

    # sheet_parser = ParseSpreadSheet()

    def run_parser(self, sheet_path: str) -> None:
        self.sheet_parser = ParseSpreadSheet()
        self.sheet_parser.sheet = sheet_path
        self.sheet_parser.get_spread_sheet()
        self.sheet_parser.parse_spread_sheet()

    def test_error_messages_for_empty_values_in_sheet(self):
        self.run_parser('excel_files/Test File 1.xlsx')
        self.assertEqual(len(self.sheet_parser.error_list), 2)

    def test_end_of_file_separator_case(self):
        self.run_parser('excel_files/Test File 2.xlsx')
        self.assertNotIn('Items', self.sheet_parser.sheet_data)

    def test_excel_file_with_more_data(self):
        self.run_parser('excel_files/Test File 3.xlsx')
        self.assertEqual(len(self.sheet_parser.sheet_data['Items']), 12)
        self.assertEqual(self.sheet_parser.sheet_data['Ship From'], 'India')

    def test_excel_file_with_non_required_fields(self):
        self.run_parser('excel_files/Test File 3.xlsx')
        self.assertNotIn('Quantity', self.sheet_parser.sheet_data['Items'][0])


if __name__ == '__main__':
    main()
