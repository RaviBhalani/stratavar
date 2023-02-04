from unittest import main, TestCase
from excel_parsing import ParseSpreadSheet


class TestParseSpreadSheet(TestCase):

    def test_error_messages_for_empty_values_in_sheet(self):
        sheet_parser = ParseSpreadSheet()
        sheet_parser.sheet = 'Test File 1.xlsx'
        sheet_parser.get_spread_sheet()
        sheet_parser.parse_spread_sheet()
        self.assertEqual(len(sheet_parser.error_list), 2)


if __name__ == '__main__':
    main()
