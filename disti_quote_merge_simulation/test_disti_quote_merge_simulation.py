from unittest import main, TestCase

from disti_quote_merge_simulation import DistiQuoteMergeSimuation
from disti_quote_merge_simulation import convert_list_of_dict_to_dict


class TestDistiQuoteMergeSimuation(TestCase):

    def run_merge(self, bom_line, disti_line):
        self.simulation = DistiQuoteMergeSimuation()
        self.simulation.BoM, self.simulation.Disti = bom_line, disti_line
        self.simulation.DistiDict = convert_list_of_dict_to_dict(self.simulation.Disti)
        self.simulation.merge_part_lists()
        self.simulation.print_merged_lines()

    def test_merge_for_different_datasets(self):
        bom_line = [{1: 5}, {2: 6}, {1: 3}, {2: 1}, {3: 5}, {4: 9}]
        disti_line = [{1: 10}, {3: 1}, {5: 5}]
        self.run_merge(bom_line, disti_line)

        print('----------------------------------------------------------------')

        bom_line = [{1: 10}, {3: 1}, {5: 5}]
        disti_line = [{1: 5}, {2: 6}, {1: 3}, {2: 1}, {3: 5}, {4: 9}]
        self.run_merge(bom_line, disti_line)


if __name__ == '__main__':
    main()
