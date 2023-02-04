from unittest import main, TestCase

from disti_quote_merge_simulation import DistiQuoteMergeSimuation
from disti_quote_merge_simulation import convert_list_of_dict_to_dict


class TestDistiQuoteMergeSimuation(TestCase):

    def test_merge_for_given_dataset(self):
        self.simulation = DistiQuoteMergeSimuation()
        self.simulation.BoM = [{'ABC': 2}, {'XYZ': 1}, {'IJK': 1}, {'ABC': 1}, {'IJK': 1}, {'XYZ': 2}, {'DEF': 2}]
        self.simulation.Disti = [{'XYZ': 2}, {'GEF': 2}, {'ABC': 4}, {'IJK': 2}]
        self.simulation.DistiDict = convert_list_of_dict_to_dict(self.simulation.Disti)
        self.simulation.merge_part_lists()
        self.simulation.print_merged_lines()


if __name__ == '__main__':
    main()
