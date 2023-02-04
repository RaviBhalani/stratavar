def convert_list_of_dict_to_dict(list_of_dict: list) -> dict:
    return {key: elem[key] for elem in list_of_dict for key in elem}


class DistiQuoteMergeSimuation:

    def __init__(self):
        self.DistiDict = dict()
        self.merged_list = list()

        # Uncomment these lines to run the program. These are commented as input is given through testcases as asked in
        # the interview question.
        # self.BoM = [{1: 5}, {2: 6}, {1: 3}, {2: 1}, {3: 5}, {4: 9}]
        # self.Disti = [{1: 10}, {3: 1}, {5: 5}]

    def merge_part_lists(self) -> None:
        """
        :return:

        This function merges BoM & Disti lines while splitting Disti line, keeping overall quantities same for both
        lines and matching individual part quantities as much as possible.
        """

        # Iterate over BoM line and split the Disti line to match exact or nearest quantity for each BoM part number.
        # Set error flag for those Disti line part numbers which are not present for given BoM part number.
        for part in self.BoM:
            bom_pn, bom_qty = part.popitem()
            if bom_pn in self.DistiDict:
                if bom_qty <= self.DistiDict[bom_pn]:
                    disti_qty = bom_qty
                    self.DistiDict[bom_pn] = self.DistiDict[bom_pn] - disti_qty
                else:
                    disti_qty = self.DistiDict[bom_pn]
                    del self.DistiDict[bom_pn]
                error = False
            else:
                disti_qty, error = None, True

            self.merged_list.append({
                'bom_pn': bom_pn, 'bom_qty': bom_qty, 'disti_pn': bom_pn, 'disti_qty': disti_qty, 'error': error
            })

        # Iterate over remaining Disti line part numbers. These include unique and common ones.
        for part in self.DistiDict:
            self.merged_list.append({
                'bom_pn': part, 'bom_qty': None, 'disti_pn': part, 'disti_qty': self.DistiDict[part], 'error': False
            })

    def print_merged_lines(self) -> None:
        """
        :return:
        This function prints merged lines.
        """
        for line in self.merged_list:
            print(line)


if __name__ == '__main__':
    simulation = DistiQuoteMergeSimuation()

    # As Disti line parts are aggregated, it is safe to convert them to dictionary for easy manipulation.
    simulation.DistiDict = convert_list_of_dict_to_dict(simulation.Disti)

    simulation.merge_part_lists()
    simulation.print_merged_lines()
