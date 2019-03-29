from typing import Iterable


def separate_cells_set_into_contacted_cells_set(cells_set):
    """

    :param Iterable[Cell] cells_set:
    :return:
    :rtype: list[set[Cell]]
    """

    contact_dict = {}
    for cell in cells_set:
        for other_cell in cells_set:
            if cell == other_cell:
                continue
            elif cell.is_contact_with(other_cell):
                contact_set = contact_dict.setdefault(cell, set())
                contact_set.add(other_cell)

    # travel all the contacted cell to find out all the contact cell set.
    result = []
    un_contact_cells_set = set(cells_set)
    for cell in cells_set:
        if cell in un_contact_cells_set and cell in contact_dict:
            contact_cells_set_of_current_cell = contact_dict[cell]
            un_contact_cells_set -= contact_cells_set_of_current_cell
            cells_needed_to_travel = list(contact_cells_set_of_current_cell)
            cells_has_traveled = {cell}
            new_contact_cells_set = {cell} | contact_cells_set_of_current_cell
            # recursively travel all the contacted cell
            while len(cells_needed_to_travel) > 0:
                cell_traveled = cells_needed_to_travel.pop()
                cells_has_traveled.add(cell_traveled)
                if cell_traveled in contact_dict:
                    contact_cells_set_of_current_cell = contact_dict[cell_traveled]
                    un_contact_cells_set -= contact_cells_set_of_current_cell
                    cells_needed_to_travel.extend(contact_cells_set_of_current_cell - cells_has_traveled)
                    new_contact_cells_set |= contact_cells_set_of_current_cell
            result.append(new_contact_cells_set)
    return result
