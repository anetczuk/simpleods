#
# Copyright (c) 2026, Arkadiusz Netczuk <dev.arnet@gmail.com>
# All rights reserved.
#
# This source code is licensed under the BSD 3-Clause license found in the
# LICENSE file in the root directory of this source tree.
#

import logging

import copy
import tempfile

import odf
from odf.opendocument import load as odf_load

# from odf.namespaces import TABLENS

from odf.table import Table, TableColumn
from odf.table import TableRow
from odf.table import TableCell
from odf.text import P

_LOGGER = logging.getLogger(__name__)


## =======================================================


def get_spreadsheet_coords(row_index: int, cell_index: int) -> tuple[str, str]:
    ## alphabet size: 26 letters
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    col_high = int(cell_index / 26)
    col_low = int(cell_index % 26)
    col_name = ""
    if col_high > 0:
        col_name = alphabet[col_high - 1]
    low_letter = alphabet[col_low]
    col_name = col_name + low_letter
    return (f"{row_index + 1}", col_name)


## =======================================================


class Cell:

    def __init__(self, cell: TableCell):
        self.cell: TableCell = cell

    def get_parent(self) -> "Row":
        row = self.cell.parentNode
        return Row(row)

    def clear_cell(self):
        # Remove all child nodes (usually <text:p>)
        for child in list(self.cell.childNodes):  # snapshot list to avoid iteration issues
            self.cell.removeChild(child)

        # remove attributes
        for attr in ["value", "valuetype", "formula", "stylename"]:
            if self.cell.getAttribute(attr):
                self.cell.removeAttribute(attr)

    def copy_cell(self) -> "Cell":
        ## deep copy of cell object is not advised
        ## return copy.deepcopy(self.cell)

        dst_cell = TableCell()
        target_cell = Cell(dst_cell)
        self.copy_cell_to(target_cell)
        return target_cell

    def copy_cell_to(self, target_cell: "Cell"):
        attr_dict = {
            "style-name": "stylename",
            "value-type": "valuetype",
        }

        # copy attributes
        dst_cell = target_cell.cell
        for (_attrns, attr), value in self.cell.attributes.items():
            attr_value = attr_dict.get(attr, attr)
            dst_cell.setAttribute(attr_value, value)

        # copy content
        for child in self.cell.childNodes:
            dst_cell.addElement(copy.deepcopy(child))

    def move_content_to(self, target_cell: "Cell"):
        self.copy_cell_to(target_cell)
        self.clear_cell()

    def get_repeat(self) -> int:
        return Cell.get_repeat_attribute(self.cell)

    def is_repeated(self) -> bool:
        return Cell.get_repeat_attribute(self.cell) > 1

    @staticmethod
    def get_repeat_attribute(cell) -> int:
        return int(cell.getAttribute("numbercolumnsrepeated") or 1)

    def get_index(self) -> list[int]:
        row = self.cell.parentNode
        ret_index = []
        index = 0
        cells_list = row.getElementsByType(TableCell)
        for curr_item in cells_list:
            repeat_num = Cell.get_repeat_attribute(curr_item)
            if curr_item != self.cell:
                index += repeat_num
                continue
            for _i in range(repeat_num):
                ret_index.append(index)
                index += 1
        return ret_index

    def get_spreadsheet_coords(self) -> tuple[str, str]:
        row: Row = self.get_parent()
        row_index = row.get_index()
        cell_index = self.get_index()
        return get_spreadsheet_coords(row_index[0], cell_index[0])

    def get_text(self) -> str:
        text = ""
        for p in self.cell.getElementsByType(P):
            if p.firstChild:
                text += str(p.firstChild.data)
        return text

    def set_text(self, content: str):
        for child in list(self.cell.childNodes):  # snapshot list to avoid iteration issues
            self.cell.removeChild(child)
        self.cell.addElement(P(text=content))

    def expand_cell(self):
        repeat = Cell.get_repeat_attribute(self.cell)
        if repeat < 2:
            return
        self.cell.removeAttribute("numbercolumnsrepeated")

        new_items_list = []
        for _i in range(repeat - 1):
            new_cell: Cell = self.copy_cell()
            new_items_list.append(new_cell)

        parent_row = self.get_parent()
        parent_row.add_cells_before(new_items_list, self.cell)


class Row:

    def __init__(self, row: TableRow):
        self.row: TableRow = row

    def get_parent(self) -> "Sheet":
        sheet = self.row.parentNode
        return Sheet(sheet)

    def count_cells(self) -> int:
        cell_list = self.row.getElementsByType(TableCell)
        return len(cell_list)

    def add_new_column(self, content=None):
        sheet: Sheet = self.get_parent()
        sheet.add_new_column()
        cell_list = self.row.getElementsByType(TableCell)
        item = cell_list[-1]
        self.add_cell(item)
        cell = Cell(item)
        if content:
            cell.set_text(content)
        return cell

    def add_new_cell(self):
        item = TableCell()
        self.row.addElement(item)

    def add_cell(self, item: TableCell):
        self.row.addElement(item)

    def add_cells(self, item_list: list[TableCell]):
        for item in item_list:
            self.row.addElement(item)

    def add_cells_before(self, item_list: list[Cell], reference_cell):
        for item in item_list:
            self.row.insertBefore(item.cell, reference_cell)

    def get_repeat(self) -> int:
        return Row.get_repeat_attribute(self.row)

    @staticmethod
    def get_repeat_attribute(row) -> int:
        return int(row.getAttribute("numberrowsrepeated") or 1)

    def is_repeated(self) -> bool:
        return Row.get_repeat_attribute(self.row) > 1

    def get_index(self) -> list[int]:
        sheet = self.row.parentNode
        ret_index = []
        index = 0
        rows_list = sheet.getElementsByType(TableRow)
        for curr_item in rows_list:
            repeat_num = Row.get_repeat_attribute(curr_item)
            if curr_item != self.row:
                index += repeat_num
                continue
            for _i in range(repeat_num):
                ret_index.append(index)
                index += 1
        return ret_index

    def get_cell_by_index(self, cell_index: int) -> Cell:
        index = 0
        cells_list = self.row.getElementsByType(TableCell)
        for curr_item in cells_list:
            repeat_num = Cell.get_repeat_attribute(curr_item)
            for _i in range(repeat_num):
                if index == cell_index:
                    return Cell(curr_item)
                index += 1
        return None

    def get_cell_by_index_expanded(self, cell_index: int) -> Cell:
        curr_cell = self.get_cell_by_index(cell_index)
        if curr_cell is None:
            message = "unable to get cell by index"
            raise RuntimeError(message)
        if not curr_cell.is_repeated():
            return curr_cell
        curr_cell.expand_cell()
        curr_cell = self.get_cell_by_index(cell_index)
        if curr_cell is None:
            message = "unable to get cell by index"
            raise RuntimeError(message)
        if curr_cell.is_repeated():
            _LOGGER.warning("unable to join cells - target cell is repeated")
            return None
        return curr_cell

    def is_empty(self):
        cells_list = self.row.getElementsByType(TableCell)
        for cell_item in cells_list:
            cell_text = Cell(cell_item).get_text()
            if cell_text:
                return False
        return True

    def get_values(self, *, expand_repeated: bool = True) -> list[str]:
        cells_list = self.row.getElementsByType(TableCell)
        if expand_repeated is False:
            return [Cell(cell_item).get_text() for cell_item in cells_list]
        ret_list = []
        for curr_item in cells_list:
            curr_cell = Cell(curr_item)
            repeat_num = curr_cell.get_repeat()
            cell_text = curr_cell.get_text()
            ret_list.extend([cell_text] * repeat_num)
        return ret_list


class Sheet:

    def __init__(self, table: Table):
        self.table: Table = table

    def count_rows(self) -> int:
        rows_list = self.table.getElementsByType(TableRow)
        return len(rows_list)

    def count_cells(self) -> int:
        count = 0
        rows_list = self.table.getElementsByType(TableRow)
        for item in rows_list:
            row = Row(item)
            count += row.count_cells()
        return count

    def add_new_row(self):
        item = TableRow()
        self.add_row(item)
        row = Row(item)
        columns = self.table.getElementsByType(TableColumn)
        cols_len = len(columns)
        if cols_len == 0:
            self.add_new_column()
            return row
        for _i in range(cols_len):
            row.add_new_cell()
        return row

    def add_row(self, item):
        self.table.addElement(item)

    def add_rows(self, item_list: list[Row]):
        for item in item_list:
            self.table.addElement(item.row)

    def add_new_column(self):
        item = TableColumn()
        self.table.addElement(item)
        # add a new cell to each row
        for row in self.table.getElementsByType(TableRow):
            cell = TableCell()
            row.addElement(cell)

    def get_row_by_index(self, row_index: int) -> Row:
        index = 0
        rows_list = self.table.getElementsByType(TableRow)
        for curr_item in rows_list:
            repeat_num = Row.get_repeat_attribute(curr_item)
            for _i in range(repeat_num):
                if index == row_index:
                    return Row(curr_item)
                index += 1
        return None

    def get_cell_by_index(self, row_index: int, cell_index: int) -> Cell:
        curr_row = self.get_row_by_index(row_index)
        if curr_row is None:
            return None
        return curr_row.get_cell_by_index(cell_index)

    def get_rows(self, row_start_index=None, row_end_index=None) -> list[Row]:
        all_rows = list(self.table.getElementsByType(TableRow))  ## copy of list
        if row_end_index:
            all_rows = all_rows[:row_end_index]
        if row_start_index:
            all_rows = all_rows[row_start_index:]
        return [Row(item) for item in all_rows]

    def remove_empty_rows(self, row_start_index=None, row_end_index=None) -> bool:
        all_rows: list[Row] = self.get_rows(row_start_index, row_end_index)
        removed = False
        for curr_row in all_rows:
            if curr_row.is_empty():
                self.table.removeChild(curr_row.row)
                removed = True
        return removed

    def expand_rows(self):
        #
        # duplicate rows
        #
        all_rows = list(self.table.getElementsByType(TableRow))  # copy of list
        new_rows_dict = {}
        for row in all_rows:
            repeat_num = Row.get_repeat_attribute(row)
            if repeat_num < 2:
                continue
            row.removeAttribute("numberrowsrepeated")
            new_items_list = []
            parent = row.parentNode
            for _i in range(repeat_num - 1):
                new_row = copy.deepcopy(row)
                new_items_list.append(new_row)
            new_rows_dict[row] = new_items_list

        for row, new_items_list in new_rows_dict.items():
            parent = row.parentNode
            for new_item in new_items_list:
                parent.insertBefore(new_item, row)

    def sort_rows(self, row_transformer, row_start_index=None, row_end_index=None):
        all_rows = self.get_rows(row_start_index, row_end_index)

        row_data_list = []

        # for curr_row in table_rows:
        for curr_row in all_rows:
            row_data = row_transformer(curr_row)
            row_data_list.append(
                (
                    row_data,
                    curr_row,
                ),
            )

        rows_sorted = sorted(row_data_list, key=lambda item: item[0])
        if rows_sorted == row_data_list:
            return

        # remove all rows from table
        for row in self.table.getElementsByType(TableRow):
            try:
                self.table.removeChild(row)
            except ValueError as exc:
                _LOGGER.warning("unable to remove row: %s", exc)

        table_row_first = row_start_index  ## first proper index
        if not table_row_first:
            table_row_first = 0
        table_row_end = row_end_index  ## position after last proper index
        if not table_row_end:
            table_row_end = len(all_rows)

        # add all rows in changed order
        row_new_order = [item[1] for item in rows_sorted]
        self.add_rows(all_rows[:table_row_first])
        self.add_rows(row_new_order)
        self.add_rows(all_rows[table_row_end:])

    def get_values(self, *, expand_repeated: bool = True) -> list[list[str]]:
        rows_list = self.table.getElementsByType(TableRow)
        if expand_repeated is False:
            return [Row(row_item).get_values(expand_repeated=False) for row_item in rows_list]
        ret_list = []
        for curr_item in rows_list:
            curr_row = Row(curr_item)
            repeat_num = curr_row.get_repeat()
            curr_values = curr_row.get_values(expand_repeated=True)
            ret_list.extend([curr_values] * repeat_num)
        return ret_list


class Document:

    def __init__(self, document: odf.opendocument.OpenDocument):
        self.document: odf.opendocument.OpenDocument = document

    @staticmethod
    def new(sheet_name: str = "Sheet1") -> "Document":
        doc = odf.opendocument.OpenDocumentSpreadsheet()
        sheet = Table(name=sheet_name)
        doc.spreadsheet.addElement(sheet)
        return Document(doc)

    @staticmethod
    def load(file_path: str) -> "Document":
        doc = odf_load(file_path)
        return Document(doc)

    def save(self, output_file):
        """Save document to file.

        :param output_file: file path or file object
        """
        self.document.save(output_file)

    def reload(self):
        """Save to temporary file and then load from that file.

        Sometimes operations on cells and rows on odfpy level causes
        structure unstable. This method makes document stable again.
        """
        with tempfile.TemporaryFile() as temp_file:
            self.save(temp_file)
            self.document = odf_load(temp_file)

    def get_sheet(self, sheet_name) -> Sheet:
        for table in self.document.spreadsheet.getElementsByType(Table):
            if table.getAttribute("name") == sheet_name:
                return Sheet(table)
        return None
