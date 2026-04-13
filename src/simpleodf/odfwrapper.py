"""Spreadsheet wrappers for odfpy objects."""

#
# Copyright (c) 2026, Arkadiusz Netczuk <dev.arnet@gmail.com>
# All rights reserved.
#
# This source code is licensed under the BSD 3-Clause license found in the
# LICENSE file in the root directory of this source tree.
#

import logging
from collections.abc import Callable

import copy
import tempfile

import odf
from odf.opendocument import load as odf_load

from odf.table import Table, TableColumn
from odf.table import TableRow
from odf.table import TableCell
from odf.text import P

_LOGGER = logging.getLogger(__name__)


## =======================================================


def get_spreadsheet_coords(row_index: int, cell_index: int) -> tuple[str, str]:
    """Convert numeric cell index coords to spreadsheet cell address."""
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
    """Cell wrapper."""

    def __init__(self, cell: TableCell):
        """Wrap odfpy object."""
        self.cell: TableCell = cell

    def get_parent(self) -> "Row":
        """Get parent row object."""
        row = self.cell.parentNode
        return Row(row)

    def clear_cell(self):
        """Clear cell content."""
        # Remove all child nodes (usually <text:p>)
        for child in list(self.cell.childNodes):  # snapshot list to avoid iteration issues
            self.cell.removeChild(child)

        # remove attributes
        for attr in ["value", "valuetype", "formula", "stylename"]:
            if self.cell.getAttribute(attr):
                self.cell.removeAttribute(attr)

    def copy_cell(self) -> "Cell":
        """Copy cell content."""
        ## deep copy of cell object is not advised
        ## return copy.deepcopy(self.cell)

        dst_cell = TableCell()
        target_cell = Cell(dst_cell)
        self.copy_cell_to(target_cell)
        return target_cell

    def copy_cell_to(self, target_cell: "Cell"):
        """Copy cell's content into target cell."""
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
        """Move cell's content into target cell."""
        self.copy_cell_to(target_cell)
        self.clear_cell()

    def get_repeat(self) -> int:
        """Get row's repeat attribute value."""
        return Cell.get_repeat_attribute(self.cell)

    def is_repeated(self) -> bool:
        """Check if row is repeated."""
        return Cell.get_repeat_attribute(self.cell) > 1

    @staticmethod
    def get_repeat_attribute(cell) -> int:
        """Get row's repeat attribute value."""
        return int(cell.getAttribute("numbercolumnsrepeated") or 1)

    def get_index(self) -> list[int]:
        """Get cell index.

        If cell is repeated then return all it's indexes.
        """
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
        """Convert numeric cell index coords to spreadsheet cell address."""
        row: Row = self.get_parent()
        row_index = row.get_index()
        cell_index = self.get_index()
        return get_spreadsheet_coords(row_index[0], cell_index[0])

    def get_text(self) -> str:
        """Get cell's text."""
        text = ""
        for p in self.cell.getElementsByType(P):
            if p.firstChild:
                text += str(p.firstChild.data)
        return text

    def set_text(self, content: str):
        """Set cell's content."""
        for child in list(self.cell.childNodes):  # snapshot list to avoid iteration issues
            self.cell.removeChild(child)
        self.cell.addElement(P(text=content))

    def expand_cell(self):
        """Replace cell's 'repeat' attribute with corresponding number of duplicated cells."""
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
    """Row wrapper."""

    def __init__(self, row: TableRow):
        """Wrap odfpy object."""
        self.row: TableRow = row

    def get_parent(self) -> "Sheet":
        """Get parent sheet object."""
        sheet = self.row.parentNode
        return Sheet(sheet)

    def count_cells(self) -> int:
        """Count cells in row.

        Ignores repeat attribute.
        """
        cell_list = self.row.getElementsByType(TableCell)
        return len(cell_list)

    def add_new_column(self, content=None):
        """Add column object and corresponding cell."""
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
        """Insert new cell at end of row."""
        item = TableCell()
        self.row.addElement(item)

    def add_cell(self, item: TableCell):
        """Insert cell at end of row."""
        self.row.addElement(item)

    def add_cells(self, item_list: list[TableCell]):
        """Insert cells at end of row."""
        for item in item_list:
            self.row.addElement(item)

    def add_cells_before(self, item_list: list[Cell], reference_cell):
        """Insert cell into row."""
        for item in item_list:
            self.row.insertBefore(item.cell, reference_cell)

    def get_repeat(self) -> int:
        """Get row's repeat attribute value."""
        return Row.get_repeat_attribute(self.row)

    @staticmethod
    def get_repeat_attribute(row) -> int:
        """Get row's repeat attribute value."""
        return int(row.getAttribute("numberrowsrepeated") or 1)

    def is_repeated(self) -> bool:
        """Check if row is repeated."""
        return Row.get_repeat_attribute(self.row) > 1

    def get_index(self) -> list[int]:
        """Get row index.

        If row is repeated then return all it's indexes.
        """
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
        """Get cell by given index."""
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
        """Get cell by given index.

        If pointed cell is repeated then duplicate the cell before returning.
        """
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

    def is_empty(self) -> bool:
        """Check if row does not contain any text."""
        cells_list = self.row.getElementsByType(TableCell)
        for cell_item in cells_list:
            cell_text = Cell(cell_item).get_text()
            if cell_text:
                return False
        return True

    def get_values(self, *, expand_repeated: bool = True) -> list[str]:
        """Convert row's all cells to list."""
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
    """Sheet wrapper."""

    def __init__(self, table: Table):
        """Wrap odfpy object."""
        self.table: Table = table

    def count_rows(self) -> int:
        """Get number of rows in sheet."""
        rows_list = self.table.getElementsByType(TableRow)
        return len(rows_list)

    def count_cells(self) -> int:
        """Get number of cells in sheet."""
        cell_list = self.table.getElementsByType(TableCell)
        return len(cell_list)

    def add_new_row(self):
        """Add empty row to sheet."""
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
        """Add row to sheet."""
        self.table.addElement(item)

    def add_rows(self, item_list: list[Row]):
        """Add rows to sheet."""
        for item in item_list:
            self.table.addElement(item.row)

    def add_new_column(self):
        """Add empty column."""
        item = TableColumn()
        self.table.addElement(item)
        # add a new cell to each row
        for row in self.table.getElementsByType(TableRow):
            cell = TableCell()
            row.addElement(cell)

    def get_row_by_index(self, row_index: int) -> Row:
        """Get row by given index."""
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
        """Get cell by given index."""
        curr_row = self.get_row_by_index(row_index)
        if curr_row is None:
            return None
        return curr_row.get_cell_by_index(cell_index)

    def get_rows(self, row_start_index=None, row_end_index=None) -> list[Row]:
        """Get rows by given index span."""
        all_rows = list(self.table.getElementsByType(TableRow))  ## copy of list
        if row_end_index:
            all_rows = all_rows[:row_end_index]
        if row_start_index:
            all_rows = all_rows[row_start_index:]
        return [Row(item) for item in all_rows]

    def remove_empty_rows(self, row_start_index=None, row_end_index=None) -> bool:
        """Remove empty rows.

        Row is treated as empty when it's cells do not contain content.
        """
        all_rows: list[Row] = self.get_rows(row_start_index, row_end_index)
        removed = False
        for curr_row in all_rows:
            if curr_row.is_empty():
                self.table.removeChild(curr_row.row)
                removed = True
        return removed

    def expand_rows(self):
        """Duplicate rows marked as repeated."""
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

    def sort_rows(self, row_transformer: Callable, row_start_index=None, row_end_index=None):
        """Sort rows.

        :param row_transformer: callable calculating representation of rows
        :param row_start_index: optional start row
        :param row_end_index: optional end row
        """
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
        """Convert all cells to 2D matrix."""
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


class SpreadsheetDocument:
    """Document wrapper."""

    def __init__(self, document: odf.opendocument.OpenDocumentSpreadsheet):
        """Wrap odfpy object."""
        self.document: odf.opendocument.OpenDocumentSpreadsheet = document

    @staticmethod
    def new(sheet_name: str = "Sheet1") -> "SpreadsheetDocument":
        """Create new spreadsheet document with empty table sheet."""
        doc = odf.opendocument.OpenDocumentSpreadsheet()
        sheet = Table(name=sheet_name)
        doc.spreadsheet.addElement(sheet)
        return SpreadsheetDocument(doc)

    @staticmethod
    def load(file_path: str) -> "SpreadsheetDocument":
        """Load spreadsheet document."""
        doc = odf_load(file_path)
        return SpreadsheetDocument(doc)

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

    def get_sheet(self, sheet_name: str) -> Sheet:
        """Get spreadsheet by name.

        :param str sheet_name: spreadsheet name
        """
        for table in self.document.spreadsheet.getElementsByType(Table):
            if table.getAttribute("name") == sheet_name:
                return Sheet(table)
        return None
