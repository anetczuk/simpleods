#
# Copyright (c) 2025, Arkadiusz Netczuk <dev.arnet@gmail.com>
# All rights reserved.
#
# This source code is licensed under the BSD 3-Clause license found in the
# LICENSE file in the root directory of this source tree.
#

import unittest
import tempfile

from testsimpleodf.data import get_data_path

from simpleodf.odfwrapper import get_spreadsheet_coords, Document, Sheet, Row, Cell


class FreeFunctionsTest(unittest.TestCase):
    def test_get_spreadsheet_coords_01(self):
        coords = get_spreadsheet_coords(0, 0)
        self.assertEqual(("1", "A"), coords)

    def test_get_spreadsheet_coords_02(self):
        coords = get_spreadsheet_coords(10, 10)
        self.assertEqual(("11", "K"), coords)

    def test_get_spreadsheet_coords_03(self):
        coords = get_spreadsheet_coords(30, 30)
        self.assertEqual(("31", "AE"), coords)


## ========================================================================


def create_simple_document():
    document = Document.new()
    sheet: Sheet = document.get_sheet("Sheet1")
    sheet.add_new_row()
    return document


class CellTest(unittest.TestCase):
    def test_get_index_01(self):
        document = create_simple_document()
        sheet: Sheet = document.get_sheet("Sheet1")
        row = sheet.get_row_by_index(0)
        self.assertTrue(row is not None)
        cell = row.get_cell_by_index(0)
        cell_index = cell.get_index()
        self.assertEqual([0], cell_index)

    def test_get_index_02(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")
        cell = sheet.get_cell_by_index(0, 0)
        cell_index = cell.get_index()
        self.assertEqual([0, 1, 2, 3, 4, 5, 6], cell_index)

    def test_get_index_03(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")
        cell = sheet.get_cell_by_index(3, 3)
        cell_index = cell.get_index()
        self.assertEqual([3, 4], cell_index)

    def test_clear_cell(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")
        cell = sheet.get_cell_by_index(5, 6)
        self.assertEqual("cc", cell.get_text())

        cell.clear_cell()
        self.assertEqual("", cell.get_text())

    def test_move_content_to(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")

        src_cell = sheet.get_cell_by_index(5, 6)
        self.assertEqual("cc", src_cell.get_text())
        dst_cell = sheet.get_cell_by_index(0, 0)
        self.assertEqual("", dst_cell.get_text())

        src_cell.move_content_to(dst_cell)

        src_cell = sheet.get_cell_by_index(5, 6)
        self.assertEqual("", src_cell.get_text())
        dst_cell = sheet.get_cell_by_index(0, 0)
        self.assertEqual("cc", dst_cell.get_text())

        values = sheet.get_values(expand_repeated=False)
        self.assertEqual([["cc"], ["", "aa", ""], ["", "aa", ""], ["", "bb", ""]], values)


## ========================================================================


class RowTest(unittest.TestCase):
    def test_get_cell_by_index_expanded(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")
        row: Row = sheet.get_row_by_index(0)

        values = row.get_values(expand_repeated=False)
        self.assertEqual([""], values)

        row.get_cell_by_index_expanded(1)

        values = row.get_values(expand_repeated=False)
        self.assertEqual(["", "", "", "", "", "", ""], values)

    def test_get_index_03(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")
        cell = sheet.get_cell_by_index(3, 3)
        cell_index = cell.get_index()
        self.assertEqual([3, 4], cell_index)


## ========================================================================


class SheetTest(unittest.TestCase):
    def test_count_empty(self):
        document = Document.new()
        sheet = document.get_sheet("Sheet1")
        self.assertEqual(0, sheet.count_rows())
        self.assertEqual(0, sheet.count_cells())

    def test_count_filled(self):
        document = Document.new()
        sheet = document.get_sheet("Sheet1")

        sheet.add_new_row()
        ## 1 row with 1 cell
        self.assertEqual(1, sheet.count_rows())
        self.assertEqual(1, sheet.count_cells())

        row = sheet.add_new_row()
        ## 2 rows, every with 1 cell
        self.assertEqual(2, sheet.count_rows())
        self.assertEqual(2, sheet.count_cells())

        row.add_new_column()
        ## 2 rows, every with 2 cell
        self.assertEqual(2, sheet.count_rows())
        self.assertEqual(4, sheet.count_cells())

        sheet.add_new_column()
        ## 2 rows, every with 3 cell
        self.assertEqual(2, sheet.count_rows())
        self.assertEqual(6, sheet.count_cells())

        sheet.add_new_row()
        ## 3 rows, every with 3 cell
        self.assertEqual(3, sheet.count_rows())
        self.assertEqual(9, sheet.count_cells())

    def test_get_cell_by_index(self):
        document = create_simple_document()
        sheet: Sheet = document.get_sheet("Sheet1")
        cell = sheet.get_cell_by_index(0, 0)
        self.assertTrue(cell is not None)

    def test_get_values(self):
        document = create_simple_document()
        sheet: Sheet = document.get_sheet("Sheet1")
        values = sheet.get_values()
        self.assertEqual([[""]], values)

    def test_get_values_repeated_01(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")
        values = sheet.get_values(expand_repeated=False)
        self.assertEqual([[""], ["", "aa", ""], ["", "aa", ""], ["", "bb", "cc"]], values)

    def test_get_values_repeated_02(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")
        values = sheet.get_values(expand_repeated=True)
        self.assertEqual(
            [
                ["", "", "", "", "", "", ""],
                ["", "", "", "", "", "", ""],
                ["", "", "", "", "", "", ""],
                ["", "", "", "aa", "aa", "", ""],
                ["", "", "", "aa", "aa", "", ""],
                ["", "", "", "", "bb", "bb", "cc"],
            ],
            values,
        )

    def test_get_values_01(self):
        document = create_simple_document()
        sheet: Sheet = document.get_sheet("Sheet1")
        row = sheet.add_new_row()
        row.add_new_column("aaa")
        row.add_new_column("aaa")

        values = sheet.get_values()
        self.assertEqual([["", "", ""], ["", "aaa", "aaa"]], values)

        values = sheet.get_values(expand_repeated=False)
        self.assertEqual([["", "", ""], ["", "aaa", "aaa"]], values)

    def test_remove_empty_rows(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")

        values = sheet.get_values(expand_repeated=False)
        self.assertEqual([[""], ["", "aa", ""], ["", "aa", ""], ["", "bb", "cc"]], values)

        sheet.remove_empty_rows()

        values = sheet.get_values(expand_repeated=False)
        self.assertEqual([["", "aa", ""], ["", "aa", ""], ["", "bb", "cc"]], values)

    def test_expand_rows(self):
        data_path = get_data_path("repeated.ods")
        document = Document.load(data_path)
        sheet: Sheet = document.get_sheet("Sheet1")

        values = sheet.get_values(expand_repeated=False)
        self.assertEqual([[""], ["", "aa", ""], ["", "aa", ""], ["", "bb", "cc"]], values)

        sheet.expand_rows()

        values = sheet.get_values(expand_repeated=False)
        self.assertEqual([[""], [""], [""], ["", "aa", ""], ["", "aa", ""], ["", "bb", "cc"]], values)

    def test_sort(self):
        document = create_simple_document()
        sheet: Sheet = document.get_sheet("Sheet1")
        sheet.add_new_row()
        sheet.add_new_column()

        ## set unsorted
        cell01: Cell = sheet.get_cell_by_index(0, 1)
        cell10: Cell = sheet.get_cell_by_index(1, 0)
        cell01.set_text("bbb")
        cell10.set_text("aaa")

        ## check unsorted
        self.assertEqual(("1", "B"), cell01.get_spreadsheet_coords())
        self.assertEqual(("2", "A"), cell10.get_spreadsheet_coords())
        self.assertEqual("bbb", cell01.get_text())
        self.assertEqual("aaa", cell10.get_text())

        ## sort
        def rows_sorter(row: Row):
            row_vals = row.get_values(expand_repeated=False)
            return sorted(row_vals)

        sheet.sort_rows(row_transformer=rows_sorter)

        ## check sorted
        self.assertEqual(("1", "A"), cell10.get_spreadsheet_coords())  ## moved from row 1 to row 0
        self.assertEqual(("2", "B"), cell01.get_spreadsheet_coords())  ## moved from row 0 to row 1

        cell00: Cell = sheet.get_cell_by_index(0, 0)
        cell11: Cell = sheet.get_cell_by_index(1, 1)
        self.assertEqual("aaa", cell00.get_text())
        self.assertEqual("bbb", cell11.get_text())


## ========================================================================


class DocumentTest(unittest.TestCase):
    def test_load_empty_document_nonexisting(self):
        data_path = "not_existing_file_path.ods"
        self.assertRaises(FileNotFoundError, Document.load, data_path)

    def test_load_empty_document(self):
        try:
            data_path = get_data_path("empty.ods")
            Document.load(data_path)
        except Exception:  # pylint: disable=W0718 # noqa: BLE001
            self.fail("function raised Exception unexpectedly!")

    def test_load_sheet_nonexisting(self):
        data_path = get_data_path("empty.ods")
        document = Document.load(data_path)
        sheet = document.get_sheet("non_existing_sheet")
        self.assertTrue(sheet is None)

    def test_load_sheet(self):
        data_path = get_data_path("empty.ods")
        doc = Document.load(data_path)
        sheet = doc.get_sheet("Sheet1")
        self.assertTrue(sheet is not None)

    def test_save_fileobject(self):
        try:
            data_path = get_data_path("empty.ods")
            doc = Document.load(data_path)
            with tempfile.NamedTemporaryFile() as temp_file:
                doc.save(temp_file)
        except Exception:  # pylint: disable=W0718 # noqa: BLE001
            self.fail("function raised Exception unexpectedly!")

    def test_reload(self):
        data_path = get_data_path("empty.ods")
        doc = Document.load(data_path)
        old_doc = doc.document
        doc.reload()
        self.assertNotEqual(id(old_doc), id(doc.document))
