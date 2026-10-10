from __future__ import annotations

from io import BytesIO
from pathlib import Path
import sys
import unittest

import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

from app import build_targeted_export_payload
from workbook_mapper import (
    MASTER_DATA_ROW_CODES,
    SUPPORT_EXPORT_ITEMS,
    SheetTemplate,
    build_master_data_view,
    build_support_view,
)


def _template(sheet_name: str, rows: list[dict[str, str]], codes: list[str], labels: list[str]) -> SheetTemplate:
    return SheetTemplate(
        sheet_name=sheet_name,
        title=sheet_name,
        column_codes=codes,
        column_labels=labels,
        rows=rows,
    )


def _support_templates() -> dict[str, SheetTemplate]:
    codes_by_sheet = {
        sheet_name: sorted(
            {
                row_code
                for _, row_code, source_sheet, _, _ in SUPPORT_EXPORT_ITEMS
                if row_code and source_sheet == sheet_name
            }
        )
        for sheet_name in ("t7_c", "t7_b1")
    }
    return {
        "t7_c": _template(
            "t7_c",
            [
                {"row_code": code, "row_title": f"Forrás megnevezés {code}", "dyn_row_serial": "0"}
                for code in codes_by_sheet["t7_c"]
            ],
            ["3"],
            ["támogatás összege"],
        ),
        "t7_b1": _template(
            "t7_b1",
            [
                {"row_code": code, "row_title": f"Forrás megnevezés {code}", "dyn_row_serial": "0"}
                for code in codes_by_sheet["t7_b1"]
            ],
            ["3", "4"],
            ["Egység", "eFt"],
        ),
    }


def _row(farm: str, code: str, osz: str, value: str, sheet: str) -> dict[str, str]:
    return {
        "akod": farm,
        "sor": code,
        "osz": osz,
        "ertek": value,
        "dimenzio1": "",
        "dyn_row_serial": "0",
        "sheet_name": sheet,
    }


class TargetedExportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.farm = "TEST-01"
        self.templates = _support_templates()
        self.bulk_data = pd.DataFrame(
            [
                _row(self.farm, "m7094", "3", "1,25", "t7_c"),
                _row(self.farm, "m7080", "3", "1.25", "t7_c"),
                _row(self.farm, "m70801", "3", "0", "t7_c"),
                _row(self.farm, "m70802", "3", "", "t7_c"),
                _row(self.farm, "m70803", "3", "1 234,56", "t7_c"),
                _row(self.farm, "m7216", "3", "2,75", "t7_c"),
                _row(self.farm, "m7407", "3", "99", "t7_b1"),
                _row(self.farm, "m7407", "4", "2,50", "t7_b1"),
                _row(self.farm, "m0011", "3", "20 - Őstermelők családi gazdasága (ÖCsG)", "t0"),
                _row(self.farm, "m0001", "3", "2 - SZJA", "t0"),
                _row(self.farm, "m0005", "3", "4 - biogazdálkodást is folytat", "t0"),
                _row(self.farm, "m0075", "3", "2016", "t0"),
                _row(self.farm, "m0002", "3", "3 - ÁFA általános szabályai szerint adózik", "t0"),
                _row(self.farm, "m0018", "3", "2 - egyszeres könyvelés", "t0"),
                _row(self.farm, "m5411", "5", "3.5", "t5_c"),
                _row(self.farm, "m100", "3", "12.5", "t1_a"),
                _row(self.farm, "m100", "4", "0", "t1_a"),
                _row(self.farm, "m101", "3", "0", "t1_a"),
            ]
        )

    def _all_templates(self) -> dict[str, SheetTemplate]:
        templates = dict(self.templates)
        templates.update(
            {
                "t0": _template(
                    "t0",
                    [
                        {
                            "row_code": code,
                            "row_title": f"Törzsadat {code}",
                            "dyn_row_serial": "0",
                        }
                        for code in MASTER_DATA_ROW_CODES
                    ],
                    ["3"],
                    ["törzsadat"],
                ),
                "t5_c": _template(
                    "t5_c",
                    [{"row_code": "m5411", "row_title": "Őszi búza", "dyn_row_serial": "0"}],
                    ["5"],
                    ["Záróérték"],
                ),
                "t6_b": _template("t6_b", [], ["12"], ["Zárókészlet"]),
                "t6_a": _template("t6_a", [], ["12"], ["Záróállomány"]),
                "t1_a": _template(
                    "t1_a",
                    [
                        {"row_code": "m100", "row_title": "Földterületi próbasor", "dyn_row_serial": "0"},
                        {"row_code": "m101", "row_title": "Nulla értékű próbasor", "dyn_row_serial": "0"},
                    ],
                    ["3", "4", "5", "6", "7", "8", "9", ""],
                    ["Adat 3", "Adat 4", "Adat 5", "Adat 6", "Adat 7", "Adat 8", "Adat 9", ""],
                ),
            }
        )
        return templates

    def test_master_data_view_contains_t0_rows_and_derived_akg_flag(self) -> None:
        templates = self._all_templates()

        view = build_master_data_view(self.bulk_data, templates, self.farm)

        self.assertEqual(list(view.columns), ["Üzemkód", "RowCode", "Törzsadat megnevezése", "Érték"])
        self.assertEqual(len(view), len(MASTER_DATA_ROW_CODES) + 1)
        self.assertEqual(view.iloc[0]["RowCode"], "m0011")
        self.assertEqual(view.iloc[0]["Érték"], "20 - Őstermelők családi gazdasága (ÖCsG)")
        akg = view.loc[view["RowCode"].eq("AKG")].iloc[0]
        self.assertEqual(akg["Törzsadat megnevezése"], "AKG gazdálkodás")
        self.assertEqual(akg["Érték"], "Igen")

    def test_master_only_export_has_dedicated_filename_and_formatting(self) -> None:
        templates = self._all_templates()
        payload, filename, mime = build_targeted_export_payload(
            selected_farms=[self.farm],
            template_bundle={"sheets": templates},
            bulk_data=self.bulk_data,
            subjects=["master"],
        )

        self.assertEqual(filename, "TEST-01_torzsadatok.xlsx")
        self.assertEqual(
            mime,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        workbook = load_workbook(BytesIO(payload), data_only=True)
        self.assertEqual(workbook.sheetnames, ["Törzsadatok"])
        worksheet = workbook["Törzsadatok"]
        self.assertEqual(worksheet.max_row, len(MASTER_DATA_ROW_CODES) + 2)
        self.assertEqual(worksheet.cell(worksheet.max_row, 2).value, "AKG")
        self.assertEqual(worksheet.cell(worksheet.max_row, 4).value, "Igen")
        self.assertEqual(worksheet.cell(worksheet.max_row, 4).fill.fgColor.rgb, "00D9EAD3")
        self.assertEqual(worksheet.freeze_panes, "A2")
        self.assertTrue(worksheet.auto_filter.ref)

    def test_support_list_keeps_all_rows_and_distinguishes_source_states(self) -> None:
        view = build_support_view(self.bulk_data, self.templates, self.farm)

        self.assertEqual(len(view), len(SUPPORT_EXPORT_ITEMS))
        self.assertEqual(view.iloc[0]["FADN-sorkód"], "m7094")
        self.assertEqual(view.iloc[0]["Összeg (Ft)"], 1250)
        self.assertEqual(view.iloc[0]["Sor jellege"], "Összesítő")

        zero = view.loc[view["FADN-sorkód"].eq("m70801")].iloc[0]
        self.assertEqual(zero["Összeg (Ft)"], 0)
        self.assertEqual(zero["Adatállapot"], "Forrásban 0")

        blank = view.loc[view["FADN-sorkód"].eq("m70802")].iloc[0]
        self.assertEqual(blank["Összeg (Ft)"], 0)
        self.assertEqual(blank["Adatállapot"], "Nincs kitöltött forrásérték")

        converted = view.loc[view["FADN-sorkód"].eq("m70803")].iloc[0]
        self.assertEqual(converted["Összeg (Ft)"], 1_234_560)

        cis_yf = view.loc[view["FADN-sorkód"].eq("m7407")].iloc[0]
        self.assertEqual(cis_yf["Összeg (Ft)"], 2500)

        smallholder = view.loc[view["Célűrlap jogcíme"].eq("Kistermelői támogatási rendszer")].iloc[0]
        self.assertTrue(pd.isna(smallholder["Összeg (Ft)"]))
        self.assertEqual(smallholder["Adatállapot"], "Nincs azonosított FADN-sorkód")

    def test_export_contains_support_sheet_after_land_and_preserves_land_export(self) -> None:
        templates = self._all_templates()
        payload, filename, mime = build_targeted_export_payload(
            selected_farms=[self.farm],
            template_bundle={"sheets": templates},
            bulk_data=self.bulk_data,
            subjects=["autumn", "stocks", "animals", "land", "support"],
        )

        self.assertEqual(filename, "TEST-01_celzott_export.xlsx")
        self.assertEqual(
            mime,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        workbook = load_workbook(BytesIO(payload), data_only=True)
        self.assertEqual(
            workbook.sheetnames,
            [
                "Ősszel vetett terület",
                "Készletek",
                "Állatok",
                "Földterületi adatok",
                "Támogatási jogcímek",
            ],
        )

        land = workbook["Földterületi adatok"]
        self.assertEqual(land.max_row, 2)
        self.assertEqual(land.cell(2, 2).value, "m100")
        self.assertEqual(land.cell(2, 4).value, 12.5)

        support = workbook["Támogatási jogcímek"]
        self.assertEqual(support.max_row, len(SUPPORT_EXPORT_ITEMS) + 1)
        self.assertEqual(support.cell(2, 3).value, "m7094")
        self.assertEqual(support.cell(2, 5).value, 1250)
        self.assertEqual(support.cell(2, 5).number_format, "#,##0")
        self.assertEqual(support.freeze_panes, "A2")
        self.assertTrue(support.auto_filter.ref)

    def test_support_only_export_has_dedicated_filename(self) -> None:
        payload, filename, _ = build_targeted_export_payload(
            selected_farms=[self.farm],
            template_bundle={"sheets": self.templates},
            bulk_data=self.bulk_data,
            subjects=["support"],
        )

        self.assertEqual(filename, "TEST-01_tamogatasi_jogcimek.xlsx")
        workbook = load_workbook(BytesIO(payload), data_only=True)
        self.assertEqual(workbook.sheetnames, ["Támogatási jogcímek"])


if __name__ == "__main__":
    unittest.main()
