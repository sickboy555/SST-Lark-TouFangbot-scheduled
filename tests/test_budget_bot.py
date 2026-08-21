import datetime as dt
import os
import unittest
from unittest.mock import patch

from lark_budget_bot import (
    BudgetRow,
    Config,
    dedupe_duplicate_country_blocks,
    filter_zero_rows,
    load_stable_budget_data,
    parse_date_cell,
)


class BudgetParsingRegressionTests(unittest.TestCase):
    def test_config_filters_zero_rows_by_default(self):
        with patch.dict(
            os.environ,
            {
                "LARK_APP_ID": "app",
                "LARK_APP_SECRET": "secret",
                "LARK_SHEET_URL": "https://example.com/sheets/token?sheet=sheet",
                "LARK_CHAT_ID": "chat",
                "SEND_MODE": "app",
            },
            clear=False,
        ):
            os.environ.pop("FILTER_ZERO_ROWS", None)
            self.assertTrue(Config.from_env().filter_zero_rows)

    def test_filter_zero_rows_hides_only_two_day_zero_rows(self):
        rows = [
            BudgetRow("MX", "MX-001", 0, 0),
            BudgetRow("MX", "MX-002", 100, 0),
            BudgetRow("MX", "MX-003", 0, 100),
            BudgetRow("MX", "MX_汇总", 100, 100),
        ]

        result = filter_zero_rows(rows)

        self.assertEqual(
            [(row.package, row.today, row.yesterday) for row in result],
            [("MX-002", 100, 0), ("MX-003", 0, 100), ("MX_汇总", 100, 100)],
        )

    def test_duplicate_country_blocks_keep_the_last_block(self):
        rows = [
            BudgetRow("墨西哥", "MX-001", 100, 90),
            BudgetRow("墨西哥", "MX_汇总", 100, 90),
            BudgetRow("阿根廷", "AR-iOS-H5", 10, 10),
            BudgetRow("阿根廷", "AR-EMI-D01", 20, 20),
            BudgetRow("阿根廷", "AR_汇总", 30, 30),
            BudgetRow("巴西", "BR-001", 200, 180),
            BudgetRow("巴西", "BR_汇总", 200, 180),
            BudgetRow("阿根廷", "AR-EMI-D01", 20, 20),
            BudgetRow("阿根廷", "AR-IOS-H5", 10, 10),
            BudgetRow("阿根廷", "AR_汇总", 30, 30),
        ]

        result = dedupe_duplicate_country_blocks(rows)

        self.assertEqual(
            [(row.country, row.package) for row in result],
            [
                ("墨西哥", "MX-001"),
                ("墨西哥", "MX_汇总"),
                ("巴西", "BR-001"),
                ("巴西", "BR_汇总"),
                ("阿根廷", "AR-EMI-D01"),
                ("阿根廷", "AR-IOS-H5"),
                ("阿根廷", "AR_汇总"),
            ],
        )

    def test_parse_excel_serial_date(self):
        self.assertEqual(parse_date_cell(46255), dt.date(2026, 8, 21))

    def test_stable_loader_retries_when_formula_total_disagrees(self):
        stale = [
            ["全市场总预算", "", 139230, 88620],
            ["国家", "包名", "8月21日", "8月20日"],
            ["墨西哥", "MX_汇总", 26500, 24500],
            ["阿根廷", "AR_汇总", 1500, 1500],
        ]
        fresh = [
            ["全市场总预算", "", 139230, 133270],
            ["国家", "包名", "8月21日", "8月20日"],
            ["墨西哥", "MX_汇总", 26500, 24500],
            ["阿根廷", "AR_汇总", 1500, 1500],
            ["巴西", "BR_汇总", 111230, 107270],
        ]
        calls = iter([stale, fresh, fresh])

        values = load_stable_budget_data(
            lambda: next(calls),
            dt.date(2026, 8, 21),
            retries=3,
            delay_seconds=0,
            sleep_fn=lambda _: None,
        )

        self.assertEqual(values, fresh)

    def test_stable_loader_rejects_unstable_data(self):
        values = [
            ["全市场总预算", "", 140730, 88620],
            ["国家", "包名", "8月21日", "8月20日"],
            ["墨西哥", "MX_汇总", 26500, 24500],
        ]

        with self.assertRaisesRegex(RuntimeError, "总计与明细汇总不一致"):
            load_stable_budget_data(
                lambda: values,
                dt.date(2026, 8, 21),
                retries=3,
                delay_seconds=0,
                sleep_fn=lambda _: None,
            )


if __name__ == "__main__":
    unittest.main()
