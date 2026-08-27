import tempfile
import unittest
from pathlib import Path

import pandas as pd

from data_cache import MonthlyQuotaExceeded, load_with_fallback, parse_feishu_response


class FakeResponse:
    def __init__(self, status_code, payload):
        self.status_code = status_code
        self.payload = payload

    def json(self):
        return self.payload


class DataCacheTests(unittest.TestCase):
    def test_monthly_quota_response_is_recognized(self):
        response = FakeResponse(
            429,
            {
                "code": 99991403,
                "msg": "This month's API call quota has been exceeded",
            },
        )

        with self.assertRaises(MonthlyQuotaExceeded):
            parse_feishu_response(response, "读取电子表格")

    def test_successful_fetch_is_saved_and_returned(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            cache_path = Path(temp_dir) / "snapshot.pkl"
            expected = pd.DataFrame({"年份": ["2026"], "总花费": [100]})

            result = load_with_fallback(lambda: expected, cache_path)

            self.assertEqual(result.source, "live")
            self.assertTrue(cache_path.exists())
            pd.testing.assert_frame_equal(result.dataframe, expected)

    def test_quota_error_uses_last_successful_snapshot(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            cache_path = Path(temp_dir) / "snapshot.pkl"
            expected = pd.DataFrame({"年份": ["2026"], "总花费": [100]})
            load_with_fallback(lambda: expected, cache_path)

            def quota_exceeded():
                raise MonthlyQuotaExceeded("quota exhausted", code=99991403)

            result = load_with_fallback(quota_exceeded, cache_path)

            self.assertEqual(result.source, "cache")
            self.assertIsInstance(result.error, MonthlyQuotaExceeded)
            pd.testing.assert_frame_equal(result.dataframe, expected)

    def test_error_without_snapshot_is_not_hidden(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            cache_path = Path(temp_dir) / "snapshot.pkl"

            def quota_exceeded():
                raise MonthlyQuotaExceeded("quota exhausted", code=99991403)

            with self.assertRaises(MonthlyQuotaExceeded):
                load_with_fallback(quota_exceeded, cache_path)


if __name__ == "__main__":
    unittest.main()
