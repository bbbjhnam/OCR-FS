from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


WORKSPACE_MODULE = Path(__file__).parent / "pdf_12_fix" / "local_ocr_server.py"
MODULE_PATH = (
    WORKSPACE_MODULE
    if WORKSPACE_MODULE.exists()
    else Path(__file__).resolve().parents[1] / "local_ocr_server.py"
)
SPEC = importlib.util.spec_from_file_location("pdf12_upgrade", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


class UpgradeRegressions(unittest.TestCase):
    def test_income_scope_skips_table_of_contents(self):
        lines = MODULE.add_continuation_candidates(
            [
                "목 차",
                "손 익 계 산 서 ........ 5",
                "제 조 원 가 명 세 서 ........ 6",
                "현 금 흐 름 표 ........ 7",
                "이 익 잉 여 금 처 분 계 산 서 ........ 8",
                "감사보고서 본문",
                "손 익 계 산 서",
                "제 1 기 2026년 1월 1일부터 2026년 6월 30일까지",
                "매 출 액 1,200,000",
                "매 출 원 가 700,000",
                "영 업 이 익 300,000",
                "당 기 순 이 익 250,000",
                "현 금 흐 름 표",
                "영업활동현금흐름 100,000",
            ]
        )

        start = MODULE.income_statement_start(lines)
        scope = MODULE.income_statement_scope(lines)

        self.assertGreater(start, 4)
        self.assertTrue(any("매출액" in MODULE.normalize(line) for line in scope))
        self.assertTrue(MODULE.has_single_period_income_layout(lines))
        self.assertTrue(MODULE.is_first_fiscal_period_income_layout(lines))

    def test_fast_mode_never_escalates_to_precise_ocr(self):
        with tempfile.TemporaryDirectory(dir=MODULE.ROOT) as temp_name:
            root = Path(temp_name)
            source = root / "scan.pdf"
            source.write_bytes(b"dummy")
            final_dir = root / "final"
            final_dir.mkdir()

            def fake_fast(_command, _input, output, sidecar):
                output.write_bytes(b"pdf")
                sidecar.write_text("매출액 1,000 900", encoding="utf-8")
                return {
                    "ok": True,
                    "profile": MODULE.FAST_PROFILE["name"],
                    "text": "매출액 1,000 900",
                    "error": "",
                }

            with (
                patch.object(MODULE, "extract_text", return_value=""),
                patch.object(MODULE, "has_meaningful_text", return_value=False),
                patch.object(MODULE, "run_fast_ocr", side_effect=fake_fast) as fast,
                patch.object(MODULE, "run_best_ocr") as precise,
            ):
                result = MODULE.process_one_file(
                    (1, "scan.pdf", source), final_dir, ["ocrmypdf"], "fast"
                )

            self.assertEqual(fast.call_count, 1)
            precise.assert_not_called()
            self.assertIn("빠른 OCR 완료", result["ocrStatus"])

    def test_text_pdf_skips_ocr_and_duplicate_names_get_unique_outputs(self):
        meaningful = "재무상태표 " + "가나다라마바사 1,000 900\n" * 20
        with tempfile.TemporaryDirectory(dir=MODULE.ROOT) as temp_name:
            root = Path(temp_name)
            final_dir = root / "final"
            final_dir.mkdir()
            first = root / "001_report.pdf"
            second = root / "002_report.pdf"
            first.write_bytes(b"first")
            second.write_bytes(b"second")

            with (
                patch.object(MODULE, "extract_text", return_value=meaningful),
                patch.object(MODULE, "has_meaningful_text", return_value=True),
                patch.object(MODULE, "run_fast_ocr") as fast,
                patch.object(MODULE, "run_best_ocr") as precise,
            ):
                result1 = MODULE.process_one_file(
                    (1, "report.pdf", first), final_dir, ["ocrmypdf"], "fast"
                )
                result2 = MODULE.process_one_file(
                    (2, "report.pdf", second), final_dir, ["ocrmypdf"], "fast"
                )

            fast.assert_not_called()
            precise.assert_not_called()
            self.assertEqual(result1["ocrStatus"], "텍스트 PDF 감지 · OCR 생략")
            self.assertNotEqual(result1["ocrUrl"], result2["ocrUrl"])
            self.assertEqual(len(list(final_dir.glob("*_최종.pdf"))), 2)

    def test_failed_file_has_a_complete_placeholder_result(self):
        result = MODULE.failed_document((3, "broken.pdf", Path("broken.pdf")), ValueError("bad pdf"))
        self.assertEqual(result["ocrStatus"], "처리 실패")
        self.assertEqual(result["ocrError"], "bad pdf")
        self.assertEqual(len(result["rows"]), len(MODULE.REPORT_FIELDS))


if __name__ == "__main__":
    unittest.main()
