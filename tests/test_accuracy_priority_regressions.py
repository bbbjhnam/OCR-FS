from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path
from unittest.mock import patch


MODULE_PATH = Path(__file__).resolve().parents[1] / "local_ocr_server.py"
SPEC = importlib.util.spec_from_file_location("pdf12_accuracy", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


def rows_by_item(text: str) -> dict[str, dict]:
    return {row["item"]: row for row in MODULE.extract_financial_rows(text)}


class AccuracyPriorityRegressions(unittest.TestCase):
    def test_complete_period_amounts_are_not_joined(self):
        self.assertEqual(
            MODULE.extract_amounts("영업외수익 674,377 236,625"),
            [674_377, 236_625],
        )

    def test_detached_thousands_are_repaired_only_for_second_full_amount(self):
        self.assertEqual(
            MODULE.extract_amounts("매출액 77,701,095 7,426 698,918"),
            [77_701_095, 7_426_698_918],
        )
        self.assertEqual(
            MODULE.extract_amounts("제품매출 4,758,539,687 3,814,846 887"),
            [4_758_539_687, 3_814_846_887],
        )
        self.assertNotIn("48,232,660", MODULE.normalize_ocr_numbers("잡손실 48,232 660"))
        self.assertEqual(
            MODULE.financial_amounts("잡손실 48,232 660", ("잡손실",)),
            [48_232, 660],
        )

    def test_explicit_loss_applies_to_every_period(self):
        line = "영업손실 684,976,848 -274,793,339"
        self.assertEqual(
            MODULE.apply_loss_row_signs(line, MODULE.financial_amounts(line, ("영업손실",))),
            [-684_976_848, -274_793_339],
        )
        reversal = "당기순손실 1,861,760,879 ~789,046,283"
        self.assertEqual(
            MODULE.apply_loss_row_signs(
                reversal, MODULE.financial_amounts(reversal, ("당기순손실",))
            ),
            [-1_861_760_879, 789_046_283],
        )
        summary = "당기순손실 당기: 475,152,340 전기: -322,715,216"
        self.assertEqual(
            MODULE.apply_loss_row_signs(
                summary, MODULE.financial_amounts(summary, ("당기순손실",))
            ),
            [-475_152_340, 322_715_216],
        )
        mixed = "영업손익 -684,976,848 274,793,339"
        self.assertEqual(
            MODULE.apply_loss_row_signs(mixed, MODULE.financial_amounts(mixed, ("영업손익",))),
            [-684_976_848, -274_793_339],
        )

    def test_complete_row_does_not_absorb_following_account(self):
        lines = [
            "자본금 70,002,000 70,002,000",
            "자본잉여금 1,078,256,260 1,078,256,260",
            "주식발행초과금 1,078,256,260 1,078,256,260",
        ]
        expanded = MODULE.add_continuation_candidates(lines)
        self.assertEqual(expanded, lines)
        self.assertEqual(rows_by_item("\n".join(lines))["자본금"]["current"], 70_002_000)

    def test_common_and_preferred_capital_use_total(self):
        text = "\n".join(
            [
                "재무상태표",
                "Xt 꼼",
                "293,970,000 293,970,000",
                "보통주 A 본 금 200,435,000 200,435,000",
                "반복 페이지 머리글",
                "제 10 기 당기",
                "회사명 메디사피엔스",
                "금액",
                "빈 줄 대체",
                "우선주 자본금 93,535,000 93,535,000",
                "자본잉여금 7,568,756,042 7,568,756,042",
            ]
        )
        capital = rows_by_item(text)["자본금"]
        self.assertEqual((capital["current"], capital["previous"]), (293_970_000, 293_970_000))

    def test_truncated_liability_total_is_crosschecked_by_subtotals(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계 310,899,524 355,802,031",
                "유동부채 1,135,670,909 1,131,056,205",
                "비유동부채 0 0",
                "부채총계 1,1365.670,909 1,131.056,205",
                "자본금 136,892,000 135,892,000",
                "손익계산서",
                "매출액 1,000,000 900,000",
                "영업손실 100,000 90,000",
            ]
        )
        liabilities = rows_by_item(text)["부채총계"]
        self.assertEqual(
            (liabilities["current"], liabilities["previous"]),
            (1_135_670_909, 1_131_056_205),
        )

    def test_full_liability_total_beats_later_ratio_table(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계 14,058,271,974 18,520,489,676",
                "부채총계 59,987,506,266 57,221,500,003",
                "자본금 1,000,000,000 1,000,000,000",
                "주석 비율 분석",
                "부채총계 59,987 4.83% 57,221 2.7% 55,732",
            ]
        )
        liabilities = rows_by_item(text)["부채총계"]
        self.assertEqual(
            (liabilities["current"], liabilities["previous"]),
            (59_987_506_266, 57_221_500_003),
        )

    def test_quarterly_layout_uses_cumulative_columns(self):
        text = "\n".join(
            [
                "분 기 연 결 포 괄 손 익 계 산 서",
                "제 9(당)기 3분기 제 8(전)기 3분기",
                "3개월 당 3분기말 3개월 전 3분기말",
                "Ⅰ. 매출 1,492,200 73,792,200 167,380,000 275,595,375",
                "Ⅴ. 영업손실 (1,844,598,129) (5,974,521,608) (2,499,503,103) (7,313,331,222)",
                "법인세비용차감전순손실 (1,136,779,577) (8,377,299,084) (1,971,208,319) (6,113,470,163)",
                "법인세비용 0 0 0 0",
                "분기순손실 (1,136,779,577) (8,377,299,084) (1,971,208,319) (6,113,470,163)",
                "분 기 연 결 자 본 변 동 표",
                "주석 매출액 999,999,999 888,888,888",
            ]
        )
        rows = rows_by_item(text)
        self.assertEqual((rows["매출액"]["current"], rows["매출액"]["previous"]), (73_792_200, 275_595_375))
        self.assertEqual((rows["영업이익"]["current"], rows["영업이익"]["previous"]), (-5_974_521_608, -7_313_331_222))
        self.assertEqual((rows["당기순이익"]["current"], rows["당기순이익"]["previous"]), (-8_377_299_084, -6_113_470_163))

    def test_net_income_ignores_other_comprehensive_income_reclassification(self):
        text = "\n".join(
            [
                "손익계산서",
                "제13(당)기 2026년 / 제12(전)기 2025년",
                "매출액 21,651,000 0",
                "영업손실 5,932,433,584 9,267,771,774",
                "당기손익 -4,915,108,300 -16,642,044,591",
                "기타포괄손익 30,612,016",
                "후속적으로 당기손익으로 재분류되지 않는 항목 30,612,016",
                "당기손익-공정가치 측정 지정 금융부채 평가손익 30,612,016",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (-4_915_108_300, -16_642_044_591),
        )

    def test_clean_net_income_row_beats_noisy_reference_schedule(self):
        text = "\n".join(
            [
                "손익계산서",
                "과목 당기 전기",
                "매출액 10,000,000 9,000,000",
                "영업손실 800,000,000 300,000,000",
                (
                    "기타처분손실 203 오류수정손실 2101 법인세비용 218 "
                    "당기순손익 219 -920,456,247"
                ),
                "결손금처리계산서",
                "당기순손실 920,456,247 325,706,950",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (-920_456_247, -325_706_950),
        )

    def test_ocr_four_is_repaired_as_loss_suffix(self):
        text = "\n".join(
            [
                "손익계산서",
                "제2(당)기 제1(전)기",
                "매출액 5,000,000,000 4,000,000,000",
                "영업손실 1,084,019,790 1,696,354,870",
                "당 기 순 손 4 1,339,561,455 2,416,748,332",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (-1_339_561_455, -2_416_748_332),
        )

    def test_zero_tax_uses_ocr_truncated_pretax_row(self):
        text = "\n".join(
            [
                "손익계산서",
                "매출액 7,340,632,189 16,558,580,013",
                "영업이익(손실) (3,951,951,552) (3,003,098,626)",
                "법인세차 감 전 순이 (4,189,865,521) (3,195,065,127)",
                "법인세비용 - -",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (-4_189_865_521, -3_195_065_127),
        )

    def test_vertical_duplicate_cells_keep_both_periods(self):
        text = "\n".join(
            [
                "손익계산서",
                "과목 당기 전기",
                "법인세비용차감전순이익",
                "1,538,067,465 1,538,067,465 1,538,067,465 1,538,067,465",
                "1,209,654,264 1,209,654,264 1,209,654,264 1,209,654,264",
                "법인세비용",
                "0 0 0 0",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (1_538_067_465, 1_209_654_264),
        )

    def test_vertical_split_leading_group_is_rejoined(self):
        text = "\n".join(
            [
                "손익계산서",
                "과목 당기 전기",
                "법인세비용차감전순손실",
                "390",
                "535,855",
                "2,144,100,368",
                "법인세비용 0 0",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (-390_535_855, -2_144_100_368),
        )

    def test_earlier_retained_earnings_does_not_hide_later_net_row(self):
        text = "\n".join(
            [
                "재무상태표",
                (
                    "미지급비용 77,538,721 104,914,833 이익잉여금 -17,099,650,793 "
                    "-17,480,978,898 (당기순이익) 381,328,105 -3,210,944,230"
                ),
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (381_328_105, -3_210_944_230),
        )

    def test_year_only_header_is_not_net_income(self):
        text = "\n".join(
            [
                "손익계산서",
                "제12기 당기 2026년 1월 1일부터 2026년 6월 30일까지",
                "법인세비용",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual((net["current"], net["previous"]), ("", ""))

    def test_fuzzy_asset_total_label_uses_amounts_after_label(self):
        text = "\n".join(
            [
                "재무상태표 제7기 2025년 제6기 2024년",
                "자 AF 총 계 1,622,097,953 578,173,885 부채총계 739,485,920 36,690,180",
                "자본금 62,467,000 59,778,000",
            ]
        )
        assets = rows_by_item(text)["자산총계"]
        self.assertEqual(
            (assets["current"], assets["previous"]),
            (1_622_097_953, 578_173_885),
        )

    def test_explicit_asset_total_beats_combined_capital_and_balance_values(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계 1,270,797,907 2,486,923,233",
                "자본총계 부채및자본총계 24,164,824 1,270,797,907 1,391,901,886 2,486,923,233",
            ]
        )
        assets = rows_by_item(text)["자산총계"]
        self.assertEqual(
            (assets["current"], assets["previous"]),
            (1_270_797_907, 2_486_923_233),
        )

    def test_clean_combined_balance_total_crosschecks_damaged_asset_digit(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계 8,659,273,388 8,758,732,013",
                "부채및자본총계 8,559,273,388 8,758,732,013",
            ]
        )
        assets = rows_by_item(text)["자산총계"]
        self.assertEqual(
            (assets["current"], assets["previous"]),
            (8_559_273_388, 8_758_732_013),
        )

    def test_interleaved_capital_and_combined_totals_use_combined_columns(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계 50,328,302,761 49,359,027,949",
                "자본총계부채및자본총계 25,560,087,768 50,328,302,761 25,415,956,508 49,359,027,949",
            ]
        )
        assets = rows_by_item(text)["자산총계"]
        self.assertEqual(
            (assets["current"], assets["previous"]),
            (50_328_302_761, 49_359_027_949),
        )

    def test_interleaved_balance_totals_recover_liability_and_vertical_capital(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계 50,328,302,761 49,359,027,949",
                "자본금 우선주자본금 자본잉여금",
                "24,768,214,993",
                "4,209,332,000",
                "18,591,961,038",
                "23,943,071,441",
                "2,104,666,000",
                "20,711,372,629",
                "자본총계부채및자본총계 25,560,087,768 50,328,302,761 25,415,956,508 49,359,027,949",
            ]
        )
        rows = rows_by_item(text)
        self.assertEqual(
            (rows["부채총계"]["current"], rows["부채총계"]["previous"]),
            (24_768_214_993, 23_943_071_441),
        )
        self.assertEqual(
            (rows["자본금"]["current"], rows["자본금"]["previous"]),
            (4_209_332_000, 2_104_666_000),
        )

    def test_implausibly_small_comparison_asset_is_left_blank(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계 4,437,845,624 7,845,624",
            ]
        )
        assets = rows_by_item(text)["자산총계"]
        self.assertEqual((assets["current"], assets["previous"]), (4_437_845_624, ""))

    def test_single_period_asset_does_not_absorb_following_liability(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계 16,453,734,222",
                "부채",
                "유동부채 2,893,409,376",
                "자본금 100,000,000",
            ]
        )
        assets = rows_by_item(text)["자산총계"]
        self.assertEqual(
            (assets["current"], assets["previous"]),
            (16_453_734_222, ""),
        )

    def test_vertical_comparison_columns_recover_asset_total(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계",
                "제 8 기 (기준)",
                "3,750,000",
                "121,365,800",
                "9,979,195,205",
                "제 7 기 (비교)",
                "47,000,000",
                "2,183,911,123",
                "12,211,962,843",
                "2025/10/21 오후 2:00:32",
            ]
        )
        assets = rows_by_item(text)["자산총계"]
        self.assertEqual(
            (assets["current"], assets["previous"]),
            (9_979_195_205, 12_211_962_843),
        )

    def test_noncurrent_asset_zero_is_not_a_total(self):
        assets = rows_by_item("재무상태표\n자산총계\n비유동자산총계 0 0")["자산총계"]
        self.assertEqual((assets["current"], assets["previous"]), ("", ""))

    def test_single_period_capital_section_uses_plain_paid_in_capital_detail(self):
        text = "\n".join(
            [
                "요약 재무상태표",
                "Ⅰ.자본금 4,010,325,100",
                "자본금 10,325,100",
                "기타자본 4,000,000,000",
                "Ⅱ.자본잉여금 1,860,274,900",
            ]
        )
        capital = rows_by_item(text)["자본금"]
        self.assertEqual((capital["current"], capital["previous"]), (10_325_100, ""))

    def test_capital_aggregate_is_not_mixed_with_following_component_row(self):
        text = "\n".join(
            [
                "재무상태표",
                "자 본",
                "|. 자 본 a 283,750,000 283,750,000",
                "자본금 174,910,000 174,910,000",
                "우선주 자본금 108,840,000 108,840,000",
            ]
        )
        capital = rows_by_item(text)["자본금"]
        self.assertEqual(
            (capital["current"], capital["previous"]),
            (283_750_000, 283_750_000),
        )

    def test_trial_balance_ocr_capital_total_precedes_components(self):
        text = (
            "합계잔액시산표 차 본 2 150,909,500 150,909,500 "
            "자 본 금 110,611,000 110,611,000 "
            "우 선 주 자 본 금 40,298,500 40,298,500"
        )
        capital = rows_by_item(text)["자본금"]
        self.assertEqual(
            (capital["current"], capital["previous"]),
            (150_909_500, 150_909_500),
        )

    def test_flattened_share_classes_are_summed_for_each_period(self):
        text = (
            "재무상태표 자 본 L.A 본 # 9,654,800 9,654,800 "
            "자본금 5,025,100 5,025,100 "
            "우선주자본금 4,629,700 4,629,700 자본잉여금 9,081,505,626"
        )
        capital = rows_by_item(text)["자본금"]
        self.assertEqual(
            (capital["current"], capital["previous"]),
            (9_654_800, 9_654_800),
        )

    def test_last_plain_capital_before_preferred_is_common_component(self):
        text = (
            "재무상태표 자본금 223,948,000 216,950,000 "
            "자본금 103,500,000 103,500,000 "
            "우선주자본금 120,448,000 113,450,000 자본잉여금 5,000,000"
        )
        capital = rows_by_item(text)["자본금"]
        self.assertEqual(
            (capital["current"], capital["previous"]),
            (223_948_000, 216_950_000),
        )

    def test_trial_balance_misspellings_of_share_classes_are_repaired(self):
        text = (
            "표준재무제표증명 334 76,472,000 "
            "보롱주자본금 335 50,000,000 쿠선주자본금 336 26,472,000 "
            "자본잉여금 4,532,589,342"
        )
        capital = rows_by_item(text)["자본금"]
        self.assertEqual((capital["current"], capital["previous"]), (76_472_000, ""))

    def test_explicit_aggregate_capital_beats_damaged_common_component(self):
        text = (
            "재무상태표 1. 자본금 429,300,000 429,300,000 "
            "보통주 자본금 2,372,260,000 372,260,000 "
            "우선주 자본금 57,040,000 57,040,000 자본잉여금 2,838,267,000"
        )
        capital = rows_by_item(text)["자본금"]
        self.assertEqual(
            (capital["current"], capital["previous"]),
            (429_300_000, 429_300_000),
        )

    def test_year_headings_are_not_capital_amounts(self):
        text = "재무상태표 제7기 2025년 제6기 2024년 자본금 103,500,000 103,500,000"
        capital = rows_by_item(text)["자본금"]
        self.assertEqual(
            (capital["current"], capital["previous"]),
            (103_500_000, 103_500_000),
        )

    def test_trailing_income_statement_title_keeps_preceding_body(self):
        text = "\n".join(
            [
                "매출액 2,000,000,000",
                "영업손실 (1,133,640,585)",
                "법인세차감전이익 (1,079,002,045)",
                "법인세등 -",
                "당기순이익 (1,079,002,045)",
                "요약손익계산서(가결산)",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual((net["current"], net["previous"]), (-1_079_002_045, ""))

    def test_balance_summary_finds_previous_amount_before_loss_label(self):
        text = "\n".join(
            [
                "미처리결손금 1,912,185,271 1,793,462,692",
                "전기 : 368,625,303",
                "자본총계 1,589,819,625 1,708,542,204",
                "Ⅴ. 결손금",
                "( 당기순손실 )",
                "당기 : 118,722,579",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (-118_722_579, -368_625_303),
        )

    def test_ocr_damaged_period_marker_uses_explicit_loss_summary(self):
        text = "\n".join(
            [
                "미처리결손금 95,394,974 -145,516,268",
                "( 당기순손실 )",
                "담 기 ※ 240,911,242 원",
                "전기 : -90,664,545 원",
                "손익계산서",
                "영업손실 241,585,619 -91,808,273",
                "법인세차감전손실 240,911,242 -91,974,396",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (-240_911_242, -90_664_545),
        )

    def test_clean_loss_summary_beats_noisy_calculated_income(self):
        text = "\n".join(
            [
                "결손금 985,429,205 295,444,580",
                "( 당 키 순 솔 쉼 )",
                "당 기 ： 689,984,625",
                "전기 ： 295,444,580",
                "손익계산서",
                "영업손실 668,504,880 362,328,593",
                "영업외수익 475,390 70,256,529",
                "영업외비용 21,955,135 3,372,516",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual(
            (net["current"], net["previous"]),
            (-689_984_625, -295_444_580),
        )

    def test_missing_operating_total_is_not_fabricated_as_zero(self):
        text = "\n".join(
            [
                "손익계산서",
                "매출액 144,749,091 55,324,000",
                "매출총이익10. 판매비와관리비",
                "직원급여 355,553,367 95,024,859",
                "영업외수익 22,434,498 47,804,943",
            ]
        )
        operating = rows_by_item(text)["영업이익"]
        self.assertEqual((operating["current"], operating["previous"]), ("", ""))

    def test_standard_tax_form_revenue_uses_total_and_has_no_prior_period(self):
        text = "\n".join(
            [
                "국세청 표 준 재 무 제 표 2025년 12월 31일까지",
                "1. 메 출 역 |00| 10,908,009,689 TESTA -9,819,410",
                "제품매출 10,091,809,685 소모품비 46,186,198",
                "매출원가 10,088,414,689",
            ]
        )
        revenue = rows_by_item(text)["매출액"]
        self.assertEqual(
            (revenue["current"], revenue["previous"]),
            (10_908_009_689, ""),
        )

    def test_malformed_decimal_capital_is_left_blank(self):
        text = "표준재무제표\n보 동 주 자본금 1189.008.800\n자본잉여금 14,425,876,088"
        capital = rows_by_item(text)["자본금"]
        self.assertEqual((capital["current"], capital["previous"]), ("", ""))

    def test_malformed_liability_group_is_not_reported_as_truncated_number(self):
        text = "\n".join(
            [
                "재무상태표",
                "자산총계 20,000,000,000 10,000,000,000",
                "부채총계 11,006,041,39 4,035,482,742",
                "자본금 2,209,500,000 2,209,500,000",
            ]
        )
        liabilities = rows_by_item(text)["부채총계"]
        self.assertEqual((liabilities["current"], liabilities["previous"]), ("", 4_035_482_742))
        self.assertIn("자릿수 손상", liabilities["status"])

    def test_fuzzy_liability_total_label_is_repaired(self):
        text = "재무상태표\n부 채 & AH 4,623,012,021 1,353,610,096\n자본금 223,948,000 216,950,000"
        liabilities = rows_by_item(text)["부채총계"]
        self.assertEqual(
            (liabilities["current"], liabilities["previous"]),
            (4_623_012_021, 1_353_610_096),
        )

    def test_missing_liability_total_does_not_borrow_preceding_numbers(self):
        text = "\n".join(
            [
                "재무상태표",
                "기타 계정 29,776,842,509 2,569,648,294",
                "부채총계",
                "자본금 200,000,000 200,000,000",
            ]
        )
        liabilities = rows_by_item(text)["부채총계"]
        self.assertEqual((liabilities["current"], liabilities["previous"]), ("", ""))

    def test_liability_components_are_used_only_when_both_periods_are_clean(self):
        text = "\n".join(
            [
                "재무상태표",
                "유동부채 345,088,680 474,209,799",
                "비유동부채 4,277,923,341 879,400,297",
                "자본금 223,948,000 216,950,000",
            ]
        )
        liabilities = rows_by_item(text)["부채총계"]
        self.assertEqual(
            (liabilities["current"], liabilities["previous"]),
            (4_623_012_021, 1_353_610_096),
        )

    def test_explicit_liability_total_beats_small_component_discrepancy(self):
        text = "\n".join(
            [
                "재무상태표",
                "유동부채 63,611,074 63,727,617",
                "비유동부채 689,810 689,810",
                "부채총계 63,611,074 64,417,427",
            ]
        )
        liabilities = rows_by_item(text)["부채총계"]
        self.assertEqual(
            (liabilities["current"], liabilities["previous"]),
            (63_611_074, 64_417_427),
        )

    def test_reversed_spreadsheet_export_recovers_nearby_balance_totals(self):
        text = "\n".join(
            [
                "3,453,679,514 2,874,026,193",
                "863,859,889 1,165,247,918",
                "1,000,000,000 236,071",
                "1,863,859,889 1,165,483,989",
                "재무상태표",
                "부채총계",
                "자본",
                "비유동부채",
                "부채",
                "유동부채",
                "자산총계",
                "자본금 1,502,000,000 1,502,000,000",
                "3,453,679,514 2,874,026,193",
                "자본총계",
                "부채와자본총계",
            ]
        )
        rows = rows_by_item(text)
        self.assertEqual(
            (rows["자산총계"]["current"], rows["자산총계"]["previous"]),
            (3_453_679_514, 2_874_026_193),
        )
        self.assertEqual(
            (rows["부채총계"]["current"], rows["부채총계"]["previous"]),
            (1_863_859_889, 1_165_483_989),
        )

    def test_dense_vertical_liability_page_is_not_parsed_from_trailing_numbers(self):
        dense = (
            "부채총계본 자본금 자본잉여금 자본조정 기타포괄손익누계액 "
            + "장기차입금 퇴직급여충당금 " * 80
            + "5,000,000,000 533,460,149 -2,308,789,784 698,704,850"
        )
        liabilities = rows_by_item("재무상태표\n" + dense)["부채총계"]
        self.assertEqual((liabilities["current"], liabilities["previous"]), ("", ""))

    def test_dense_combined_balance_label_band_does_not_invent_totals(self):
        dense_labels = (
            "유동자산 비유동자산 자산총계 유동부채 비유동부채 부채총계 "
            "자본금 자본잉여금 자본조정 결손금 자본총계 부채및자본총계 "
            + "장기차입금 퇴직급여충당금 주식발행초과금 " * 8
        )
        text = "\n".join(
            [
                "재무상태표",
                dense_labels,
                "5,000,000,000 533,460,149 -2,308,789,784",
                "698,704,850 678,304,503 -673,870,027",
            ]
        )
        rows = rows_by_item(text)
        self.assertEqual((rows["자산총계"]["current"], rows["자산총계"]["previous"]), ("", ""))
        self.assertEqual((rows["부채총계"]["current"], rows["부채총계"]["previous"]), ("", ""))

    def test_negative_liability_from_balance_equation_is_never_reported(self):
        with (
            patch.object(MODULE, "find_balance_component_total", return_value=None),
            patch.object(
                MODULE,
                "find_combined_balance_total",
                return_value="부채및자본총계 100,000,000 90,000,000",
            ),
            patch.object(
                MODULE,
                "find_vertical_alias_segment",
                return_value="자본총계 200,000,000 80,000,000",
            ),
        ):
            liabilities = MODULE.extract_liabilities_total(["재무상태표"], "부채총계")
        assert liabilities is not None
        self.assertEqual((liabilities["current"], liabilities["previous"]), ("", 10_000_000))

    def test_reordered_current_liability_label_uses_clean_components(self):
        text = "\n".join(
            [
                "재무상태표",
                "채유 부채 2,344,308,963 2,155,337,677 비 유 동 부 A 5,810,471,203 4,810,471,203",
                "자본금 598,905,000 598,905,000",
            ]
        )
        liabilities = rows_by_item(text)["부채총계"]
        self.assertEqual(
            (liabilities["current"], liabilities["previous"]),
            (8_154_780_166, 6_965_808_880),
        )

    def test_zero_tax_pretax_beats_contaminated_standard_form_net_row(self):
        text = "\n".join(
            [
                "국세청 표준손익계산서 2025년 12월 31일까지",
                "법 인 셰 비 용 차 감 전 손익 2,206,641,098",
                "법 인 세 비 용 0",
                "임차료 191,591,787 ※. 망 기 순 손익 2,906,041,098",
            ]
        )
        net = rows_by_item(text)["당기순이익"]
        self.assertEqual((net["current"], net["previous"]), (2_206_641_098, 0))


if __name__ == "__main__":
    unittest.main()
