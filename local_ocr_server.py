from __future__ import annotations

import cgi
import html
import json
import re
import shutil
import os
import subprocess
import sys
import time
import threading
import uuid
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


ROOT = Path(__file__).resolve().parent
OUTPUTS = ROOT / "outputs"
TARGET_ITEMS = [
    ("자산총계", ["자산총계", "자 산 총 계"]),
    ("부채총계", ["부채총계", "부 채 총 계", "채총계", "채 총 계"]),
    ("자본총계", []),
    ("자본금", ["자본금", "자 본 금"]),
    ("매출액", ["매출액", "매 출 액", "출액", "출 액", "영업수익", "수수료", "제품매출"]),
    ("영업이익", ["영업이익", "영업 이익", "영업손실", "영업 손실", "영업손"]),
    ("당기순이익", ["당기순이익", "당기순손실", "당기 순손실", "당기순손", "분기순이익", "분기순손실"]),
]
REPORT_FIELDS = [
    ("자산총계", ("current",)),
    ("부채총계", ("current",)),
    ("자본금", ("current",)),
    ("매출액", ("current", "previous")),
    ("영업이익", ("current", "previous")),
    ("당기순이익", ("current", "previous")),
]
AMOUNT_PATTERN = re.compile(
    r"\(?-?(?:"
    r"\d{1,3}(?:,\s?\d{3})+(?:\.\d+)?|"
    r"\d{1,3}(?:\.\s?\d{3})+|"
    r"\d{4,}(?:\.\d+)?|"
    r"(?<!\d)0(?:\.\d+)?(?!\d)"
    r")\)?"
)
OCR_PROFILES = [
    {"name": "balanced_psm6_300", "lang": "kor+eng", "psm": "6", "oversample": "300"},
    {"name": "table_psm4_400", "lang": "kor+eng", "psm": "4", "oversample": "400"},
    {"name": "sparse_psm11_400", "lang": "kor+eng", "psm": "11", "oversample": "400"},
    {"name": "korean_only_psm6_400", "lang": "kor", "psm": "6", "oversample": "400"},
    {"name": "compat_default", "lang": "kor+eng", "psm": "", "oversample": ""},
]
FAST_PROFILE = OCR_PROFILES[0]
MAX_PARALLEL_FILES = max(1, min(3, (os.cpu_count() or 2) // 2))
JOBS: dict[str, dict] = {}
JOBS_LOCK = threading.Lock()


class Handler(SimpleHTTPRequestHandler):
    def end_headers(self):
        if self.path in ("/", "/server", "/server.html") or self.path.endswith(".html"):
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
        super().end_headers()

    def do_GET(self):
        if self.path.startswith("/api/progress?"):
            job_id = self.path.partition("jobId=")[2].split("&", 1)[0]
            with JOBS_LOCK:
                job = dict(JOBS.get(job_id, {}))
            self.send_json(200 if job else 404, {"ok": bool(job), **job})
            return
        if self.path == "/api/status":
            self.send_json(
                200,
                {
                    "ok": True,
                    "serverUrl": f"http://127.0.0.1:{self.server.server_port}/server",
                    "root": str(ROOT),
                    "outputs": str(OUTPUTS),
                    "ocrAvailable": bool(find_ocrmypdf_command()),
                    "ocrProfiles": [profile["name"] for profile in OCR_PROFILES],
                    "parallelFiles": MAX_PARALLEL_FILES,
                },
            )
            return
        if self.path in ("/", "/server"):
            self.path = "/server.html"
        return super().do_GET()

    def do_POST(self):
        if self.path != "/api/process":
            self.send_error(404, "Not found")
            return

        try:
            result = process_upload(self)
            self.send_json(200, result)
        except Exception as exc:
            self.send_json(500, {"ok": False, "error": str(exc)})

    def send_json(self, status: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def process_upload(request: Handler) -> dict:
    form = cgi.FieldStorage(
        fp=request.rfile,
        headers=request.headers,
        environ={
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": request.headers.get("Content-Type", ""),
        },
    )
    if "files" in form:
        fields = form["files"]
    elif "pdfs" in form:
        fields = form["pdfs"]
    else:
        fields = []
    if not isinstance(fields, list):
        fields = [fields]

    mode = form.getfirst("mode", "precise")
    if mode not in ("fast", "precise"):
        mode = "precise"
    job_id = form.getfirst("jobId", "") or uuid.uuid4().hex

    # A timestamp alone can collide when two browser requests start during the
    # same second. Keep the readable timestamp and add a short random suffix.
    session_dir = OUTPUTS / f"{time.strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
    upload_dir = session_dir / "uploads"
    final_dir = session_dir / "final"
    upload_dir.mkdir(parents=True, exist_ok=True)
    final_dir.mkdir(parents=True, exist_ok=True)

    upload_items = []
    ocrmypdf_command = find_ocrmypdf_command()
    update_job(job_id, total=len(fields), completed=0, current="파일 업로드 완료", status="running")

    for index, field in enumerate(fields, start=1):
        if not getattr(field, "filename", ""):
            continue

        original_name = safe_filename(field.filename)
        # Prefix the stored name so two equally named files from different
        # folders cannot overwrite one another in the same upload.
        input_path = upload_dir / f"{index:03d}_{original_name}"
        with input_path.open("wb") as file:
            shutil.copyfileobj(field.file, file)
        upload_items.append((index, original_name, input_path))

    update_job(job_id, total=len(upload_items), completed=0, current="파일 업로드 완료", status="running")

    documents = []
    try:
        with ThreadPoolExecutor(max_workers=min(MAX_PARALLEL_FILES, max(1, len(upload_items)))) as executor:
            futures = {
                executor.submit(process_one_file, item, final_dir, ocrmypdf_command, mode): item
                for item in upload_items
            }
            for completed, future in enumerate(as_completed(futures), start=1):
                item = futures[future]
                try:
                    document = future.result()
                    current = item[1]
                except Exception as exc:
                    # One damaged PDF must not discard every successfully
                    # processed company in a multi-file batch.
                    document = failed_document(item, exc)
                    current = f"{item[1]} (실패)"
                documents.append(document)
                update_job(job_id, completed=completed, current=current)
    finally:
        shutil.rmtree(upload_dir, ignore_errors=True)

    documents.sort(key=lambda document: document["index"])
    documents = merge_company_documents(documents)

    if not documents:
        raise ValueError("처리할 파일을 찾지 못했습니다.")

    excel_path = session_dir / "재무제표_추출.xlsx"
    build_workbook(excel_path, documents)
    update_job(job_id, completed=len(upload_items), current="완료", status="done")

    return {
        "ok": True,
        "ocrAvailable": bool(ocrmypdf_command),
        "ocrCommand": " ".join(ocrmypdf_command) if ocrmypdf_command else "",
        "excelUrl": to_url(excel_path),
        "outputDirUrl": to_url(session_dir),
        "outputDirPath": str(session_dir),
        "documents": documents,
        "mode": mode,
        "completedFiles": len(upload_items),
        "totalFiles": len(upload_items),
        "parallelFiles": min(MAX_PARALLEL_FILES, max(1, len(upload_items))),
        "installHint": install_hint() if not ocrmypdf_command else "",
    }


def update_job(job_id: str, **values) -> None:
    with JOBS_LOCK:
        JOBS.setdefault(job_id, {}).update(values)


def process_one_file(item, final_dir: Path, ocrmypdf_command: list[str], mode: str) -> dict:
    index, original_name, input_path = item
    base_name = Path(original_name).stem
    storage_base_name = f"{index:03d}_{base_name}"
    suffix = input_path.suffix.lower()
    final_pdf_path = final_dir / f"{storage_base_name}_최종.pdf"
    final_text_path = final_dir / f"{storage_base_name}_텍스트.txt"
    ocr_status = "OCRmyPDF 미설치"
    ocr_error = ""
    text = ""

    if suffix in (".xlsx", ".xlsm"):
        ocr_status = "엑셀 직접 추출"
        text = extract_excel_text(input_path)
    elif suffix == ".pdf":
        embedded_text = extract_text(input_path)
        if has_meaningful_text(embedded_text):
            text = embedded_text
            shutil.copyfile(input_path, final_pdf_path)
            ocr_status = "텍스트 PDF 감지 · OCR 생략"
        elif ocrmypdf_command:
            if mode == "precise":
                result = run_best_ocr(
                    ocrmypdf_command,
                    input_path,
                    final_dir,
                    storage_base_name,
                    final_pdf_path,
                    final_text_path,
                )
            else:
                result = run_fast_ocr(ocrmypdf_command, input_path, final_pdf_path, final_text_path)
            if result["ok"]:
                text = result["text"]
                ocr_status = f"{('정밀' if mode == 'precise' else '빠른')} OCR 완료: {result['profile']}"
            else:
                text = embedded_text
                ocr_status = "OCR 실패"
                ocr_error = result["error"]
                shutil.copyfile(input_path, final_pdf_path)
        else:
            text = embedded_text
    else:
        ocr_status = "지원하지 않는 파일"
        ocr_error = "PDF, XLSX, XLSM 파일만 처리할 수 있습니다."

    if text:
        final_text_path.write_text(text, encoding="utf-8")
    extracted_rows = extract_financial_rows(text)
    return {
        "index": index,
        "fileName": original_name,
        "sheetName": unique_sheet_name(base_name, index),
        "ocrStatus": ocr_status,
        "ocrError": ocr_error,
        "ocrUrl": to_url(final_pdf_path) if final_pdf_path.exists() else "",
        "textUrl": to_url(final_text_path) if final_text_path.exists() else "",
        "sourceText": text[:6000],
        "rows": report_rows(extracted_rows),
    }


def failed_document(item, error: Exception) -> dict:
    """Represent an isolated file failure without aborting the whole batch."""
    index, original_name, _input_path = item
    base_name = Path(original_name).stem
    return {
        "index": index,
        "fileName": original_name,
        "sheetName": unique_sheet_name(base_name, index),
        "ocrStatus": "처리 실패",
        "ocrError": str(error),
        "ocrUrl": "",
        "textUrl": "",
        "sourceText": "",
        "rows": report_rows([]),
    }


def report_rows(rows: list[dict]) -> list[dict]:
    values = {row["item"]: row for row in rows}
    result = []
    for label, periods in REPORT_FIELDS:
        row = dict(values.get(label, {"item": label, "current": "", "previous": "", "status": "미발견"}))
        row["includePrevious"] = "previous" in periods
        if "previous" not in periods:
            row["previous"] = ""
        result.append(row)
    return result


def statement_kind(file_name: str, source_text: str = "") -> str:
    stem = Path(file_name).stem
    compact = re.sub(r"\s+", "", stem)
    has_bs_token = bool(re.search(r"(?:^|[\s_.#-])BS(?:$|[\s_.#-])", stem, flags=re.IGNORECASE))
    has_pl_token = bool(re.search(r"(?:^|[\s_.#-])PL(?:$|[\s_.#-])", stem, flags=re.IGNORECASE))
    text_sample = re.sub(r"\s+", "", source_text[:4000]).casefold()
    if (
        has_bs_token
        or
        "재무상태표" in compact
        or "대차대조표" in compact
        or "balancesheet" in compact.casefold()
        or "consolidatedbalancesheet" in text_sample
    ):
        return "balance"
    if (
        has_pl_token
        or
        "손익계산서" in compact
        or "손익게산서" in compact
        or "포괄손익계산서" in compact
        or "incomestatement" in compact.casefold()
        or "consolidatedincomestatement" in text_sample
    ):
        return "income"
    return ""


def statement_period_key(file_name: str) -> str:
    stem = Path(file_name).stem
    patterns = (
        (r"(20\d{2})\s*[._년/-]\s*(0?[1-9]|1[0-2])\s*[._월/-]\s*([0-3]?\d)", lambda m: f"{m[1]}{int(m[2]):02d}{int(m[3]):02d}"),
        (r"(?<!\d)(\d{2})\s*년\s*(0?[1-9]|1[0-2])\s*월", lambda m: f"20{m[1]}{int(m[2]):02d}"),
        (r"(20\d{2})\s*[._년/-]?\s*([1-4])\s*분기", lambda m: f"{m[1]}Q{m[2]}"),
        (r"(20\d{2})\s*년", lambda m: m[1]),
        (r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])([0-3]\d)(?!\d)", lambda m: f"{m[1]}{m[2]}{m[3]}"),
    )
    matches = []
    for pattern, formatter in patterns:
        for match in re.finditer(pattern, stem, flags=re.IGNORECASE):
            matches.append((match.start(), formatter(match)))
    return min(matches, default=(0, ""), key=lambda item: item[0])[1]


def company_key_from_file(file_name: str) -> str:
    stem = Path(file_name).stem.casefold()
    stem = re.sub(r"^\s*\d+(?:[-.]\d+)*[.)]?\s*", "", stem)
    stem = re.sub(r"\(\s*주\s*\)|㈜|주식회사", "", stem)
    stem = re.sub(
        r"재무상태표|대차대조표|포괄손익계산서|손익계산서|손익게산서|재무제표|가결산|확정|별도|연결|감사보고서",
        " ",
        stem,
    )
    stem = re.sub(r"(?<![a-z0-9])(?:bs|pl)(?![a-z0-9])", " ", stem, flags=re.IGNORECASE)
    stem = re.sub(r"20\d{6}|\d{2,4}\s*년|\d{1,2}\s*월|\d{1,2}\s*일|[1-4]\s*분기|[1-4]\s*q", " ", stem)
    stem = re.sub(r"\d+(?:[._/-]\d+)+|\d+", " ", stem)
    return re.sub(r"[^0-9a-z가-힣]+", "", stem)


def periods_compatible(left: dict, right: dict) -> bool:
    left_period = statement_period_key(left["fileName"])
    right_period = statement_period_key(right["fileName"])
    return not left_period or not right_period or left_period == right_period


def merged_statement_document(balance: dict, income: dict) -> dict:
    balance_rows = {row["item"]: row for row in balance["rows"]}
    income_rows = {row["item"]: row for row in income["rows"]}
    merged_rows = []
    for label, periods in REPORT_FIELDS:
        primary = balance_rows if label in ("자산총계", "부채총계", "자본금") else income_rows
        secondary = income_rows if primary is balance_rows else balance_rows
        row = dict(primary.get(label) or secondary.get(label) or {"item": label, "current": "", "previous": "", "status": "미발견"})
        row["includePrevious"] = "previous" in periods
        if "previous" not in periods:
            row["previous"] = ""
        merged_rows.append(row)

    company_name = company_name_from_file(balance["fileName"])
    source_files = []
    for document in (balance, income):
        source_files.append({
            "fileName": document["fileName"],
            "ocrUrl": document.get("ocrUrl", ""),
            "textUrl": document.get("textUrl", ""),
        })
    return {
        "index": min(balance["index"], income["index"]),
        "fileName": f"{company_name} (재무상태표 + 손익계산서)",
        "companyName": company_name,
        "sheetName": balance["sheetName"],
        "ocrStatus": "동일 회사 재무상태표·손익계산서 병합",
        "ocrError": "\n".join(filter(None, (balance.get("ocrError"), income.get("ocrError")))),
        "ocrUrl": "",
        "textUrl": "",
        "sourceFiles": source_files,
        "sourceText": "\n\n".join(filter(None, (balance.get("sourceText"), income.get("sourceText")))),
        "rows": merged_rows,
    }


def merge_company_documents(documents: list[dict]) -> list[dict]:
    result = []
    grouped: dict[str, list[dict]] = {}
    for document in documents:
        kind = statement_kind(document["fileName"], document.get("sourceText", ""))
        company_key = company_key_from_file(document["fileName"])
        if kind and len(company_key) >= 2:
            grouped.setdefault(company_key, []).append(document)
        else:
            result.append(document)

    for group in grouped.values():
        balances = [
            document for document in group
            if statement_kind(document["fileName"], document.get("sourceText", "")) == "balance"
        ]
        incomes = [
            document for document in group
            if statement_kind(document["fileName"], document.get("sourceText", "")) == "income"
        ]
        used_income_ids = set()
        for balance in balances:
            income = next(
                (
                    candidate
                    for candidate in incomes
                    if id(candidate) not in used_income_ids and periods_compatible(balance, candidate)
                ),
                None,
            )
            if income:
                used_income_ids.add(id(income))
                result.append(merged_statement_document(balance, income))
            else:
                result.append(balance)
        result.extend(income for income in incomes if id(income) not in used_income_ids)

    result.sort(key=lambda document: document["index"])
    return result


def has_meaningful_text(text: str) -> bool:
    content = re.sub(r"--- \d+페이지 ---|\s+", "", text)
    readable = re.findall(r"[가-힣A-Za-z0-9]", content)
    nonempty_lines = [line.strip() for line in text.splitlines() if line.strip() and not line.startswith("--- ")]
    single_character_ratio = sum(len(re.sub(r"\s+", "", line)) <= 1 for line in nonempty_lines) / max(1, len(nonempty_lines))
    amount_count = len(AMOUNT_PATTERN.findall(text))
    return len(readable) >= 80 and single_character_ratio < 0.55 and amount_count >= 2


def extraction_quality(text: str) -> int:
    rows = extract_financial_rows(text)
    values = {row["item"]: row for row in rows}
    complete = sum(row.get("current") != "" and row.get("previous") != "" for row in rows if row["item"] != "자본총계")
    score = complete * 100
    assets = values.get("자산총계", {}).get("current")
    liabilities = values.get("부채총계", {}).get("current")
    if assets not in (None, "") and liabilities not in (None, ""):
        score += 30 if assets != liabilities and 0 <= liabilities <= assets * 2 else -150
    revenue_source = normalize(values.get("매출액", {}).get("sourceLine", ""))
    if any(term in revenue_source for term in ("매출원가", "제품매출원가", "매출총이익")):
        score -= 150
    return score


def needs_precise_retry(text: str) -> bool:
    """Report extraction risk for diagnostics; mode selection stays user-controlled."""
    rows = {row["item"]: row for row in extract_financial_rows(text)}
    required = ("자산총계", "부채총계", "자본금", "매출액", "영업이익", "당기순이익")
    missing = sum(rows.get(item, {}).get("current") in ("", None) for item in required)
    suspicious_period = any(
        isinstance(row.get("current"), int)
        and isinstance(row.get("previous"), int)
        and row["current"] != 0
        and row["previous"] != 0
        and abs(row["previous"]) * 1000 < abs(row["current"])
        for row in rows.values()
    )
    assets = rows.get("자산총계", {}).get("current")
    liabilities = rows.get("부채총계", {}).get("current")
    capital = rows.get("자본금", {}).get("current")
    impossible_balance = (
        isinstance(assets, int)
        and (
            isinstance(liabilities, int) and (liabilities < 0 or liabilities > assets * 2)
            or isinstance(capital, int) and capital > assets * 2
        )
    )
    return extraction_quality(text) < 600 or missing >= 2 or suspicious_period or impossible_balance


def find_ocrmypdf_command() -> list[str]:
    configured = read_configured_ocr_command()
    candidates: list[list[str]] = []
    if configured:
        candidates.append(configured)

    found = shutil.which("ocrmypdf")
    if found:
        candidates.append([found])

    for path in common_ocrmypdf_paths():
        candidates.append([str(path)])

    for python_command in python_candidates():
        candidates.append([*python_command, "-m", "ocrmypdf"])

    seen = set()
    unique_candidates = []
    for candidate in candidates:
        key = tuple(candidate)
        if key not in seen:
            seen.add(key)
            unique_candidates.append(candidate)

    for candidate in unique_candidates:
        if command_works([*candidate, "--version"]):
            return candidate
    return []


def run_best_ocr(
    base_command: list[str],
    input_path: Path,
    ocr_dir: Path,
    base_name: str,
    final_pdf_path: Path,
    final_sidecar_path: Path,
) -> dict:
    attempts = []
    command_logs = []
    with tempfile.TemporaryDirectory(prefix="ocr_compare_", dir=ocr_dir) as temp_name:
        temp_dir = Path(temp_name)
        for profile in OCR_PROFILES:
            candidate_pdf = temp_dir / f"{base_name}_{profile['name']}.pdf"
            candidate_text = temp_dir / f"{base_name}_{profile['name']}.txt"
            command = build_ocr_command(base_command, input_path, candidate_pdf, candidate_text, profile)
            command_logs.append(f"[{profile['name']}]\n{format_command(command)}\n")
            try:
                completed = subprocess.run(command, capture_output=True, text=True, timeout=1200)
            except Exception as exc:
                attempts.append({"ok": False, "profile": profile["name"], "score": 0, "error": str(exc)})
                continue
            output = "\n".join(part for part in [completed.stdout, completed.stderr] if part)
            if completed.returncode != 0 or not candidate_pdf.exists():
                attempts.append({"ok": False, "profile": profile["name"], "score": 0, "error": output.strip()})
                continue
            text = candidate_text.read_text(encoding="utf-8", errors="ignore") if candidate_text.exists() else extract_text(candidate_pdf)
            attempts.append({"ok": True, "profile": profile["name"], "score": score_ocr_text(text), "text": text, "pdf_path": candidate_pdf, "error": output.strip()})

        successful = [attempt for attempt in attempts if attempt["ok"]]
        if not successful:
            return {"ok": False, "profile": "", "score": 0, "text": "", "command_log": "\n".join(command_logs), "error": "\n\n".join(f"[{a['profile']}]\n{a.get('error', '')}" for a in attempts)}
        best = max(successful, key=lambda attempt: attempt["score"])
        best_text = repair_accounting_spacing(best["text"])
        shutil.copyfile(best["pdf_path"], final_pdf_path)
        final_sidecar_path.write_text(best_text, encoding="utf-8")
        return {"ok": True, "profile": best["profile"], "score": best["score"], "text": best_text, "command_log": "\n".join(command_logs), "error": ""}


def run_fast_ocr(base_command: list[str], input_path: Path, final_pdf_path: Path, final_sidecar_path: Path) -> dict:
    command = build_ocr_command(base_command, input_path, final_pdf_path, final_sidecar_path, FAST_PROFILE)
    try:
        completed = subprocess.run(command, capture_output=True, text=True, timeout=1200)
    except Exception as exc:
        return {"ok": False, "profile": FAST_PROFILE["name"], "text": "", "error": str(exc)}
    output = "\n".join(part for part in (completed.stdout, completed.stderr) if part).strip()
    if completed.returncode != 0 or not final_pdf_path.exists():
        final_pdf_path.unlink(missing_ok=True)
        final_sidecar_path.unlink(missing_ok=True)
        return {"ok": False, "profile": FAST_PROFILE["name"], "text": "", "error": output}
    text = final_sidecar_path.read_text(encoding="utf-8", errors="ignore") if final_sidecar_path.exists() else extract_text(final_pdf_path)
    text = repair_accounting_spacing(text)
    final_sidecar_path.write_text(text, encoding="utf-8")
    return {"ok": True, "profile": FAST_PROFILE["name"], "text": text, "error": ""}


def score_ocr_text(text: str) -> int:
    rows = extract_financial_rows(text)
    found_count = sum(1 for row in rows if row.get("current") != "" or row.get("previous") != "")
    source_count = sum(1 for row in rows if row.get("sourceLine"))
    amount_count = len(AMOUNT_PATTERN.findall(text))
    hangul_count = len(re.findall(r"[가-힣]", text))
    target_bonus = sum(1 for keyword in ("자산", "부채", "자본", "매출", "영업", "당기", "순", "손익") if keyword in text)
    return (
        extraction_quality(text) * 1000
        + found_count * 10000
        + source_count * 1000
        + target_bonus * 250
        + min(amount_count, 400) * 5
        + min(hangul_count, 3000)
    )


def build_ocr_command(base_command: list[str], input_path: Path, output_path: Path, sidecar_path: Path, profile: dict) -> list[str]:
    command = [
        *base_command,
        "--force-ocr",
        "--invalidate-digital-signatures",
        "--deskew",
        "--rotate-pages",
        "--pdf-renderer",
        "sandwich",
        "--sidecar",
        str(sidecar_path),
        "--optimize",
        "0",
        "-l",
        profile["lang"],
    ]
    if profile.get("psm"):
        command.extend(["--tesseract-pagesegmode", profile["psm"]])
    if profile.get("oversample"):
        command.extend(["--oversample", profile["oversample"]])
    if shutil.which("unpaper"):
        command.append("--clean")
    command.extend([str(input_path), str(output_path)])
    return command


def read_configured_ocr_command() -> list[str]:
    config_path = ROOT / "ocrmypdf_path.txt"
    if not config_path.exists():
        return []
    raw = config_path.read_text(encoding="utf-8").strip().strip('"')
    if not raw:
        return []
    if raw.lower().endswith(".exe") or Path(raw).exists():
        return [raw]
    return raw.split()


def common_ocrmypdf_paths() -> list[Path]:
    roots = [
        Path(os.environ.get("USERPROFILE", "")) / ".local" / "bin",
        Path(os.environ.get("USERPROFILE", "")) / "pipx",
        Path(os.environ.get("APPDATA", "")) / "Python",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Python",
        Path(os.environ.get("PROGRAMFILES", "")),
        Path(os.environ.get("PROGRAMFILES(X86)", "")),
    ]
    paths = []
    for root in roots:
        if root.exists():
            paths.extend(root.glob("**/Scripts/ocrmypdf.exe"))
            paths.extend(root.glob("**/ocrmypdf.exe"))
    return paths[:30]


def python_candidates() -> list[list[str]]:
    candidates = []
    for command in ("python", "python3", "py"):
        if shutil.which(command):
            candidates.append([command])
    local_python = Path(os.environ.get("LOCALAPPDATA", "")) / "Python"
    if local_python.exists():
        for python_exe in local_python.glob("pythoncore-*-64/python.exe"):
            candidates.append([str(python_exe)])
        for python_exe in local_python.glob("Python*/python.exe"):
            candidates.append([str(python_exe)])
    candidates.append([sys.executable])
    return candidates


def command_works(command: list[str]) -> bool:
    try:
        completed = subprocess.run(command, capture_output=True, text=True, timeout=20)
        return completed.returncode == 0
    except Exception:
        return False


def format_command(command: list[str]) -> str:
    return " ".join(quote_arg(part) for part in command)


def quote_arg(value: str) -> str:
    if re.search(r"\s|[()&]", value):
        return f'"{value}"'
    return value


def extract_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    chunks = []
    for page_index, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        chunks.append(f"--- {page_index}페이지 ---\n{text}")
    return "\n\n".join(chunks)


def extract_excel_text(excel_path: Path) -> str:
    workbook = load_workbook(excel_path, data_only=True, read_only=True)
    chunks = []
    try:
        for sheet in workbook.worksheets:
            chunks.append(f"--- 시트: {sheet.title} ---")
            for row in sheet.iter_rows(values_only=True):
                values = [format_excel_cell(value) for value in row]
                while values and values[-1] == "":
                    values.pop()
                if any(values):
                    chunks.append(" ".join(value for value in values if value != ""))
    finally:
        workbook.close()
    return "\n".join(chunks)


def format_excel_cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return f"{value:,}"
    if isinstance(value, float):
        if value.is_integer():
            return f"{int(value):,}"
        return str(value)
    return str(value).strip()


def extract_financial_rows(text: str) -> list[dict]:
    text = repair_accounting_spacing(normalize_ocr_numbers(text))
    base_lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]
    base_lines = [line for line in base_lines if line]
    base_lines = prefer_separate_statement_scope(base_lines)
    base_lines.extend(english_financial_candidates(base_lines))
    lines = add_continuation_candidates(base_lines)
    rows = []

    for label, aliases in TARGET_ITEMS:
        if label == "자산총계":
            assets_row = extract_assets_total(lines, label)
            if assets_row:
                rows.append(assets_row)
            else:
                rows.append({"item": label, "current": "", "previous": "", "status": "미발견"})
            continue
        if label == "부채총계":
            liabilities_row = extract_liabilities_total(lines, label)
            if liabilities_row:
                rows.append(liabilities_row)
            else:
                rows.append({"item": label, "current": "", "previous": "", "status": "미발견"})
            continue
        if label == "자본총계":
            values = {row["item"]: row for row in rows}
            assets = values.get("자산총계", {})
            liabilities = values.get("부채총계", {})
            current = subtract_values(assets.get("current"), liabilities.get("current"))
            previous = subtract_values(assets.get("previous"), liabilities.get("previous"))
            rows.append({
                "item": label,
                "current": current if current is not None else "",
                "previous": previous if previous is not None else "",
                "status": "자동 계산" if current is not None or previous is not None else "계산 불가",
                "sourceLine": "자산총계 - 부채총계",
            })
            continue
        if label == "자본금":
            capital_row = extract_capital(lines, label)
            if capital_row:
                rows.append(capital_row)
            else:
                rows.append({"item": label, "current": "", "previous": "", "status": "미발견"})
            continue
        if label == "영업이익":
            operating_row = extract_operating_profit(lines, label)
            if operating_row:
                rows.append(operating_row)
            else:
                rows.append({"item": label, "current": "", "previous": "", "status": "미발견"})
            continue
        if label == "당기순이익":
            net_income_row = extract_net_income(lines, label)
            if net_income_row:
                rows.append(net_income_row)
            else:
                rows.append({"item": label, "current": "", "previous": "", "status": "미발견"})
            continue
        if label == "매출액":
            revenue_row = extract_revenue(base_lines, label)
            if revenue_row:
                rows.append(revenue_row)
            else:
                rows.append({"item": label, "current": "", "previous": "", "status": "미발견"})
            continue
        match = find_best_line(lines, aliases, label)
        if not match:
            rows.append({"item": label, "current": "", "previous": "", "status": "미발견"})
            continue

        compact_match = normalize(match)
        is_loss = "손실" in compact_match or label in ("영업이익", "당기순이익") and "손" in compact_match
        numbers = extract_amounts(match)
        if is_loss:
            numbers = [-abs(num) for num in numbers]

        current, previous = pick_current_previous(numbers)
        rows.append(
            {
                "item": label,
                "current": current if current is not None else "",
                "previous": previous if previous is not None else "",
                "status": "확인 필요" if current is not None or previous is not None else "금액 없음",
                "sourceLine": match,
            }
        )

    return rows


def normalize_ocr_numbers(text: str) -> str:
    normalized = re.sub(r"[ \t]*,[ \t]*", ",", text)
    normalized = re.sub(
        r"(?<!\d)\d{1,3}(?:[,.]\d{3}){2,}(?!\d)",
        lambda match: match.group(0).replace(".", ","),
        normalized,
    )
    normalized = re.sub(r"(?<=\d)\.,(?=\d{3})", ",", normalized)
    for _ in range(2):
        normalized = re.sub(
            r"(?<![\d,])(\d{1,3})(\d{3})(?=,\d{3}(?!\d))",
            r"\1,\2",
            normalized,
        )
    normalized = re.sub(
        r"(?<![\d,])(\d{1,3})[ \t]+(\d{3})(?=,\d{3}(?!\d))",
        r"\1,\2",
        normalized,
    )
    def merge_split_six_digit_groups(match: re.Match) -> str:
        # If a complete current-period amount already appears on this row, two
        # six-digit groups after it are usually one damaged previous-period
        # amount (``77,701,095 7,426 698,918``).  Without that anchor they are
        # two legitimate period columns (``674,377 236,625``).
        line_start = match.string.rfind("\n", 0, match.start()) + 1
        prefix = match.string[line_start : match.start()]
        if AMOUNT_PATTERN.search(prefix):
            return f"{match.group(1)},{match.group(2)}"
        return match.group(0)

    normalized = re.sub(
        r"(?<![\d,])(\d{1,3},\d{3})[ \t]+(\d{3},\d{3})(?![\d,])",
        merge_split_six_digit_groups,
        normalized,
    )
    normalized = re.sub(
        r"(?<!\d)-[ \t]+(?=\d{1,3}(?:,\d{3})+)",
        "-",
        normalized,
    )
    normalized = re.sub(r"\([ \t]+(?=-?\d)", "(", normalized)
    normalized = re.sub(r"(?<=\d)[ \t]+\)", ")", normalized)
    def merge_trailing_thousands(match: re.Match) -> str:
        # A detached final thousands group is repaired only when another full
        # amount already precedes it on the same row.  This covers OCR such as
        # ``4,758,539,687 3,814,846 887`` while preserving a legitimate small
        # comparison value such as ``48,232 660``.
        line_start = match.string.rfind("\n", 0, match.start()) + 1
        prefix = match.string[line_start : match.start()]
        if AMOUNT_PATTERN.search(prefix):
            return f"{match.group(1)},{match.group(2)}"
        return match.group(0)

    normalized = re.sub(
        r"(\d{1,3}(?:,\d{3})+)[ \t]+(\d{3})(?![\d,])",
        merge_trailing_thousands,
        normalized,
    )
    return normalized


def subtract_values(left, right) -> int | float | None:
    if left in (None, "") or right in (None, ""):
        return None
    return left - right


def repair_accounting_spacing(text: str) -> str:
    terms = [
        "자산총계",
        "부채총계",
        "자본총계",
        "부채와자본총계",
        "부채및자본총계",
        "자본금",
        "매출액",
        "매출원가",
        "매출총이익",
        "영업이익",
        "영업손실",
        "당기순이익",
        "당기순손실",
        "법인세차감전손익",
        "법인세차감전손실",
        "법인세비용차감전손익",
        "법인세비용차감전손실",
        "수수료",
        "제품매출",
        "영업수익",
    ]
    repaired = text
    fuzzy_terms = (
        (r"자\s*[몬븐본봄]\s*[금급]", "자본금"),
        (r"당\s*[기키]\s*순\s*[손솔]\s*[실쉼]", "당기순손실"),
        (r"영\s*업\s*손\s*[실쉼]", "영업손실"),
        (r"영\s*업\s*[드도]\s*실", "영업손실"),
        (r"염\s*업\s*손\s*실", "영업손실"),
        (r"당\s*[기키]\s*순\s*[손=]\s*[실a4]", "당기순손실"),
        (r"당\s*[기키]\s*[순=]\s*이\s*익", "당기순이익"),
        (r"(?:부\s*)?채\s*[총촉]\s*[계A]", "부채총계"),
        (r"부\s*채\s*&\s*A[HIl1]", "부채총계"),
        (r"자\s*산\s*[총춤]\s*[계A]", "자산총계"),
        (r"자\s*A\s*F\s*[총춤]\s*[계A]", "자산총계"),
        (r"유\s*[동통]\s*[.ㆍ]?\s*자\s*산", "유동자산"),
        (r"비\s*유\s*[동톨]\s*[.ㆍ]?\s*자\s*산", "비유동자산"),
        (r"유\s*[동통]\s*[.ㆍ]?\s*부\s*[채A]", "유동부채"),
        (r"채\s*유\s*부\s*[채A]", "유동부채"),
        (r"비\s*유\s*[동톨]\s*[.ㆍ]?\s*부\s*[채A]", "비유동부채"),
        (r"자\s*(?:본|=)\s*[금3aACl]+", "자본금"),
        # Some dense trial-balance exports OCR ``자본금`` as ``차 본 2``.
        # Repairing that aggregate label lets us keep the total row instead of
        # mistaking the following common/preferred-stock components for periods.
        (r"차\s*본\s*2", "자본금"),
        (r"(?:At|Xt)\s*본\s*금", "자본금"),
        (r"보\s*[통롱동]\s*주\s*[AX자차]?\s*본\s*[금급]", "보통주 자본금"),
        (r"[우쿠]\s*선\s*주\s*[AX자차]?\s*본\s*[금급]", "우선주 자본금"),
        (r"(?:매|[oO0]f)\s*[출줄술]\s*(?:액|[oO0]f)", "매출액"),
        (r"메\s*출\s*[액역]", "매출액"),
        (r"(?:판|©)\s*(?:매|[oO0]f)\s*비\s*와\s*관\s*리\s*비", "판매비와관리비"),
        (r"판\s*매\s*비\s*와\s*관\s*리[^\d\n]{0,5}", "판매비와관리비 "),
        (r"법\s*인\s*셰", "법인세"),
        (r"법\s*인\s*세\s*[등릉]", "법인세등"),
        (r"[당망]\s*[기키]\s*순\s*손\s*익", "당기손익"),
        (r"(?:매|[oO0]f)\s*[출줄]\s*총\s*이\s*[익9¢]", "매출총이익"),
    )
    for pattern, replacement in fuzzy_terms:
        repaired = re.sub(pattern, replacement, repaired)
    for term in terms:
        pattern = r"\s*".join(map(re.escape, term))
        repaired = re.sub(pattern, term, repaired)
    return repaired


def add_continuation_candidates(lines: list[str]) -> list[str]:
    expanded: list[str] = []
    for index, line in enumerate(lines):
        expanded.append(line)
        compact = normalize(line)
        has_target_hint = any(
            hint in compact
            for hint in (
                "자산",
                "부채",
                "자본",
                "매출",
                "출액",
                "영업",
                "당기",
                "순이익",
                "순손실",
                "수수료",
            )
        )
        if not has_target_hint:
            continue
        if len(line) > 500 or is_dense_account_label_blob(line):
            # A page-wide OCR label blob followed by vertical numeric columns
            # has no reliable row alignment.  Joining its first numbers would
            # fabricate totals from unrelated accounts.
            continue
        # A complete two-period row must not absorb the following account.
        # Doing so contaminated totals with values from 자본잉여금 and other
        # adjacent rows.
        if len(extract_amounts(line)) >= 2:
            continue
        lookahead = 2
        if index + lookahead < len(lines):
            combined = " ".join(lines[index : index + lookahead + 1])
            if extract_amounts(combined):
                expanded.append(combined)
    return expanded


def is_dense_account_label_blob(line: str) -> bool:
    """Identify page-wide OCR label bands that have lost their row alignment.

    Spreadsheet-like PDFs sometimes expose every account label as one text line
    and every numeric column as later lines.  Even when the label band is under
    the old 500-character guard, joining its following numbers invents totals.
    These blocks are deliberately left for manual review because there is no
    trustworthy positional relationship left in the extracted text.
    """
    compact = normalize(line)
    labels = (
        "유동자산",
        "비유동자산",
        "자산총계",
        "유동부채",
        "비유동부채",
        "부채총계",
        "자본금",
        "자본총계",
        "부채와자본총계",
        "부채및자본총계",
    )
    label_count = sum(label in compact for label in labels)
    return len(compact) >= 120 and label_count >= 4


def prefer_separate_statement_scope(lines: list[str]) -> list[str]:
    """Do not mix separate-company and consolidated statements in one PDF.

    Some signed report packages contain a full separate set followed by a full
    consolidated set.  Accounting totals from the two sets must never be mixed.
    When a separate statement appears first, keep that complete set and stop at
    the first consolidated statement heading.
    """
    def is_statement_heading(compact: str) -> bool:
        return any(term in compact for term in ("재무상태표", "손익계산서", "포괄손익계산서"))

    separate_indexes = [
        index
        for index, line in enumerate(lines)
        if "별도" in (compact := normalize(line)) and is_statement_heading(compact)
    ]
    consolidated_indexes = [
        index
        for index, line in enumerate(lines)
        if "연결" in (compact := normalize(line)) and is_statement_heading(compact)
    ]
    # A note saying "별도로" is not a separate financial statement.  The old
    # broad check truncated some consolidated reports before their first real
    # statement.
    if not separate_indexes or not consolidated_indexes:
        return lines
    # Preserve the established preference for the separate-company set.  The
    # important correction above is that ordinary prose containing "별도로"
    # no longer activates this truncation.
    return lines[: min(consolidated_indexes)]


def is_amount_only_line(line: str) -> bool:
    if not extract_amounts(line):
        return False
    remainder = AMOUNT_PATTERN.sub("", line)
    remainder = re.sub(r"[\s|:：.·_\-~원₩]", "", remainder)
    return not re.search(r"[가-힣A-Za-z]", remainder)


def find_vertical_alias_segment(
    lines: list[str],
    aliases: tuple[str, ...],
    *,
    amount_count: int = 2,
    positive_only: bool = False,
) -> str:
    """Join values that OCR emitted vertically beneath a row label."""
    for index, line in enumerate(lines):
        compact = normalize(line)
        if not any(alias in compact for alias in aliases):
            continue
        if len(line) > 500 or is_dense_account_label_blob(line):
            continue
        if extract_amounts(line):
            continue
        collected: list[int | float] = []
        pending_prefix: int | None = None
        for following in lines[index + 1 : index + 13]:
            if is_amount_only_line(following):
                following_values = extract_amounts(following)
                # Some spreadsheet-to-PDF exports duplicate the same visual
                # cell four times on one OCR line.  Count that line as one
                # period so the next line can supply the comparison period.
                if len(following_values) >= 2 and len(set(following_values)) == 1:
                    following_values = following_values[:1]
                else:
                    following_values = collapse_ocr_repetitions(following_values)
                if (
                    pending_prefix is not None
                    and following_values
                    and 1_000 <= abs(following_values[0]) < 1_000_000
                    and following.count(",") == 1
                ):
                    combined = pending_prefix * 1_000_000 + abs(following_values[0])
                    following_values[0] = -combined if following_values[0] < 0 else combined
                pending_prefix = None
                for value in following_values:
                    if positive_only and value <= 0:
                        continue
                    collected.append(value)
                    if len(collected) == amount_count:
                        return " ".join([line, *(format_parser_amount(item) for item in collected)])
                continue
            short_prefix = re.fullmatch(r"\s*(\d{1,3})\s*", following)
            if short_prefix and int(short_prefix.group(1)) > 0:
                pending_prefix = int(short_prefix.group(1))
                continue
            pending_prefix = None
            if collected and re.search(r"[가-힣A-Za-z]", following):
                break
    return ""


def has_cumulative_income_layout(lines: list[str]) -> bool:
    for index, line in enumerate(lines):
        compact = normalize(" ".join(lines[index : index + 24]))
        cumulative_markers = (
            compact.count("누적")
            + compact.count("누계")
            + compact.count("두적")
            + compact.count("분기말")
            + compact.count("반기말")
            + compact.count("9개월")
            + compact.count("6개월")
        )
        if (
            compact.count("3개월") >= 2
            and cumulative_markers >= 2
        ):
            return True
    return False


def has_single_period_income_layout(lines: list[str]) -> bool:
    """Return True when the income statement exposes only a current-period column."""
    compact_all = normalize(" ".join(lines))
    if "표준재무제표" in compact_all and "전기" not in compact_all:
        return True
    found_current_only = False
    for index, line in enumerate(lines):
        if "손익계산서" not in normalize(line):
            continue
        raw_header = " ".join(lines[index : index + 18])
        header = normalize(raw_header)
        years = set(re.findall(r"(?:19|20)\d{2}", raw_header))
        periods = set(re.findall(r"제\s*(\d+)\s*기", raw_header))
        if "전기" in header or "(전)기" in header or len(years) >= 2 or len(periods) >= 2:
            return False
        if "당기" in header or "(당)기" in header or len(years) == 1 or len(periods) == 1:
            found_current_only = True
    if found_current_only:
        return True
    return (
        ("손익계산서" in compact_all or "익계산서" in compact_all)
        and "전기" not in compact_all
        and ("당기" in compact_all or "2026" in compact_all)
    )


def has_trailing_income_statement_title(lines: list[str]) -> bool:
    """Detect spreadsheet exports whose statement title is printed last.

    Such files contain one current-period column only.  Returning zero for the
    missing comparison period fabricates a prior-year value, so it must remain
    blank.
    """
    headings = [
        index
        for index, line in enumerate(lines)
        if any(term in normalize(line) for term in INCOME_STATEMENT_HEADINGS)
    ]
    if not headings or headings[-1] < max(0, len(lines) - 5):
        return False
    trailing = headings[-1]
    body_terms = ("매출액", "영업이익", "영업손실", "당기순이익", "당기순손실", "법인세차감전")
    return any(
        extract_amounts(lines[index])
        and any(term in normalize(lines[index]) for term in body_terms)
        for index in range(max(0, trailing - 180), trailing)
    )


def is_first_fiscal_period_income_layout(lines: list[str]) -> bool:
    """Return True for a one-column income statement explicitly labeled 제1기."""
    for index, line in enumerate(lines):
        if "손익계산서" not in normalize(line):
            continue
        raw_header = " ".join(lines[index : index + 18])
        header = normalize(raw_header)
        periods = set(re.findall(r"제\s*(\d+)\s*기", raw_header))
        if periods == {"1"} and "전기" not in header and "(전)기" not in header:
            return True
    return False


def find_cumulative_alias_segment(lines: list[str], aliases: tuple[str, ...]) -> str:
    if not has_cumulative_income_layout(lines):
        return ""
    for line in lines:
        compact = normalize(line)
        for alias in aliases:
            if alias not in compact:
                continue
            if len(extract_amounts(line)) >= 4:
                return line
    return find_vertical_alias_segment(lines, aliases, amount_count=4)


def pick_income_current_previous(numbers: list[int | float], lines: list[str]) -> tuple[int | float | None, int | float | None]:
    numbers = collapse_ocr_repetitions(numbers)
    if has_cumulative_income_layout(lines) and len(numbers) >= 4:
        # Header order: current 3 months, current cumulative,
        # previous 3 months, previous cumulative.
        return numbers[1], numbers[3]
    if has_single_period_income_layout(lines) and (len(numbers) < 2 or numbers[0] == numbers[-1]):
        return (numbers[0], None) if numbers else (None, None)
    return pick_current_previous(numbers)


def english_financial_candidates(lines: list[str]) -> list[str]:
    """Translate common English statement rows into the parser's canonical labels."""
    if not any(
        re.search(pattern, line, flags=re.IGNORECASE)
        for line in lines
        for pattern in (r"balance\s+sheet", r"income\s+statement", r"revenue\s*\(sales\)")
    ):
        return []

    def find_row(patterns: tuple[str, ...]) -> str:
        for line in lines:
            if any(re.search(pattern, line, flags=re.IGNORECASE) for pattern in patterns) and len(extract_amounts(line)) >= 2:
                return line
        return ""

    def translated(label: str, source: str, *, force_loss: bool = False) -> str:
        amounts = extract_amounts(source)
        if force_loss:
            amounts = [-abs(value) for value in amounts]
        return f"{label} " + " | ".join(format_parser_amount(value) for value in amounts[:2])

    candidates: list[str] = []
    current_assets = find_row((r"\bcurrent\s+assets\b",))
    noncurrent_assets = find_row((r"\bnon[- ]?current\s+assets\b",))
    current_liabilities = find_row((r"\bcurrent\s+liabilities\b",))
    noncurrent_liabilities = find_row((r"\bnon[- ]?current\s+liabilities\b",))

    def add_section_total(label: str, first: str, second: str) -> None:
        left = extract_amounts(first)
        right = extract_amounts(second)
        if len(left) >= 2 and len(right) >= 2:
            candidates.append(
                f"{label} {format_parser_amount(left[0] + right[0])} | "
                f"{format_parser_amount(left[1] + right[1])}"
            )

    add_section_total("자산총계", current_assets, noncurrent_assets)
    add_section_total("부채총계", current_liabilities, noncurrent_liabilities)

    capital = find_row((r"\bcontributed\s+capital\b", r"\bshare\s+capital\b", r"\bcapital\s+stock\b"))
    if capital:
        candidates.append(translated("영문 자본금", capital))

    revenue = find_row((r"\brevenue\s*\(sales\)", r"^\s*(?:[ivx]+[.)]?\s*)?revenue\b", r"^\s*sales\b"))
    if revenue:
        candidates.append(translated("매출액", revenue))

    operating = find_row((r"\boperati+ng\s+(?:profit|income|loss)\b",))
    if operating:
        candidates.append(translated("영업손익", operating, force_loss=bool(re.search(r"loss", operating, re.I))))

    net_income = find_row((r"\bnet\s+(?:profit|income|loss)\b", r"\bprofit\s*\(loss\)\s+for\s+the\s+period\b"))
    if net_income:
        candidates.append(translated("당기순이익", net_income, force_loss=bool(re.search(r"loss", net_income, re.I))))
    return candidates


def format_parser_amount(value: int | float) -> str:
    if isinstance(value, float) and not value.is_integer():
        return f"{value:,.6f}".rstrip("0").rstrip(".")
    return f"{int(value):,}"


def find_nearby_reversed_total(lines: list[str], aliases: tuple[str, ...]) -> str:
    """Recover exports that place a value row immediately before its label."""
    for index, line in enumerate(lines):
        compact = normalize(line)
        matched_alias = next((alias for alias in aliases if compact == alias), "")
        if not matched_alias or extract_amounts(line):
            continue
        for previous_index in range(index - 1, max(-1, index - 5), -1):
            candidate = lines[previous_index]
            amounts = extract_amounts(candidate)
            if (
                len(amounts) == 2
                and all(value > 0 for value in amounts)
                and re.fullmatch(r"[\s\d,.]+", candidate)
            ):
                return (
                    f"{matched_alias} {format_parser_amount(amounts[0])} "
                    f"{format_parser_amount(amounts[1])}"
                )
    return ""


def find_combined_balance_total(lines: list[str]) -> str:
    aliases = ("부채와자본총계", "부채및자본총계")
    exact_candidates: list[str] = []
    fallback_candidates: list[str] = []
    for line in lines:
        matched_alias = next((alias for alias in aliases if alias in normalize(line)), "")
        if not matched_alias or len(line) > 800 or is_dense_account_label_blob(line):
            continue
        alias_position = normalize(line).find(matched_alias)
        values_after_alias = financial_amounts(line, aliases)
        if (
            "자본총계" in normalize(line)[:alias_position]
            and len(values_after_alias) >= 4
        ):
            return (
                f"{matched_alias} {format_parser_amount(values_after_alias[1])} "
                f"{format_parser_amount(values_after_alias[3])}"
            )
        segment = slice_alias_amounts(line, matched_alias)
        if not segment:
            continue
        numbers = financial_amounts(segment, aliases)
        if numbers and numbers[0] > 0:
            fallback_candidates.append(segment)
            if len(numbers) == 2:
                exact_candidates.append(segment)
    if exact_candidates:
        return min(exact_candidates, key=len)
    if fallback_candidates:
        return fallback_candidates[0]
    reversed_total = find_nearby_reversed_total(lines, aliases)
    if reversed_total:
        return reversed_total
    return find_vertical_alias_segment(lines, aliases, amount_count=2, positive_only=True)


def find_interleaved_equity_balance_totals(
    lines: list[str],
) -> tuple[str, list[int | float], list[int | float]] | None:
    """Parse ``자본총계/부채및자본총계`` printed as alternating columns."""
    combined_aliases = ("부채와자본총계", "부채및자본총계")
    for line in lines:
        compact = normalize(line)
        if (
            len(line) > 800
            or is_dense_account_label_blob(line)
            or "자본총계" not in compact
            or not any(alias in compact for alias in combined_aliases)
        ):
            continue
        first_label = compact.find("자본총계")
        # Use the original line from the equity label onward.  The visual
        # column order is equity-current, balance-current, equity-previous,
        # balance-previous.
        normalized_line = normalize_ocr_numbers(line)
        label_match = re.search(r"자\s*본\s*총\s*계", normalized_line)
        tail = normalized_line[label_match.start() :] if label_match else normalized_line
        amounts = financial_amounts(tail, ("자본총계",))
        if len(amounts) >= 4 and all(value > 0 for value in amounts[:4]):
            return line, [amounts[0], amounts[2]], [amounts[1], amounts[3]]
    return None


def find_vertical_share_capital_from_balance(
    lines: list[str],
) -> tuple[str, list[int | float]] | None:
    """Recover vertical share-capital rows using the balance equation anchors."""
    interleaved = find_interleaved_equity_balance_totals(lines)
    if not interleaved:
        return None
    _source, equity_values, balance_values = interleaved
    liabilities = [balance - equity for balance, equity in zip(balance_values, equity_values)]
    recovered: list[int | float] = []
    sources: list[str] = []
    for liability in liabilities:
        capital_value = None
        for index, line in enumerate(lines[:-1]):
            amounts = extract_amounts(line)
            if len(amounts) != 1 or amounts[0] != liability:
                continue
            for following in lines[index + 1 : index + 3]:
                following_amounts = extract_amounts(following)
                if (
                    len(following_amounts) == 1
                    and 10_000 <= following_amounts[0] < liability
                    and re.fullmatch(r"[\s\d,.]+", following)
                ):
                    capital_value = following_amounts[0]
                    sources.extend((line, following))
                    break
            if capital_value is not None:
                break
        if capital_value is None:
            return None
        recovered.append(capital_value)
    return " / ".join(sources), recovered


def find_section_heading_total(lines: list[str], section: str) -> str:
    child = "유동자산" if section == "자산" else "유동부채"
    for index, line in enumerate(lines):
        letters = re.sub(r"[^가-힣]", "", line)
        if letters != section or len(extract_amounts(line)) < 1:
            continue
        nearby = normalize(" ".join(lines[index + 1 : index + 8]))
        if child in nearby:
            return line
    return ""


def find_columnar_period_total(lines: list[str], alias: str) -> str:
    """Recover a total from OCR that emitted each comparison column vertically.

    This intentionally requires an explicit ``(비교)`` period header and takes
    the final standalone amount in each column.  Without both anchors, returning
    blank is safer than borrowing an unrelated account value.
    """
    for label_index, line in enumerate(lines):
        if normalize(line) != alias or extract_amounts(line):
            continue
        comparison_index = next(
            (
                index
                for index in range(label_index + 1, min(len(lines), label_index + 80))
                if "비교" in normalize(lines[index])
                and re.search(r"제\s*\d+\s*기", lines[index])
            ),
            -1,
        )
        if comparison_index < 0:
            continue

        def standalone_candidates(start: int, end: int) -> list[int | float]:
            values: list[int | float] = []
            for candidate in lines[start:end]:
                amounts = extract_amounts(candidate)
                if len(amounts) != 1 or re.search(r"(?:19|20)\d{2}[./-]", candidate):
                    continue
                if not re.fullmatch(r"[\s()\-\d,.]+", candidate):
                    continue
                if abs(amounts[0]) >= 1_000_000:
                    values.append(amounts[0])
            return values

        current_values = standalone_candidates(label_index + 1, comparison_index)
        previous_end = min(len(lines), comparison_index + 45)
        for index in range(comparison_index + 1, previous_end):
            compact = normalize(lines[index])
            if index > comparison_index + 1 and (
                lines[index].startswith("---")
                or re.search(r"(?:19|20)\d{2}/\d{2}/\d{2}.*\d{1,2}:\d{2}", lines[index])
            ):
                previous_end = index
                break
        previous_values = standalone_candidates(comparison_index + 1, previous_end)
        if current_values and previous_values:
            return (
                f"{alias} {format_parser_amount(current_values[-1])} "
                f"{format_parser_amount(previous_values[-1])}"
            )
    return ""


def find_balance_component_total(
    lines: list[str], section: str
) -> tuple[str, int | float | None, int | float | None] | None:
    """Sum current/non-current subtotals inside the primary balance sheet.

    This is a guarded cross-check for a damaged total row, not a general
    replacement for an explicit total.
    """
    if section == "자산":
        current_alias, noncurrent_alias = "유동자산", "비유동자산"
    else:
        current_alias, noncurrent_alias = "유동부채", "비유동부채"

    income_start = income_statement_start(lines)
    balance_lines = lines[:income_start] if income_start > 0 else lines[:400]

    def best_component(alias: str, *, exclude: str = "") -> tuple[str, list[int | float]] | None:
        candidates: list[tuple[str, list[int | float]]] = []
        for line in balance_lines:
            compact = normalize(line)
            if alias not in compact:
                continue
            # ``유동부채`` is a substring of ``비유동부채``.  Reject only a
            # substring-only hit; a flattened OCR line may legitimately hold
            # both current and non-current rows and must be sliced by label.
            if exclude and alias not in compact.replace(exclude, ""):
                continue
            if alias == noncurrent_alias and f"기타{noncurrent_alias}" in compact:
                continue
            segment = slice_alias_amounts(line, alias)
            if not segment or len(segment) > 800 or len(extract_amounts(segment)) > 4:
                continue
            amounts = financial_amounts(segment, (alias,))
            raw_amounts = AMOUNT_PATTERN.findall(normalize_ocr_numbers(segment))
            clean_cells = sum(
                bool(re.fullmatch(r"\(?-?(?:\d{1,3}(?:,\d{3})+|0)\)?", raw.strip()))
                for raw in raw_amounts[:2]
            )
            if len(amounts) >= 2 and clean_cells >= 2:
                candidates.append((segment, amounts))
        if not candidates:
            return None
        return max(
            candidates,
            key=lambda item: (
                min(len(item[1]), 2),
                sum(value != 0 for value in item[1][:2]),
                0,
            ),
        )

    current_part = best_component(current_alias, exclude=noncurrent_alias)
    noncurrent_part = best_component(noncurrent_alias)
    if not current_part or not noncurrent_part:
        return None
    current_values = pick_current_previous(current_part[1])
    noncurrent_values = pick_current_previous(noncurrent_part[1])
    summed = []
    for current_value, noncurrent_value in zip(current_values, noncurrent_values):
        if current_value is None or noncurrent_value is None:
            summed.append(None)
        else:
            summed.append(current_value + noncurrent_value)
    return f"{current_part[0]} / {noncurrent_part[0]}", summed[0], summed[1]


def prefer_component_crosscheck(
    direct: int | float | None, component: int | float | None
) -> bool:
    if component in (None, 0) or direct in (None, 0):
        return False
    if direct == component:
        return False
    direct_abs = abs(direct)
    component_abs = abs(component)
    # Only override a direct total when its magnitude is unmistakably damaged
    # (typically a dropped group of three OCR digits).  A full-sized direct
    # total remains authoritative even if a noisy note subtotal is much larger.
    obvious_magnitude_damage = (
        direct_abs < 1_000_000
        and component_abs >= 1_000_000
        and direct_abs * 100 < component_abs
        and component_abs <= direct_abs * 1_000_000
    )
    return obvious_magnitude_damage


def extract_assets_total(lines: list[str], label: str) -> dict | None:
    # A clean two-period ``부채및자본총계`` is an accounting-identity crosscheck
    # for an OCR-damaged asset row.  It is safe only when it has exactly two
    # values; a four-value capital/combined row must never donate its first pair.
    # Take the first explicit row even when it contains only the current period.
    # Requiring two values first can select a synthetic continuation that joins
    # the next liability row and mislabels it as the previous-period asset total.
    direct_match = find_columnar_period_total(lines, "자산총계")
    if not direct_match:
        for line in lines:
            compact = normalize(line)
            if "자산총계" not in compact or any(
                subtotal in compact
                for subtotal in ("유동자산총계", "비유동자산총계", "기타비유동자산총계")
            ):
                continue
            segment = slice_alias_amounts(line, "자산총계")
            direct_amounts = financial_amounts(segment, ("자산총계",)) if segment else []
            if not direct_amounts:
                continue
            first_periods = direct_amounts[:2]
            if first_periods and all(1900 <= abs(value) <= 2100 for value in first_periods):
                continue
            if (
                len(first_periods) >= 2
                and any(1900 <= abs(value) <= 2100 for value in first_periods)
                and any(abs(value) >= 1_000_000 for value in first_periods)
            ):
                continue
            direct_match = segment
            break
    combined_match = find_combined_balance_total(lines)
    combined_amounts = (
        financial_amounts(
            combined_match,
            ("부채와자본총계", "부채및자본총계", "자본과부채총계"),
        )
        if combined_match
        else []
    )
    combined_has_year = (
        len(combined_amounts) >= 2
        and any(1900 <= abs(value) <= 2100 for value in combined_amounts[:2])
        and any(abs(value) >= 1_000_000 for value in combined_amounts[:2])
    )
    match = direct_match
    if not match and combined_match and not combined_has_year:
        match = combined_match
    if not match:
        match = find_section_heading_total(lines, "자산")
    if not match:
        # OCR often damages only the first character (for example "Xt 산 총 계").
        # A line ending in "산총계" immediately before the liabilities section is
        # still a much stronger total candidate than a non-current asset subtotal.
        fuzzy_totals = [
            line
            for line in lines
            if "산총계" in normalize(line)
            and not any(term in normalize(line) for term in ("부채", "자본"))
            and not any(
                subtotal in normalize(line)
                for subtotal in ("유동자산총계", "비유동자산총계", "기타비유동자산총계")
            )
            and extract_amounts(line)
            and any(value != 0 for value in extract_amounts(line))
        ]
        if fuzzy_totals:
            match = max(
                fuzzy_totals,
                key=lambda line: max(
                    (abs(value) for value in financial_amounts(line, ("산총계",))),
                    default=0,
                ),
            )
    if not match:
        match = find_best_line(lines, ["자산총계", "자 산 총 계"], label)
        if match and any(
            subtotal in normalize(match)
            for subtotal in ("유동자산총계", "비유동자산총계", "기타비유동자산총계")
        ):
            match = ""
    if match and not extract_amounts(match):
        match = ""
    if not match:
        for index, line in enumerate(lines):
            compact = normalize(line)
            if "자산" in compact and extract_amounts(line):
                following = "".join(normalize(item) for item in lines[index + 1 : index + 5])
                previous = "".join(normalize(item) for item in lines[max(0, index - 20) : index])
                if "부채" in following and ("유동자산" in previous or "비유동자산" in previous or "자산" in previous):
                    match = line
                    break
            if "산" in compact and extract_amounts(line):
                if any(
                    subtotal in compact
                    for subtotal in (
                        "유동자산",
                        "비유동자산",
                        "당좌자산",
                        "재고자산",
                        "투자자산",
                        "유형자산",
                        "무형자산",
                        "기타비유동자산",
                    )
                ) or not any(value != 0 for value in extract_amounts(line)):
                    continue
                following = "".join(normalize(item) for item in lines[index + 1 : index + 5])
                previous = "".join(normalize(item) for item in lines[max(0, index - 20) : index])
                if "부채" in following and ("유동" in previous or "비유동" in previous):
                    match = line
                    break
    if not match:
        return None
    numbers = financial_amounts(
        match,
        ("자산총계", "자산 총계", "산총계", "부채와자본총계", "부채및자본총계"),
    )
    current, previous = pick_current_previous(numbers)
    current = abs(current) if current is not None else None
    previous = abs(previous) if previous is not None else None
    # A clean balance-equation total is an independent check on the asset row.
    # When the two differ only slightly, the asset row is commonly missing or
    # misreading one OCR digit while the combined total remains intact.
    combined_current, combined_previous = pick_current_previous(combined_amounts)
    combined_replaced = False
    if not combined_has_year:
        if (
            prefer_component_crosscheck(current, combined_current)
            or current not in (None, 0)
            and combined_current not in (None, 0)
            and abs(abs(current) - abs(combined_current))
            <= max(1_000, max(abs(current), abs(combined_current)) * 0.02)
        ):
            current = abs(combined_current) if combined_current is not None else None
            combined_replaced = True
        if (
            prefer_component_crosscheck(previous, combined_previous)
            or previous not in (None, 0)
            and combined_previous not in (None, 0)
            and abs(abs(previous) - abs(combined_previous))
            <= max(1_000, max(abs(previous), abs(combined_previous)) * 0.02)
        ):
            previous = abs(combined_previous) if combined_previous is not None else None
            combined_replaced = True
    component_total = find_balance_component_total(lines, "자산")
    status = "부채·자본 합계로 교차검증" if combined_replaced else "확인 필요"
    if combined_replaced:
        match = f"{match} / 교차검증: {combined_match}"
    if component_total:
        component_source, component_current, component_previous = component_total
        replace_current = prefer_component_crosscheck(current, component_current)
        replace_previous = prefer_component_crosscheck(previous, component_previous)
        if replace_current:
            current = abs(component_current) if component_current is not None else None
        if replace_previous:
            previous = abs(component_previous) if component_previous is not None else None
        if replace_current or replace_previous:
            match = f"{match} / 교차검증: {component_source}"
            status = "유동·비유동 합계로 교차검증"
    if current not in (None, 0) and previous not in (None, 0):
        larger = max(abs(current), abs(previous))
        smaller = min(abs(current), abs(previous))
        if smaller < 10_000_000 and larger > smaller * 500:
            if abs(current) == smaller:
                current = None
            else:
                previous = None
            status = "OCR 숫자 자릿수 손상 - 원문 확인 필요"
    return {
        "item": label,
        "current": current if current is not None else "",
        "previous": previous if previous is not None else "",
        "status": status if current is not None or previous is not None else "금액 없음",
        "sourceLine": match,
    }


def extract_liabilities_total(lines: list[str], label: str) -> dict | None:
    malformed_group_pattern = re.compile(
        r"(?<![\d,])\(?-?\d{1,3}(?:,\s?\d{3})+,\s?\d{1,2}(?!\d)"
    )

    def malformed_period_indexes(candidate: str) -> set[int]:
        normalized = normalize_ocr_numbers(candidate)
        label_positions = [
            normalized.find(alias)
            for alias in ("부채총계", "부채 총계", "채총계", "채 계", "부채")
            if normalized.find(alias) >= 0
        ]
        if label_positions:
            normalized = normalized[min(label_positions):]
        ordinary = list(AMOUNT_PATTERN.finditer(normalized))
        damaged: set[int] = set()
        for malformed in malformed_group_pattern.finditer(normalized):
            for index, amount in enumerate(ordinary[:2]):
                if amount.start() == malformed.start():
                    damaged.add(index)
                    break
        return damaged

    direct_segment = find_quality_alias_segment(lines, ("부채총계",), require_two=False)
    if not direct_segment:
        direct_segment = find_section_heading_total(lines, "부채")
    if not direct_segment:
        direct_segment = find_nearby_reversed_total(lines, ("부채총계",))
    candidates = []
    for index, line in enumerate(lines):
        compact = normalize(line)
        compact_label = re.sub(r"[^가-힣A-Za-z0-9]", "", line)
        if any(
            combined in compact
            for combined in ("부채와자본총계", "부채및자본총계", "자본과부채총계")
        ):
            continue
        liability_total_label = (
            "부채총계" in compact_label
            or "채총계" in compact_label
            or compact_label.startswith("채계")
            or bool(re.search(r"부채총(?:A[lI1]|[Il1])", compact, re.I))
        )
        if liability_total_label:
            candidate = slice_alias_amounts(line, "부채총계") or line
            if (
                len(extract_amounts(candidate)) >= 1
                and len(candidate) <= 800
                and len(extract_amounts(candidate)) <= 4
            ):
                candidates.append(candidate)
    if (
        direct_segment
        and len(direct_segment) <= 800
        and len(extract_amounts(direct_segment)) <= 4
        and direct_segment not in candidates
    ):
        candidates.append(direct_segment)
    if not candidates:
        interleaved = find_interleaved_equity_balance_totals(lines)
        if interleaved:
            source_line, equity_values, balance_values = interleaved
            current = balance_values[0] - equity_values[0]
            previous = balance_values[1] - equity_values[1]
            if current >= 0 and previous >= 0:
                return {
                    "item": label,
                    "current": current,
                    "previous": previous,
                    "status": "부채·자본 합계 - 자본총계 계산",
                    "sourceLine": source_line,
                }
        component_total = find_balance_component_total(lines, "부채")
        if component_total:
            component_source, current, previous = component_total
            return {
                "item": label,
                "current": abs(current) if current is not None else "",
                "previous": abs(previous) if previous is not None else "",
                "status": "유동·비유동 부채 합산",
                "sourceLine": component_source,
            }
        combined_match = find_combined_balance_total(lines)
        capital_match = find_vertical_alias_segment(lines, ("자본총계",), amount_count=2)
        if combined_match and capital_match:
            combined_current, combined_previous = pick_current_previous(
                financial_amounts(combined_match, ("부채와자본총계", "부채및자본총계"))
            )
            capital_current, capital_previous = pick_current_previous(
                financial_amounts(capital_match, ("자본총계",))
            )
            current = subtract_values(combined_current, capital_current)
            previous = subtract_values(combined_previous, capital_previous)
            # A balance total minus equity can never produce a negative
            # liability.  If OCR alignment makes it negative, omit that period
            # rather than publishing a mathematically impossible value.
            if current is not None and current < 0:
                current = None
            if previous is not None and previous < 0:
                previous = None
            if current is not None or previous is not None:
                return {
                    "item": label,
                    "current": current if current is not None else "",
                    "previous": previous if previous is not None else "",
                    "status": "부채와자본총계 - 자본총계 계산",
                    "sourceLine": f"{combined_match} / {capital_match}",
                }
        return None
    def liability_amounts(candidate: str) -> list[int | float]:
        compact_label = re.sub(r"[^가-힣A-Za-z0-9]", "", candidate)
        aliases = (
            ("부채총계", "부채 총계", "채총계", "채 계")
            if "부채총계" in compact_label or "채총계" in compact_label
            else ("부채",)
        )
        return financial_amounts(candidate, aliases)

    def liability_candidate_rank(candidate: str) -> tuple[int, int, int, int, int, int]:
        amounts = liability_amounts(candidate)
        raw_amounts = AMOUNT_PATTERN.findall(candidate)
        comma_grouped = sum(
            bool(re.fullmatch(r"\(?-?\d{1,3}(?:,\d{3})+\)?", raw.strip()))
            for raw in raw_amounts[:2]
        )
        thousands_groups = sum(raw.strip().count(",") for raw in raw_amounts[:2])
        nonzero = sum(value != 0 for value in amounts[:2])
        return (
            int(not malformed_period_indexes(candidate)),
            min(len(amounts), 2),
            nonzero,
            comma_grouped,
            thousands_groups,
            lines.index(candidate) if candidate in lines else -1,
        )

    match = max(candidates, key=liability_candidate_rank)
    amounts = liability_amounts(match)
    current, previous = pick_current_previous(amounts)
    current = abs(current) if current is not None else None
    previous = abs(previous) if previous is not None else None
    malformed_periods = malformed_period_indexes(match)
    component_total = find_balance_component_total(lines, "부채")
    status = (
        "OCR 숫자 자릿수 손상 - 원문 확인 필요"
        if malformed_periods
        else "확인 필요"
    )
    replaced_malformed: set[int] = set()
    if component_total:
        component_source, component_current, component_previous = component_total
        replace_current = (
            0 in malformed_periods and component_current not in (None, 0)
        ) or prefer_component_crosscheck(current, component_current)
        replace_previous = (
            1 in malformed_periods and component_previous not in (None, 0)
        ) or prefer_component_crosscheck(previous, component_previous)
        if replace_current:
            current = abs(component_current) if component_current is not None else None
            if 0 in malformed_periods:
                replaced_malformed.add(0)
        if replace_previous:
            previous = abs(component_previous) if component_previous is not None else None
            if 1 in malformed_periods:
                replaced_malformed.add(1)
        if replace_current or replace_previous:
            match = f"{match} / 교차검증: {component_source}"
            status = "유동·비유동 합계로 교차검증"
    if 0 in malformed_periods and 0 not in replaced_malformed:
        current = None
    if 1 in malformed_periods and 1 not in replaced_malformed:
        previous = None
    return {
        "item": label,
        "current": current if current is not None else "",
        "previous": previous if previous is not None else "",
        "status": status if current is not None or previous is not None else "금액 없음",
        "sourceLine": match,
    }


def find_capital_component_total(lines: list[str], capital_lines: list[str]) -> tuple[str, list[int]] | None:
    anchor_indexes = [index for index, line in enumerate(lines) if line in capital_lines]
    for anchor_index in anchor_indexes:
        for total_index in range(max(0, anchor_index - 4), anchor_index + 1):
            total_amounts = extract_amounts(lines[total_index])
            if len(total_amounts) != 2 or total_amounts[0] <= 0:
                continue

            component_amounts: list[list[int]] = []
            for component_index in range(total_index + 1, min(len(lines), total_index + 5)):
                amounts = extract_amounts(lines[component_index])
                if len(amounts) != 2:
                    continue
                if abs(amounts[0]) > abs(total_amounts[0]) or abs(amounts[1]) > abs(total_amounts[1]):
                    break
                component_amounts.append(amounts)
                if len(component_amounts) < 2:
                    continue
                current_sum = sum(amounts[0] for amounts in component_amounts)
                previous_sum = sum(amounts[1] for amounts in component_amounts)
                if current_sum == total_amounts[0] and previous_sum == total_amounts[1]:
                    return lines[total_index], total_amounts
                if current_sum > total_amounts[0] or previous_sum > total_amounts[1]:
                    break
    return None


def find_share_class_capital_total(
    lines: list[str],
) -> tuple[str, list[int | float]] | None:
    """Sum common and preferred share capital without treating them as periods.

    OCR frequently flattens the aggregate, common-share and preferred-share
    rows onto one physical line.  In that layout the old parser read the two
    component values as current/previous periods.  Parse each labelled row
    independently and add the share classes instead.
    """
    explicit_pattern = re.compile(
        r"(?P<common>보\s*통\s*주\s*자본금)|(?P<preferred>우\s*선\s*주\s*자본금)"
    )
    plain_pattern = re.compile(r"자본금")
    trailing_boundary = re.compile(
        r"자본잉여금|주식발행초과금|자본조정|기타포괄손익|이익잉여금|"
        r"결손금|자본총계|부채(?:와|및)?자본총계"
    )

    candidates: list[tuple[int, str, str, list[int | float], int]] = []
    for line_index, line in enumerate(lines):
        explicit = list(explicit_pattern.finditer(line))
        explicit_spans = [match.span() for match in explicit]
        labelled: list[tuple[int, int, str]] = []
        for match in explicit:
            labelled.append(
                (match.start(), match.end(), "common" if match.lastgroup == "common" else "preferred")
            )
        for match in plain_pattern.finditer(line):
            if any(start <= match.start() < end for start, end in explicit_spans):
                continue
            labelled.append((match.start(), match.end(), "plain"))
        labelled.sort()
        for position, (start, end, kind) in enumerate(labelled):
            stop = labelled[position + 1][0] if position + 1 < len(labelled) else len(line)
            boundary = trailing_boundary.search(line, end, stop)
            if boundary:
                stop = boundary.start()
            segment = line[start:stop]
            amounts = financial_amounts(segment, ("자본금",))[:2]
            if (
                not amounts
                or all(1900 <= abs(value) <= 2100 for value in amounts)
                or any(not isinstance(value, int) or abs(value) < 10_000 for value in amounts)
            ):
                continue
            candidates.append((line_index, kind, segment, amounts, start))

    preferred = [candidate for candidate in candidates if candidate[1] == "preferred"]
    if not preferred:
        return None

    def pair_rank(pair: tuple[tuple, tuple]) -> tuple[int, int, int]:
        common_row, preferred_row = pair
        return (
            int(common_row[0] == preferred_row[0]),
            min(len(common_row[3]), len(preferred_row[3])),
            -abs(common_row[0] - preferred_row[0]),
        )

    pairs: list[tuple[tuple, tuple]] = []
    explicit_common = [candidate for candidate in candidates if candidate[1] == "common"]
    for preferred_row in preferred:
        nearby_common = [
            candidate
            for candidate in explicit_common
            if abs(candidate[0] - preferred_row[0]) <= 4
        ]
        if not nearby_common:
            # If OCR lost the word "보통주", the last plain 자본금 immediately
            # before 우선주자본금 is the common-share component.  Choosing the
            # last one avoids an earlier aggregate 자본금 row.
            nearby_common = [
                candidate
                for candidate in candidates
                if candidate[1] == "plain"
                and (
                    candidate[0] < preferred_row[0]
                    and preferred_row[0] - candidate[0] <= 3
                    or candidate[0] == preferred_row[0] and candidate[4] < preferred_row[4]
                )
            ]
            if nearby_common:
                nearby_common = [max(nearby_common, key=lambda item: (item[0], item[4]))]
        pairs.extend((common_row, preferred_row) for common_row in nearby_common)
    if not pairs:
        return None
    common_row, preferred_row = max(pairs, key=pair_rank)
    current = abs(common_row[3][0]) + abs(preferred_row[3][0])
    amounts: list[int | float] = [current]
    if len(common_row[3]) >= 2 and len(preferred_row[3]) >= 2:
        amounts.append(abs(common_row[3][1]) + abs(preferred_row[3][1]))
    aggregate_rows = [
        candidate
        for candidate in candidates
        if candidate[1] == "plain"
        and candidate != common_row
        and (
            candidate[0] < common_row[0] and common_row[0] - candidate[0] <= 1
            or candidate[0] == common_row[0] and candidate[4] < common_row[4]
        )
        and (
            abs(candidate[3][0]) >= abs(common_row[3][0])
            and abs(candidate[3][0]) >= abs(preferred_row[3][0])
            or len(candidate[3]) >= 2
            and len(amounts) >= 2
            and abs(candidate[3][1]) == abs(amounts[1])
        )
    ]
    if aggregate_rows:
        aggregate = max(aggregate_rows, key=lambda item: (item[0], item[4]))
        return aggregate[2], [abs(value) for value in aggregate[3]]
    if common_row[1] == "plain":
        nearby_values: list[int | float] = []
        for line_index in range(max(0, common_row[0] - 1), min(len(lines), common_row[0] + 2)):
            nearby_values.extend(abs(value) for value in extract_amounts(lines[line_index]))
        has_aggregate_evidence = all(abs(value) in nearby_values for value in amounts)
        if not has_aggregate_evidence:
            # With no explicit 보통주 label or independent summed total, the
            # plain 자본금 row is itself more likely the aggregate (some forms
            # list additional preferred-share subclasses on later rows).
            return common_row[2], [abs(value) for value in common_row[3]]
    return f"{common_row[2]} / {preferred_row[2]}", amounts


def find_capital_section_total(lines: list[str]) -> tuple[str, list[int]] | None:
    """Recover the bold capital total when OCR mangles its Korean row label."""
    for section_index, line in enumerate(lines):
        section_letters = re.sub(r"[^가-힣A-Za-z]", "", line)
        exact_heading = section_letters in ("자본", "쟈본", "At", "Xt")
        noisy_latin_heading = (
            not exact_heading
            and len(section_letters) <= 6
            and section_letters.startswith(("At", "Xt"))
        )
        if not exact_heading and not noisy_latin_heading:
            continue
        for candidate_index in range(section_index + 1, min(len(lines), section_index + 5)):
            candidate = lines[candidate_index]
            amounts = extract_amounts(candidate)
            if len(amounts) < 2 or amounts[0] <= 0:
                continue
            if all(1900 <= abs(value) <= 2100 for value in amounts[:2]):
                continue
            search_end = candidate_index + (16 if noisy_latin_heading else 6)
            following_lines = lines[candidate_index + 1 : search_end]
            following = normalize(" ".join(following_lines))
            has_plain_capital = "자본금" in following or "자봄금" in following
            has_preferred_capital = bool(re.search(r"우.{0,3}주.{0,3}(?:자)?본.{0,2}금", following))
            if noisy_latin_heading:
                common_lines = [
                    item for item in following_lines
                    if "보통주" in normalize(item) and "자본금" in normalize(item)
                ]
                preferred_lines = [
                    item for item in following_lines
                    if "우선주" in normalize(item) and "자본금" in normalize(item)
                ]
                if not common_lines or not preferred_lines:
                    continue
                common_values = financial_amounts(common_lines[0], ("자본금",))
                preferred_values = financial_amounts(preferred_lines[0], ("자본금",))
                if len(common_values) < 2 or len(preferred_values) < 2:
                    continue
                if (
                    abs(common_values[0]) + abs(preferred_values[0]) == abs(amounts[0])
                    and abs(common_values[1]) + abs(preferred_values[1]) == abs(amounts[1])
                ):
                    return candidate, amounts[:2]
                continue
            if has_plain_capital and (has_preferred_capital or "자본잉여금" in following):
                return candidate, amounts[:2]
    return None


def extract_capital(lines: list[str], label: str) -> dict | None:
    capital_lines = [
        line
        for line in lines
        if "자본금" in normalize(line) and "자본금계" not in normalize(line) and extract_amounts(line)
    ]
    if not capital_lines:
        return None

    vertical_total = find_vertical_share_capital_from_balance(lines)
    if vertical_total:
        source_line, amounts = vertical_total
        current, previous = pick_current_previous(amounts)
        return {
            "item": label,
            "current": current if current is not None else "",
            "previous": previous if previous is not None else "",
            "status": "수직 열 자본금 교차검증",
            "sourceLine": source_line,
        }

    section_total = find_capital_section_total(lines)
    if section_total:
        source_line, amounts = section_total
        current, previous = pick_current_previous(amounts)
        return {
            "item": label,
            "current": abs(current) if current is not None else "",
            "previous": abs(previous) if previous is not None else "",
            "status": "자본 구성 총액 사용",
            "sourceLine": source_line,
        }

    share_class_total = find_share_class_capital_total(lines)
    if share_class_total:
        source_line, amounts = share_class_total
        current, previous = pick_current_previous(amounts)
        return {
            "item": label,
            "current": current if current is not None else "",
            "previous": previous if previous is not None else "",
            "status": "보통주·우선주 자본금 합산",
            "sourceLine": source_line,
        }

    component_total = find_capital_component_total(lines, capital_lines)
    if component_total:
        source_line, amounts = component_total
        current, previous = pick_current_previous(amounts)
        return {
            "item": label,
            "current": current if current is not None else "",
            "previous": previous if previous is not None else "",
            "status": "구성항목 합계 확인",
            "sourceLine": source_line,
        }

    def capital_segment(line: str) -> str:
        return slice_alias_amounts(line, "자본금") or line

    def capital_candidate_rank(line: str) -> tuple[int, int, int | float]:
        amounts = financial_amounts(capital_segment(line), ("자본금",))
        normalized = normalize_ocr_numbers(line)
        label_match = re.search(r"자\s*본\s*금", normalized)
        prefix = normalized[: label_match.start()] if label_match else ""
        # Roman-numbered ``Ⅰ.자본금`` is often a section aggregate that also
        # includes 기타자본.  Prefer the following plain ``자본금`` detail row.
        is_section_heading = bool(re.search(r"[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]", prefix))
        return (
            int(not is_section_heading),
            min(len(amounts), 2),
            max((abs(amount) for amount in amounts[:2]), default=-1),
        )

    plain = [
        line
        for line in capital_lines
        if not re.search(r"보\s*통\s*주|우\s*선\s*주|발행|초과", line)
    ]
    common = [line for line in capital_lines if re.search(r"보\s*통\s*주", line)]
    if plain:
        source_lines = [
            max(
                plain,
                key=capital_candidate_rank,
            )
        ]
    elif common:
        source_lines = common[:1]
    else:
        source_lines = capital_lines[:1]
    # A continuation candidate may contain ``aggregate / common / preferred``
    # capital rows on one line.  Limit parsing to the first capital row so the
    # second component is never mistaken for the comparison period.
    source_segment = capital_segment(source_lines[0])
    if len(source_segment) > 800 or len(extract_amounts(source_segment)) > 4:
        return None
    amounts = financial_amounts(source_segment, ("자본금",))
    current, previous = pick_current_previous(amounts)
    current = abs(current) if current is not None else None
    previous = abs(previous) if previous is not None else None
    allow_decimal = any("영문" in normalize(line) for line in source_lines)
    if current is not None and ((not allow_decimal and not isinstance(current, int)) or abs(current) < 10_000):
        current = None
    if previous is not None and ((not allow_decimal and not isinstance(previous, int)) or abs(previous) < 10_000):
        previous = None
    if current not in (None, 0) and previous not in (None, 0):
        larger = max(abs(current), abs(previous))
        smaller = min(abs(current), abs(previous))
        if smaller < 10_000_000 and larger > smaller * 500:
            if abs(current) == smaller:
                current = None
            else:
                previous = None
    return {
        "item": label,
        "current": current if current is not None else "",
        "previous": previous if previous is not None else "",
        "status": "확인 필요",
        "sourceLine": " / ".join(source_lines),
    }


def apply_loss_row_signs(line: str, numbers: list[int]) -> list[int]:
    if not numbers:
        return numbers
    compact = normalize(line)
    explicit_loss = any(
        alias in compact
        for alias in ("영업손실", "당기순손실", "분기순손실", "반기순손실", "순손실")
    ) and not ("이익" in compact and "손실" in compact)
    if explicit_loss:
        # The row label applies to every displayed period.  A leading minus on
        # one OCR cell is still a loss marker, not evidence that the cell should
        # be flipped back to profit.
        normalized_line = normalize_ocr_numbers(line)
        raw_amounts = AMOUNT_PATTERN.findall(normalized_line)
        signed = [-abs(number) for number in numbers]
        for index in range(1, len(signed)):
            previous_raw = raw_amounts[index - 1] if index - 1 < len(raw_amounts) else ""
            raw = raw_amounts[index] if index < len(raw_amounts) else ""
            previous_end = normalized_line.find(previous_raw) + len(previous_raw) if previous_raw else 0
            current_start = normalized_line.find(raw, previous_end) if raw else -1
            between = normalized_line[previous_end:current_start] if current_start >= previous_end else ""
            # Some legacy OCR uses a tilde to show that the comparison column
            # reverses from loss to profit.  Preserve that verified convention,
            # while a literal '-' alone remains a loss marker.
            if "~" in between or (
                "당기" in compact
                and "전기" in compact
                and raw.strip().startswith("-")
            ):
                signed[index] = abs(numbers[index])
        return signed
    # For mixed 손익 labels, retain the established OCR inference: a row whose
    # first cell is visibly a loss often lost the parentheses on later cells.
    normalized_line = normalize_ocr_numbers(line)
    aliases = ("당기순손익", "영업손익", "영업손", "당기순손")
    positions = [normalized_line.find(alias) for alias in aliases if alias in normalized_line]
    if positions:
        normalized_line = normalized_line[min(positions):]
    raw_amounts = AMOUNT_PATTERN.findall(normalized_line)
    signed = [-abs(numbers[0])]
    first_parenthesized = bool(raw_amounts and raw_amounts[0].strip().startswith("("))
    for index, number in enumerate(numbers[1:], start=1):
        raw = raw_amounts[index] if index < len(raw_amounts) else ""
        previous_raw = raw_amounts[index - 1] if index - 1 < len(raw_amounts) else ""
        previous_end = normalized_line.find(previous_raw) + len(previous_raw) if previous_raw else 0
        current_start = normalized_line.find(raw, previous_end) if raw else -1
        between = normalized_line[previous_end:current_start] if current_start >= previous_end else ""
        explicit_profit = (
            first_parenthesized and not raw.strip().startswith(("-", "("))
            or "~" in between
        )
        signed.append(abs(number) if explicit_profit else -abs(number))
    return signed


def find_period_net_summary(lines: list[str]) -> tuple[str, list[int]] | None:
    period_pattern = (
        r"{period}\s*[^0-9가-힣A-Za-z(+-]{{0,6}}\s*"
        r"(\(?-?(?:\d{{1,3}}(?:,\d{{3}})+|0)\)?)"
        r"(?:\s*원)?"
    )
    net_labels = ("당기순이익", "당기순손실", "분기순이익", "분기순손실", "결손금")
    for index, line in enumerate(lines):
        if not any(label in normalize(line) for label in net_labels):
            continue
        # Balance-sheet summaries often print the prior-period amount several
        # lines *before* the net-income label and the current amount after it.
        # Search symmetrically around the label instead of only the next rows.
        combined = " ".join(lines[max(0, index - 8) : min(len(lines), index + 9)])
        current_match = re.search(
            period_pattern.format(period=r"[당담탕]\s*[기키]"),
            combined,
        )
        previous_match = re.search(
            period_pattern.format(period=r"전\s*기"),
            combined,
        )
        if not current_match or not previous_match:
            continue
        current = parse_money(current_match.group(1))
        previous = parse_money(previous_match.group(1))
        if current is None or previous is None:
            continue
        context = normalize(combined)
        explicit_profit = "당기순이익" in context or "분기순이익" in context
        if not explicit_profit and (
            "당기순손실" in context
            or "분기순손실" in context
            or "결손" in context
        ):
            current = -abs(current)
            previous = -abs(previous)
            tilde_reversal = any(
                "~" in candidate
                and abs(previous) in [abs(value) for value in extract_amounts(candidate)]
                and any(label in normalize(candidate) for label in ("당기순손실", "분기순손실"))
                for candidate in lines
            )
            statement_reversal = any(
                "전기" not in normalize(window)
                and -abs(previous) in extract_amounts(window)
                for statement_index, candidate in enumerate(lines)
                if any(label in normalize(candidate) for label in ("당기순손실", "분기순손실"))
                for window in [" ".join(lines[statement_index : statement_index + 3])]
            )
            if tilde_reversal or statement_reversal:
                previous = abs(previous)
        return combined, [current, previous]
    return None


def calculated_net_income(lines: list[str]) -> tuple[str, int | float | None, int | float | None] | None:
    operating = extract_operating_profit(lines, "영업이익")
    if not operating:
        return None

    def period_values(aliases: tuple[str, ...], *, absolute: bool = False):
        match = find_alias_segment(lines, aliases, require_two=False)
        if not match:
            match = find_context_candidate(lines, aliases)
        if not match:
            return "", None, None
        numbers = financial_amounts(match, aliases)
        current, previous = pick_income_current_previous(numbers, lines)
        if absolute:
            current = abs(current) if current is not None else None
            previous = abs(previous) if previous is not None else None
        return match, current, previous

    income_line, income_current, income_previous = period_values(("영업외수익",))
    expense_line, expense_current, expense_previous = period_values(
        ("영업외비용", "영레외비용", "영외비용"),
        absolute=True,
    )
    # Preserve a negative tax amount.  It represents a tax benefit and must be
    # added back when reconciling pretax income to net income.
    tax_line, tax_current, tax_previous = period_values(("법인세등", "법인세비용"))
    if income_current is None or expense_current is None:
        return None
    tax_current = tax_current or 0
    current = operating.get("current")
    if current in (None, ""):
        return None
    current = current + income_current - expense_current - tax_current

    previous = operating.get("previous")
    if previous in (None, "") or income_previous is None or expense_previous is None:
        previous = None
    else:
        previous = previous + income_previous - expense_previous - (tax_previous or 0)
    source = " / ".join(filter(None, (operating.get("sourceLine", ""), income_line, expense_line, tax_line)))
    return source, current, previous


INCOME_STATEMENT_HEADINGS = ("포괄손익계산서", "손익계산서", "익계산서")
INCOME_STATEMENT_ENDINGS = (
    "제조원가명세서",
    "현금흐름표",
    "자본변동표",
    "결손금처리계산서",
    "이익잉여금처분계산서",
)
INCOME_STATEMENT_BODY_TERMS = (
    "매출액",
    "영업수익",
    "매출원가",
    "매출총이익",
    "판매비와관리비",
    "영업이익",
    "영업손실",
    "당기순이익",
    "당기순손실",
    "법인세차감전",
)


def income_statement_start(lines: list[str]) -> int:
    """Choose the actual statement body instead of a table-of-contents entry."""
    candidates = [
        index
        for index, line in enumerate(lines)
        if any(term in normalize(line) for term in INCOME_STATEMENT_HEADINGS)
    ]
    if not candidates:
        return -1

    def has_financial_body(start: int) -> bool:
        seen_terms: set[str] = set()
        for index in range(start + 1, min(len(lines), start + 180)):
            compact = normalize(lines[index])
            if any(term in compact for term in INCOME_STATEMENT_ENDINGS) or any(
                term in compact for term in INCOME_STATEMENT_HEADINGS
            ):
                break
            matched = [term for term in INCOME_STATEMENT_BODY_TERMS if term in compact]
            if matched and extract_amounts(lines[index]):
                seen_terms.update(matched)
                if len(seen_terms) >= 2:
                    return True
        return False

    # Preserve the original first-statement behaviour and skip only headings
    # that have no financial rows before the next heading (typically a TOC).
    valid_start = next((index for index in candidates if has_financial_body(index)), None)
    if valid_start is not None:
        return valid_start

    # Some spreadsheet-to-PDF exports print the title at the bottom, after the
    # actual statement body.  If such a trailing title has at least two distinct
    # financial row types immediately before it, start at the first body row.
    trailing = candidates[-1]
    if trailing >= max(0, len(lines) - 5):
        body_indexes: list[int] = []
        seen_terms: set[str] = set()
        for index in range(max(0, trailing - 180), trailing):
            compact = normalize(lines[index])
            matched = [term for term in INCOME_STATEMENT_BODY_TERMS if term in compact]
            if matched and extract_amounts(lines[index]):
                body_indexes.append(index)
                seen_terms.update(matched)
        if body_indexes and len(seen_terms) >= 2:
            return min(body_indexes)
    return candidates[0]


def income_statement_scope(lines: list[str]) -> list[str]:
    start = income_statement_start(lines)
    if start < 0:
        return lines
    end = len(lines)
    for index in range(start + 1, len(lines)):
        compact = normalize(lines[index])
        if any(term in compact for term in INCOME_STATEMENT_ENDINGS):
            end = index
            break
    return lines[start:end]


def zero_tax_pretax_income(lines: list[str]) -> tuple[str, int | float | None, int | float | None] | None:
    aliases = (
        "법인세비용차감전손익",
        "법인세비용차감전손실",
        "법인세비용차감전순이익",
        "법인세비용차감전순손실",
        "법인세차감전순이",
        "법인세차감전순이익",
        "법인세차감전순손실",
        "법인세차감전손익",
        "법인세차감전손실",
        "법인세차김전순손실",
        "인세차감전손실",
    )
    match = find_cumulative_alias_segment(lines, aliases)
    if not match:
        match = find_vertical_alias_segment(lines, aliases, amount_count=2)
    if not match:
        match = find_alias_segment(lines, aliases, require_two=False)
    if not match:
        match = find_context_candidate(lines, aliases)
    if not match:
        return None
    tax_candidates = [
        line
        for line in lines
        if "차감전" not in normalize(line)
        and any(alias in normalize(line) for alias in ("법인세등", "법인세비용"))
        and extract_amounts(line)
    ]
    tax_match = min(tax_candidates, key=len, default="")
    tax_numbers = financial_amounts(tax_match, ("법인세등", "법인세비용")) if tax_match else []
    if tax_numbers and any(value != 0 for value in tax_numbers):
        return None
    numbers = financial_amounts(match, aliases)
    compact = normalize(match)
    if "손실" in compact:
        numbers = [-abs(value) for value in numbers]
    current, previous = pick_income_current_previous(numbers, lines)
    if has_single_period_income_layout(lines):
        previous = None if (
            is_first_fiscal_period_income_layout(lines)
            or has_trailing_income_statement_title(lines)
        ) else 0
    return match, current, previous


def extract_net_income(lines: list[str], label: str) -> dict | None:
    all_lines = lines
    single_period = has_single_period_income_layout(all_lines)
    first_fiscal_period = is_first_fiscal_period_income_layout(all_lines)
    missing_comparison_period = (
        first_fiscal_period or has_trailing_income_statement_title(all_lines)
    )
    lines = income_statement_scope(lines)
    aliases = (
        "당기순손실",
        "당기순이익",
        "당기손익",
        "분기순손실",
        "분기순이익",
        "반기순손익",
    )
    excluded_net_terms = (
        "기타포괄",
        "재분류",
        "재측정",
        "공정가치측정",
        "금융자산평가",
        "금융부채평가",
        "평가손익",
        "이익잉여금",
        "항목의법인세",
    )

    def has_excluded_net_context(line: str) -> bool:
        compact = normalize(line)
        alias_positions = [compact.find(alias) for alias in aliases if alias in compact]
        if not alias_positions:
            return any(term in compact for term in excluded_net_terms)
        # In a one-line OCR page, an earlier 이익잉여금 heading must not hide a
        # later explicit 당기순이익 row.  Exclusions only apply from the target
        # label onward (for example 당기손익-공정가치측정금융자산).
        target_tail = compact[min(alias_positions) :]
        return any(term in target_tail for term in excluded_net_terms)

    net_lines = [
        line
        for line in lines
        if not has_excluded_net_context(line)
    ]
    all_net_lines = [
        line
        for line in all_lines
        if not has_excluded_net_context(line)
    ]
    period_summary = find_period_net_summary(all_lines)

    def confirmed_period_summary(
        candidate_current: int | float | None,
    ) -> tuple[str, int | float | None, int | float | None] | None:
        if not period_summary or candidate_current is None:
            return None
        summary_source, summary_numbers = period_summary
        summary_current, summary_previous = pick_current_previous(summary_numbers)
        if summary_current == candidate_current:
            return summary_source, summary_current, summary_previous
        return None

    match = find_cumulative_alias_segment(net_lines, aliases)
    if not match:
        match = find_quality_alias_segment(
        net_lines,
        aliases,
        require_two=True,
        )
    if not match:
        match = find_context_candidate(
        net_lines,
        ("당기순손실", "당기순이익", "당기손익", "당기순손", "순이익"),
        exclude=("제조원가", "매출원가", "원가"),
        )
    pretax = zero_tax_pretax_income(lines)
    expected_amount_count = 4 if has_cumulative_income_layout(all_lines) else 2
    match_amounts = financial_amounts(match, aliases) if match else []
    has_reference_number = (
        any(1_000 <= abs(value) < 10_000 for value in match_amounts)
        and any(abs(value) >= 1_000_000 for value in match_amounts)
    )
    match_is_noisy = bool(match) and (
        len(match) > 500
        or len(match_amounts) < expected_amount_count
        or len(match_amounts) > expected_amount_count
        or has_reference_number
    )
    if pretax and period_summary:
        _pretax_source, pretax_current, pretax_previous = pretax
        summary_source, summary_numbers = period_summary
        summary_current, summary_previous = pick_current_previous(summary_numbers)
        pretax_is_incomplete = (
            pretax_current in (None, 0)
            and summary_current not in (None, 0)
        ) or (
            pretax_previous is None
            and summary_previous is not None
        )
        summary_confirms_current = (
            pretax_current not in (None, 0)
            and summary_current not in (None, 0)
            and pretax_current == summary_current
            and (
                pretax_previous in (None, 0)
                or summary_previous in (None, 0)
                or abs(pretax_previous - summary_previous)
                <= max(1_000, max(abs(pretax_previous), abs(summary_previous)) * 0.05)
            )
        )
        if pretax_is_incomplete or summary_confirms_current:
            return {
                "item": label,
                "current": summary_current if summary_current is not None else "",
                "previous": summary_previous if summary_previous is not None else "",
                "status": "당기/전기 요약 금액 사용",
                "sourceLine": summary_source,
            }
    if pretax and (not match or match_is_noisy):
        source_line, current, previous = pretax
        return {
            "item": label,
            "current": current if current is not None else "",
            "previous": previous if previous is not None else "",
            "status": "법인세 0: 세전손익 사용",
            "sourceLine": source_line,
        }
    document_match = find_quality_alias_segment(all_net_lines, aliases, require_two=True)
    if document_match and match_is_noisy:
        # A clean row elsewhere in the same report is safer than a synthetic
        # three-line context containing reference numbers and unrelated
        # accounts.  This commonly occurs when the primary statement is a
        # low-quality scan but the deficit schedule repeats net income clearly.
        match = document_match
    if not match and pretax:
        source_line, current, previous = pretax
        return {
            "item": label,
            "current": current if current is not None else "",
            "previous": previous if previous is not None else "",
            "status": "법인세 0: 세전손익 사용",
            "sourceLine": source_line,
        }
    if not match:
        calculated = calculated_net_income(lines)
        if calculated:
            source_line, current, previous = calculated
            confirmed_summary = confirmed_period_summary(current)
            if confirmed_summary:
                source_line, current, previous = confirmed_summary
                status = "당기/전기 요약 금액 사용"
            else:
                status = "영업손익과 영업외손익으로 계산"
            return {
                "item": label,
                "current": current if current is not None else "",
                "previous": previous if previous is not None else "",
                "status": status,
                "sourceLine": source_line,
            }
        if document_match:
            match = document_match
    if not match or len(extract_amounts(match)) < 2:
        start = next((index for index, line in enumerate(lines) if "손익계산서" in normalize(line) or "익계산서" in normalize(line)), -1)
        if start >= 0:
            candidates = []
            for line in lines[start:]:
                compact = normalize(line)
                line_amounts = extract_amounts(line)
                if (
                    line_amounts
                    and any(value == 0 or abs(value) >= 10_000 for value in line_amounts)
                    and "법인세" not in compact
                    and ("당기" in compact or "순" in compact or "손익" in compact or compact.startswith("XS"))
                ):
                    candidates.append(line)
            if candidates:
                match = candidates[-1]
    if not match and period_summary:
        match, numbers = period_summary
        current, previous = pick_current_previous(numbers)
        return {
            "item": label,
            "current": current if current is not None else "",
            "previous": previous if previous is not None else "",
            "status": "당기/전기 요약 금액 사용",
            "sourceLine": match,
        }
    if not match:
        calculated = calculated_net_income(lines)
        if not calculated:
            return None
        source_line, current, previous = calculated
        confirmed_summary = confirmed_period_summary(current)
        if confirmed_summary:
            source_line, current, previous = confirmed_summary
            status = "당기/전기 요약 금액 사용"
        else:
            status = "영업손익과 영업외손익으로 계산"
        return {
            "item": label,
            "current": current if current is not None else "",
            "previous": previous if previous is not None else "",
            "status": status,
            "sourceLine": source_line,
        }

    compact_match = normalize(match)
    is_mixed_profit_loss = "손익" in compact_match or "이익" in compact_match and "손실" in compact_match
    is_loss = ("손실" in compact_match or "손" in compact_match) and not is_mixed_profit_loss
    numbers = financial_amounts(match, ("당기순손실", "당기순이익", "당기손익", "당기순손", "분기순이익", "반기순손익", "순이익", "법인세차감전"))
    if is_loss:
        numbers = apply_loss_row_signs(match, numbers)
    elif "순손익" in compact_match and len(numbers) >= 2 and numbers[0] < 0 < numbers[1]:
        numbers = apply_loss_row_signs(match, numbers)
    current, previous = pick_income_current_previous(numbers, lines)
    if single_period:
        previous = None if missing_comparison_period else 0
    if pretax:
        pretax_source, pretax_current, pretax_previous = pretax
        suspicious_current = current in (None, 0) or (
            pretax_current not in (None, 0)
            and current not in (None, 0)
            and (
                current == -pretax_current
                or max(abs(current), abs(pretax_current)) / min(abs(current), abs(pretax_current)) > 3
            )
        )
        suspicious_previous = previous in (None, 0) or (
            pretax_previous not in (None, 0)
            and previous not in (None, 0)
            and (
                previous == -pretax_previous
                or max(abs(previous), abs(pretax_previous)) / min(abs(previous), abs(pretax_previous)) > 3
            )
        )
        if suspicious_current:
            match = pretax_source
            current = pretax_current
        if suspicious_previous:
            match = pretax_source
            previous = pretax_previous
    if period_summary and (
        current is None
        or previous is None
        or abs(current) < 10_000
        or abs(previous) < 10_000
        or current == previous
    ):
        match, numbers = period_summary
        current, previous = pick_current_previous(numbers)
    calculated = calculated_net_income(lines)
    if calculated:
        calculated_source, calculated_current, calculated_previous = calculated
        suspicious_year_value = current is not None and 1900 <= abs(current) <= 2100
        if suspicious_year_value:
            match = calculated_source
            current, previous = calculated_current, calculated_previous
        else:
            tolerance = lambda value: max(100, abs(value) * 1e-8)
            if (
                current is not None
                and calculated_current is not None
                and abs(current - calculated_current) <= tolerance(calculated_current)
            ):
                current = calculated_current
            if (
                previous is not None
                and calculated_previous is not None
                and abs(previous - calculated_previous) <= tolerance(calculated_previous)
            ):
                previous = calculated_previous
    if single_period:
        previous = None if missing_comparison_period else 0
    return {
        "item": label,
        "current": current if current is not None else "",
        "previous": previous if previous is not None else "",
        "status": "확인 필요" if current is not None or previous is not None else "금액 없음",
        "sourceLine": match,
    }


def operating_profit_from_components(lines: list[str]) -> tuple[str, int | float | None, int | float | None] | None:
    gross_match = find_alias_segment(
        lines,
        ("매출총이익", "매출총손실"),
        require_two=False,
    )
    selling_match = find_alias_segment(
        lines,
        ("판매비와관리비",),
        require_two=False,
    )
    if not gross_match or not selling_match:
        return None
    gross = financial_amounts(gross_match, ("매출총이익", "매출총손실"))
    selling = financial_amounts(selling_match, ("판매비와관리비",))
    gross_current, gross_previous = pick_income_current_previous(gross, lines)
    selling_current, selling_previous = pick_income_current_previous(selling, lines)
    if "매출총손실" in normalize(gross_match):
        gross_current = -abs(gross_current) if gross_current is not None else None
        gross_previous = -abs(gross_previous) if gross_previous is not None else None
    current = (
        gross_current - selling_current
        if gross_current is not None and selling_current is not None
        else None
    )
    previous = (
        gross_previous - selling_previous
        if gross_previous is not None and selling_previous is not None
        else None
    )
    return f"{gross_match} / {selling_match}", current, previous


def extract_operating_profit(lines: list[str], label: str) -> dict | None:
    all_lines = lines
    scoped_lines = income_statement_scope(all_lines)
    single_period = has_single_period_income_layout(all_lines)
    aliases = ("영업손실", "영업이익", "영업손익")
    # Prefer the audited statement body when it can be identified, but retain
    # the historical whole-document fallbacks for poorly headed OCR files.
    exact_match = find_cumulative_alias_segment(scoped_lines, aliases)
    if not exact_match:
        exact_match = find_alias_segment(
            scoped_lines,
            ("영업손실", "영업이익"),
            prefer_last=True,
            require_two=True,
        )
    lines = all_lines
    if not exact_match:
        exact_match = find_cumulative_alias_segment(lines, aliases)
    if not exact_match:
        exact_match = find_alias_segment(
            lines,
            ("영업손실", "영업이익"),
            prefer_last=True,
            require_two=True,
        )
    if not exact_match:
        exact_match = find_context_candidate(lines, ("영업손실", "영업이익"))
    match = exact_match or find_context_candidate(
        lines,
        ("영업손실", "영업이익", "영업손익", "영업손"),
        exclude=("영업외", "매출총", "제조원가", "원가"),
    )
    if not match:
        calculated = operating_profit_from_components(lines)
        if calculated:
            source_line, current, previous = calculated
            return {
                "item": label,
                "current": current if current is not None else "",
                "previous": previous if previous is not None else "",
                "status": "매출총손익 - 판매비와관리비 계산",
                "sourceLine": source_line,
            }
        return None
    compact = normalize(match)
    numbers = financial_amounts(match, ("영업손실", "영업이익", "영업손익", "영업손"))
    is_mixed_profit_loss = "이익" in compact and "손실" in compact
    if not is_mixed_profit_loss and ("손실" in compact or "영업손" in compact and "영업손익" not in compact):
        numbers = apply_loss_row_signs(match, numbers)
    elif "영업손익" in compact and len(numbers) >= 2 and numbers[0] < 0 < numbers[1]:
        numbers = apply_loss_row_signs(match, numbers)
    current, previous = pick_income_current_previous(numbers, lines)
    if single_period:
        previous = None
    calculated = operating_profit_from_components(lines)
    if calculated:
        _source_line, calculated_current, calculated_previous = calculated
        tolerance = lambda value: max(100, abs(value) * 1e-8)
        if (
            current is not None
            and calculated_current is not None
            and abs(current - calculated_current) <= tolerance(calculated_current)
        ):
            current = calculated_current
        if (
            previous is not None
            and calculated_previous is not None
            and abs(previous - calculated_previous) <= tolerance(calculated_previous)
        ):
            previous = calculated_previous
    if single_period:
        previous = None
    return {
        "item": label,
        "current": current if current is not None else "",
        "previous": previous if previous is not None else "",
        "status": "확인 필요" if current is not None or previous is not None else "금액 없음",
        "sourceLine": match,
    }


def find_context_candidate(lines: list[str], aliases: tuple[str, ...], exclude: tuple[str, ...] = ()) -> str:
    candidates = []
    for index, line in enumerate(lines):
        compact = normalize(line)
        if not any(alias in compact for alias in aliases) or any(term in compact for term in exclude):
            continue
        combined = line
        if len(financial_amounts(combined, aliases)) < 2:
            combined = " ".join(lines[index : index + 3])
        if financial_amounts(combined, aliases):
            candidates.append(combined)
    two_period_candidates = [
        candidate for candidate in candidates
        if len(financial_amounts(candidate, aliases)) >= 2
    ]
    if two_period_candidates:
        return two_period_candidates[-1]
    return candidates[0] if candidates else ""


def slice_alias_amounts(line: str, alias: str, count: int = 2) -> str:
    normalized_line = normalize_ocr_numbers(line)
    pattern = r"[ \t]*".join(map(re.escape, alias))
    label_match = re.search(pattern, normalized_line)
    if not label_match:
        return ""
    if normalized_line[label_match.end():].lstrip().startswith("원가"):
        return ""
    tail = normalized_line[label_match.start():]
    amount_tail = re.sub(r"(?<![\d,])\d{1,3}\.\d{1,2}(?!\d)", " ", tail)
    amounts = list(AMOUNT_PATTERN.finditer(amount_tail))
    if not amounts:
        return ""
    text_before_amount = normalize(amount_tail[len(label_match.group(0)):amounts[0].start()])
    other_labels = (
        "자산총계",
        "부채총계",
        "부채와자본총계",
        "부채및자본총계",
        "자본금",
        "매출액",
        "매출원가",
        "매출총이익",
        "판매비와관리비",
        "영업수익",
        "영업이익",
        "영업손실",
        "당기순이익",
        "당기순손실",
        "법인세비용차감전",
    )
    if any(other in text_before_amount for other in other_labels):
        return ""
    effective_count = count
    if len(amounts) >= 2:
        between_periods = normalize(amount_tail[amounts[0].end() : amounts[1].start()])
        row_boundary_labels = (
            "자산총계",
            "부채총계",
            "부채와자본총계",
            "부채및자본총계",
            "자본총계",
            "자본금",
            "기타자본",
            "자본잉여금",
            "주식발행초과금",
            "보통주자본금",
            "우선주자본금",
            "자본조정",
            "유동자산",
            "비유동자산",
            "유동부채",
            "비유동부채",
            "매출액",
            "매출원가",
            "매출총이익",
            "영업수익",
            "영업이익",
            "영업손실",
            "당기순이익",
            "당기순손실",
            "법인세비용차감전",
        )
        if any(boundary in between_periods for boundary in row_boundary_labels):
            effective_count = 1
        elif re.search(
            r"(?:비용|급여|소모품비|지급수수료|감가상각|임차료|보험료|세금과공과|통신비|운반비)",
            between_periods,
        ):
            effective_count = 1
    if effective_count == count and count == 2 and len(amounts) >= 4:
        parsed = [parse_money(match.group(0)) for match in amounts[:4]]
        if (
            all(value is not None for value in parsed)
            and abs(parsed[0]) >= 1_000
            and abs(parsed[2]) >= 1_000
            and abs(parsed[1]) <= 100
            and abs(parsed[3]) <= 100
        ):
            # Many management reports place a percentage after each amount.
            # Keep both amount/percentage pairs so financial_amounts can remove
            # the percentage columns without losing the previous-period amount.
            effective_count = 4
    end = amounts[min(effective_count, len(amounts)) - 1].end()
    return amount_tail[:end]


def find_alias_segment(
    lines: list[str],
    aliases: tuple[str, ...],
    *,
    prefer_last: bool = False,
    require_two: bool = False,
) -> str:
    segments = []
    for line in lines:
        compact = normalize(line)
        for alias in aliases:
            if alias not in compact:
                continue
            segment = slice_alias_amounts(line, alias)
            amount_count = len(extract_amounts(segment))
            if amount_count and (not require_two or amount_count >= 2):
                segments.append(segment)
    if not segments:
        return ""
    return segments[-1] if prefer_last else segments[0]


def find_quality_alias_segment(
    lines: list[str],
    aliases: tuple[str, ...],
    *,
    require_two: bool = False,
) -> str:
    segments: list[str] = []
    for line in lines:
        compact = normalize(line)
        for alias in aliases:
            if alias not in compact:
                continue
            segment = slice_alias_amounts(line, alias)
            amount_count = len(extract_amounts(segment))
            if amount_count and (not require_two or amount_count >= 2):
                segments.append(segment)
    if not segments:
        return ""

    parsed = [extract_amounts(segment)[:2] for segment in segments]
    current_frequency: dict[int, int] = {}
    previous_frequency: dict[int, int] = {}
    for amounts in parsed:
        if amounts:
            current_frequency[abs(amounts[0])] = current_frequency.get(abs(amounts[0]), 0) + 1
        if len(amounts) >= 2:
            previous_frequency[abs(amounts[1])] = previous_frequency.get(abs(amounts[1]), 0) + 1

    def quality(segment: str) -> tuple[int, int, int]:
        amounts = extract_amounts(segment)
        raw_amounts = AMOUNT_PATTERN.findall(segment)
        comma_grouped = sum(
            bool(re.fullmatch(r"\(?-?\d{1,3}(?:,\d{3})+\)?", raw.strip()))
            for raw in raw_amounts[:2]
        )
        repeated = 0
        if amounts:
            repeated += current_frequency.get(abs(amounts[0]), 0)
        if len(amounts) >= 2:
            repeated += previous_frequency.get(abs(amounts[1]), 0)
        return comma_grouped, repeated, segments.index(segment)

    return max(segments, key=quality)


def extract_revenue(lines: list[str], label: str) -> dict | None:
    all_lines = lines
    single_period = has_single_period_income_layout(all_lines)
    aliases = ("매출액", "배출액", "영업수익")
    scoped_lines = income_statement_scope(all_lines)
    scoped_match = ""
    if income_statement_start(all_lines) >= 0:
        scoped_match = find_cumulative_alias_segment(scoped_lines, aliases)
        if not scoped_match:
            scoped_match = find_quality_alias_segment(scoped_lines, aliases, require_two=True)
        if not scoped_match:
            scoped_totals = [
                line
                for line in scoped_lines
                if len(extract_amounts(line)) >= 2
                and "매출" in normalize(line)
                and not is_revenue_component_line(line)
                and not is_revenue_cost_line(line)
            ]
            if scoped_totals:
                scoped_match = min(scoped_totals, key=revenue_candidate_rank)
    income_start = next(
        (
            index for index, line in enumerate(all_lines)
            if any(term in normalize(line) for term in ("손익계산서", "부채와자본총계", "부채및자본총계"))
        ),
        0,
    )
    lines = all_lines[income_start:]
    single_direct_match = find_quality_alias_segment(lines, aliases, require_two=False)
    direct_match = find_cumulative_alias_segment(lines, aliases)
    if not direct_match:
        direct_match = find_quality_alias_segment(lines, aliases, require_two=True)
    if not direct_match:
        direct_match = find_context_candidate(lines, ("매출액", "배출액", "영업수익"), exclude=("원가",))
    if not direct_match:
        direct_match = next(
            (
                segment
                for line in lines
                if (segment := slice_alias_amounts(line, "제품매출"))
                and len(extract_amounts(segment)) >= 1
            ),
            "",
        )
    revenue_candidates = [
        line for line in lines
        if len(extract_amounts(line)) >= 2
        and any(hint in normalize(line) for hint in ("매출액", "출액", "영업수익", "수수료수익", "제품매출"))
        and not is_revenue_cost_line(line)
    ]
    match = scoped_match or direct_match or (min(revenue_candidates, key=revenue_candidate_rank) if revenue_candidates else "")
    if not match or revenue_candidate_rank(match)[0] >= 2:
        generic_totals = [
            line for line in lines
            if len(extract_amounts(line)) >= 2
            and "매출" in normalize(line)
            and not is_revenue_component_line(line)
            and not is_revenue_cost_line(line)
        ]
        if generic_totals:
            match = min(generic_totals, key=revenue_candidate_rank)
    if not match or revenue_candidate_rank(match)[0] >= 2:
        component_lines = []
        for line in all_lines:
            compact = normalize(line)
            compact_letters = re.sub(r"[^가-힣A-Za-z0-9]", "", line)
            if is_revenue_cost_line(line) or "매출원가" in compact:
                if component_lines:
                    break
                continue
            if is_revenue_component_line(line):
                component_lines.append(line)
            elif component_lines and "매출" in compact_letters and len(extract_amounts(line)) == 1 and not is_revenue_cost_line(line):
                component_lines.append(line)
        current_parts = []
        previous_parts = []
        single_parts = []
        for line in dict.fromkeys(component_lines):
            amounts = financial_amounts(line, ("제품매출", "상품매출", "용역매출", "서비스매출"))
            if len(amounts) >= 2:
                current_parts.append(amounts[0])
                previous_parts.append(amounts[1])
            elif len(amounts) == 1:
                single_parts.append(amounts[0])
        if current_parts:
            current = sum(current_parts)
            previous = sum(previous_parts)
            if single_parts and previous > abs(current) * 2:
                previous += sum(single_parts)
            selling_match = find_context_candidate(lines, ("판매비와관리비",))
            cost_match = find_context_candidate(lines, ("매출원가",))
            operating_row = extract_operating_profit(lines, "영업이익")
            selling_numbers = financial_amounts(selling_match, ("판매비와관리비",)) if selling_match else []
            cost_numbers = financial_amounts(cost_match, ("매출원가",)) if cost_match else []
            if operating_row and selling_numbers and cost_numbers:
                selling_current, selling_previous = pick_income_current_previous(selling_numbers, lines)
                cost_current, cost_previous = pick_income_current_previous(cost_numbers, lines)
                operating_current = operating_row.get("current")
                operating_previous = operating_row.get("previous")
                calculated_current = (
                    operating_current + abs(selling_current) + abs(cost_current)
                    if operating_current not in (None, "") and selling_current is not None and cost_current is not None
                    else None
                )
                calculated_previous = (
                    operating_previous + abs(selling_previous) + abs(cost_previous)
                    if operating_previous not in (None, "") and selling_previous is not None and cost_previous is not None
                    else None
                )
                if (
                    calculated_current is not None
                    and current == calculated_current
                    and (previous is None or abs(previous) < max(1, abs(current)) * 0.01)
                ):
                    previous = calculated_previous
            return {
                "item": label,
                "current": current,
                "previous": previous,
                "status": "구성 매출 합산",
                "sourceLine": " / ".join(component_lines),
            }
    if not match or len(extract_amounts(match)) < 2:
        start = next((index for index, line in enumerate(lines) if "손익계산서" in normalize(line) or "익계산서" in normalize(line)), -1)
        if start >= 0:
            header_seen = False
            for line in lines[start : start + 35]:
                compact = normalize(line)
                if "금액" in compact or "액" == compact:
                    header_seen = True
                    continue
                if not header_seen and not any(hint in compact for hint in ("매출", "수수료", "영업수익", "제품매출", "출액")):
                    continue
                if (
                    extract_amounts(line)
                    and len(extract_amounts(line)) >= 2
                    and any(hint in compact for hint in ("매출액", "출액", "수수료", "영업수익", "제품매출", "매출"))
                    and not is_revenue_cost_line(line)
                    and not ("매출" in compact and "가" in compact and "액" not in compact and "수수료" not in compact)
                ):
                    match = line
                    break
    if not match:
        return None

    numbers = financial_amounts(match, ("매출액", "배출액", "출액", "영업수익", "수수료", "제품매출", "매출"))
    current, previous = pick_income_current_previous(numbers, lines)
    if single_period:
        previous = None
    raw_match_amounts = extract_amounts(match)
    if (
        len(raw_match_amounts) >= 3
        and raw_match_amounts[-2:] == [0, 0]
        and abs(raw_match_amounts[0]) < 1_000_000
        and "," not in AMOUNT_PATTERN.findall(match)[0]
    ):
        current, previous = 0, 0
    single_direct_numbers = financial_amounts(single_direct_match, aliases) if single_direct_match else []
    if len(single_direct_numbers) == 1 and current == previous == single_direct_numbers[0]:
        current, previous = single_direct_numbers[0], 0
    if current is not None:
        gross_profit = find_context_candidate(lines, ("매출총이익", "매출총손실"))
        gross_numbers = financial_amounts(gross_profit, ("매출총이익", "매출총손실"))
        if len(gross_numbers) >= 2:
            gross_current, gross_previous = pick_income_current_previous(gross_numbers, lines)
            if previous is None and gross_current is not None and abs(gross_current) == abs(current):
                previous = gross_previous
            elif (
                gross_current is not None
                and gross_previous is not None
                and previous is not None
                and abs(gross_previous) == abs(current)
                and abs(gross_current) == abs(previous)
                and current != previous
            ):
                # A one-line OCR fragment can place the previous-period revenue
                # before a following zero.  The gross-profit row confirms the
                # actual current/previous column order.
                current, previous = gross_current, gross_previous
        cost_match = find_context_candidate(lines, ("매출원가",))
        cost_numbers = financial_amounts(cost_match, ("매출원가",))
        if (
            cost_numbers
            and previous == 0
            and current is not None
            and abs(current) < 1_000_000
            and "~" in match
            and all(value == 0 for value in cost_numbers)
        ):
            current, previous = 0, 0
        if gross_numbers and cost_numbers:
            gross_current, gross_previous = pick_income_current_previous(gross_numbers, lines)
            cost_current, cost_previous = pick_income_current_previous(cost_numbers, lines)
            calculated_current = (
                gross_current + abs(cost_current)
                if gross_current is not None and cost_current is not None
                else None
            )
            calculated_previous = (
                gross_previous + abs(cost_previous)
                if gross_previous is not None and cost_previous is not None
                else None
            )
            current_confirms_identity = calculated_current is not None and current == calculated_current
            previous_confirms_identity = calculated_previous is not None and previous == calculated_previous
            if previous_confirms_identity or (
                calculated_current == 0
                and previous == 0
                and current is not None
                and abs(current) < 1_000_000
                and "," not in match
            ):
                current = calculated_current
            if (
                current_confirms_identity
                and (previous is None or abs(previous) < max(1, abs(current)) * 0.01)
            ) or (calculated_current == 0 and calculated_previous == 0):
                previous = calculated_previous
        selling_match = find_context_candidate(lines, ("판매비와관리비",))
        selling_numbers = financial_amounts(selling_match, ("판매비와관리비",)) if selling_match else []
        operating_row = extract_operating_profit(lines, "영업이익")
        if cost_numbers and selling_numbers and operating_row:
            cost_current, cost_previous = pick_income_current_previous(cost_numbers, lines)
            selling_current, selling_previous = pick_income_current_previous(selling_numbers, lines)
            operating_current = operating_row.get("current")
            operating_previous = operating_row.get("previous")
            component_current = (
                operating_current + abs(selling_current) + abs(cost_current)
                if operating_current not in (None, "") and selling_current is not None and cost_current is not None
                else None
            )
            component_previous = (
                operating_previous + abs(selling_previous) + abs(cost_previous)
                if operating_previous not in (None, "") and selling_previous is not None and cost_previous is not None
                else None
            )
            if (
                component_current is not None
                and current == component_current
                and (previous is None or abs(previous) < max(1, abs(current)) * 0.01)
            ):
                previous = component_previous
    if single_period:
        previous = None
    return {
        "item": label,
        "current": current if current is not None else "",
        "previous": previous if previous is not None else "",
        "status": "확인 필요" if current is not None or previous is not None else "금액 없음",
        "sourceLine": match,
    }


def is_revenue_cost_line(line: str) -> bool:
    condensed = re.sub(r"[^가-힣A-Za-z0-9]", "", line)
    fuzzy_cost = bool(re.search(r"매출.{0,2}가", condensed)) and "매출액" not in condensed
    return fuzzy_cost or any(skip in condensed for skip in ("매출원가", "제품매출원가", "매출총이익", "매출총손실", "매출채권", "외상매출금", "매출할인", "지급수수료")) or any(
        term in condensed for term in ("상각", "판매비", "비용")
    )


def is_revenue_component_line(line: str) -> bool:
    compact = normalize(line)
    return any(component in compact for component in (
        "제품매출",
        "상품매출",
        "용역매출",
        "서비스매출",
        "보조금매출",
        "행사매출",
        "음원매출",
        "기타매출",
        "연구개발매출",
        "설계용역매출",
        "광고대행매출",
        "콘텐츠제작매출",
    )) or bool(
        re.search(r"서비.{0,2}매출", compact)
    )


def revenue_candidate_rank(line: str) -> tuple[int, int]:
    compact = normalize(line)
    if "매출액" in compact:
        priority = 0
    elif re.search(r"(?:^|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ0-9.])매출(?:$|\d)", compact):
        priority = 0
    elif "매출" in compact and not is_revenue_component_line(line):
        priority = 0
    elif "영업수익" in compact or "수수료" in compact:
        priority = 1
    else:
        priority = 2
    return priority, len(line)


def find_best_line(lines: list[str], aliases: list[str], label: str) -> str:
    normalized_aliases = [normalize(alias) for alias in aliases]
    fallback = ""
    for index, line in enumerate(lines):
        compact = normalize(line)
        if any(alias in compact for alias in normalized_aliases):
            if label == "자본금" and "자본금계" in compact:
                continue
            if extract_amounts(line) and alias_before_first_amount(line, aliases):
                return line
            fallback = fallback or line
    return fallback


def alias_before_first_amount(line: str, aliases: list[str]) -> bool:
    first_amount = AMOUNT_PATTERN.search(line)
    if not first_amount:
        return False
    for alias in aliases:
        pattern = r"\s*".join(map(re.escape, alias.replace(" ", "")))
        match = re.search(pattern, line)
        if match and match.start() < first_amount.start():
            return True
    return False


def extract_amounts(line: str) -> list[int | float]:
    line = normalize_ocr_numbers(line)
    numbers = [parse_money(raw) for raw in AMOUNT_PATTERN.findall(line)]
    return [num for num in numbers if num is not None]


def financial_amounts(line: str, aliases: tuple[str, ...]) -> list[int | float]:
    normalized_line = normalize_ocr_numbers(line)
    normalized_line = re.sub(r"(?<![\d,])\d{1,3}\.\d{1,2}(?!\d)", " ", normalized_line)
    label_matches = []
    for alias in aliases:
        pattern = r"[ \t]*".join(map(re.escape, alias.replace(" ", "")))
        match = re.search(pattern, normalized_line)
        if match:
            label_matches.append(match)
    if label_matches:
        normalized_line = normalized_line[min(match.start() for match in label_matches):]
    numbers = extract_amounts(normalized_line)
    if len(numbers) == 1:
        short_previous = re.search(
            r"(?:\d{1,3}(?:,\d{3})+|\d{4,})[ \t]+(\(?-?\d{1,3}\)?)\s*$",
            normalized_line,
        )
        if short_previous:
            raw = short_previous.group(1)
            short_value = int(raw.strip("()-"))
            if raw.startswith(("-", "(")):
                short_value = -short_value
            numbers.append(short_value)
    if len(numbers) >= 3 and abs(numbers[0]) < max(1, abs(numbers[1])) * 0.01:
        numbers = numbers[1:]
    if (
        len(numbers) >= 4
        and abs(numbers[0]) >= 1_000
        and abs(numbers[2]) >= 1_000
        and abs(numbers[1]) <= 100
        and abs(numbers[3]) <= 100
    ):
        numbers = [numbers[0], numbers[2], *numbers[4:]]
    return numbers


def normalize(value: str) -> str:
    return re.sub(r"[\s()]", "", value)


def parse_money(raw: str) -> int | float | None:
    compact = re.sub(r"\s+", "", raw)
    unsigned = compact.strip("()-")
    digits_only = re.sub(r"[^0-9]", "", unsigned)
    if digits_only != "0" and "," not in unsigned and "." not in unsigned and len(digits_only) < 4:
        return None
    sign = -1 if compact.startswith("-") or compact.startswith("(") and compact.endswith(")") else 1
    if not digits_only:
        return None
    if (
        "," in unsigned
        and "." in unsigned
        and re.fullmatch(r"\d{1,3}(?:[,.]\d{3}){2,}", unsigned)
    ):
        # OCR occasionally substitutes the final thousands comma with a dot
        # (for example 19,309,036.569). Financial statement totals are won
        # integers, so preserve the full magnitude rather than parsing a float.
        value = float(unsigned.replace(",", "").replace(".", ""))
    elif "," in unsigned and "." in unsigned:
        value = float(unsigned.replace(",", ""))
    elif "." in unsigned and not re.fullmatch(r"\d{1,3}(?:\.\d{3})+", unsigned):
        value = float(unsigned)
    else:
        value = float(unsigned.replace(",", "").replace(".", ""))
    signed = value * sign
    return int(signed) if signed.is_integer() else signed


def collapse_ocr_repetitions(numbers: list[int | float]) -> list[int | float]:
    if len(numbers) < 4:
        return numbers
    runs: list[tuple[int | float, int]] = []
    for number in numbers:
        if runs and runs[-1][0] == number:
            runs[-1] = (number, runs[-1][1] + 1)
        else:
            runs.append((number, 1))
    if len(runs) >= 2 and all(count >= 2 for _value, count in runs[:2]):
        return [value for value, _count in runs]
    return numbers


def pick_current_previous(numbers: list[int | float]) -> tuple[int | float | None, int | float | None]:
    numbers = collapse_ocr_repetitions(numbers)
    if len(numbers) >= 4 and numbers[0] == 0 and numbers[2] == 0:
        return numbers[1], numbers[3]
    if len(numbers) >= 3:
        return numbers[0], numbers[1]
    if len(numbers) >= 2:
        return numbers[0], numbers[-1]
    if len(numbers) == 1:
        return numbers[0], None
    return None, None


def build_workbook(path: Path, documents: list[dict]):
    wb = Workbook()
    wb.remove(wb.active)
    header_fill = PatternFill("solid", fgColor="116D6E")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style="thin", color="D7DCE3"),
        right=Side(style="thin", color="D7DCE3"),
        top=Side(style="thin", color="D7DCE3"),
        bottom=Side(style="thin", color="D7DCE3"),
    )

    summary = wb.create_sheet("전체요약")
    summary_headers = ["회사명", "파일명"]
    for label, periods in REPORT_FIELDS:
        summary_headers.extend(f"{label}({'당기' if period == 'current' else '전기'})" for period in periods)
    summary_headers.append("확인 필요 항목")
    summary.append(summary_headers)

    for doc in documents:
        values_by_item = {row["item"]: row for row in doc["rows"]}
        summary_row = [doc.get("companyName") or company_name_from_file(doc["fileName"]), doc["fileName"]]
        missing_or_check = []
        for label, periods in REPORT_FIELDS:
            row = values_by_item.get(label, {})
            summary_row.extend(row.get(period, "") for period in periods)
            if row.get("status") != "확인 필요":
                missing_or_check.append(label)
        summary_row.append(", ".join(missing_or_check))
        summary.append(summary_row)

    report_column_count = sum(len(periods) for _label, periods in REPORT_FIELDS)
    format_sheet(summary, header_fill, header_font, border, money_columns=range(3, 3 + report_column_count))
    summary.freeze_panes = "C2"
    summary.column_dimensions["A"].width = 18
    summary.column_dimensions["B"].width = 34
    for col_index in range(3, 3 + report_column_count):
        summary.column_dimensions[get_column_letter(col_index)].width = 18
    summary.column_dimensions[get_column_letter(3 + report_column_count)].width = 24

    for doc in documents:
        ws = wb.create_sheet(doc["sheetName"])
        ws.append(["항목", "당기 금액", "전기 금액", "확인 상태", "OCR 원문 줄"])
        for row in doc["rows"]:
            ws.append([row["item"], row["current"], row["previous"], row["status"], row.get("sourceLine", "")])

        format_sheet(ws, header_fill, header_font, border, money_columns=(2, 3))
        ws.freeze_panes = "A2"
        widths = [18, 18, 18, 14, 80]
        for col_index, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(col_index)].width = width

    wb.save(path)


def format_sheet(ws, header_fill, header_font, border, money_columns):
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    money_columns = set(money_columns)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            if cell.column in money_columns and cell.row > 1:
                cell.number_format = '#,##0.##;[Red]-#,##0.##'


def company_name_from_file(file_name: str) -> str:
    bracket_match = re.search(r"\[([^\]]+)\]", file_name)
    if bracket_match:
        return bracket_match.group(1).strip()
    stem = Path(file_name).stem
    stem = re.sub(r"^\s*#?\d+(?:[-.]\d+)*[.)]?\s*", "", stem)
    stem = re.sub(r"\(\s*주\s*\)|㈜|주식회사", "", stem)
    stem = re.sub(
        r"재무상태표|대차대조표|포괄손익계산서|손익계산서|손익게산서|재무제표|가결산|확정|별도|연결|감사보고서",
        " ",
        stem,
        flags=re.IGNORECASE,
    )
    stem = re.sub(r"(?<![A-Za-z0-9])(?:BS|PL)(?![A-Za-z0-9])", " ", stem, flags=re.IGNORECASE)
    stem = re.sub(r"20\d{6}|\d{2,4}\s*년|\d{1,2}\s*월|\d{1,2}\s*일|[1-4]\s*분기|[1-4]\s*q", " ", stem, flags=re.IGNORECASE)
    stem = re.sub(r"\d+(?:[._/-]\d+)+", " ", stem)
    stem = re.sub(r"\(\s*\)", " ", stem)
    stem = re.sub(r"[_\s-]+", " ", stem).strip()
    return stem or Path(file_name).stem


def install_hint() -> str:
    return (
        "현재 PC에서 ocrmypdf 실행 파일을 찾지 못했습니다. 설치되어 있다면 ocrmypdf.exe 전체 경로를 "
        "ocrmypdf_path.txt 파일에 한 줄로 저장한 뒤 서버를 다시 실행하세요. 예: "
        "C:\\Users\\사용자\\AppData\\Roaming\\Python\\Python312\\Scripts\\ocrmypdf.exe"
    )


def safe_filename(name: str) -> str:
    name = Path(name).name
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def unique_sheet_name(name: str, index: int) -> str:
    clean = re.sub(r"[\\/*?:\[\]]", " ", name).strip() or f"PDF{index}"
    return f"{index}_{clean}"[:31]


def to_url(path: Path) -> str:
    relative = path.resolve().relative_to(ROOT)
    return "/" + "/".join(relative.parts)


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8777
    OUTPUTS.mkdir(exist_ok=True)
    (OUTPUTS / "server_last_start.txt").write_text(
        f"started={time.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"url=http://127.0.0.1:{port}/server\n"
        f"root={ROOT}\n"
        f"outputs={OUTPUTS}\n",
        encoding="utf-8",
    )
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"Local OCR server: http://127.0.0.1:{port}/server")
    server.serve_forever()


if __name__ == "__main__":
    main()
