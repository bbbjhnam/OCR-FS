# 고품질 OCR 설치 안내

`start_ocr_server.bat`은 PC에 `OCRmyPDF`가 설치되어 있으면 자동으로 고품질 OCR을 사용합니다.

## 필요한 구성

- OCRmyPDF
- Tesseract OCR
- Tesseract 한국어 언어 데이터(`kor`)

## 동작 방식

설치가 완료되면 프로그램은 다음 방식으로 PDF를 처리합니다.

```text
ocrmypdf --force-ocr --deskew --rotate-pages --pdf-renderer sandwich --sidecar OCR텍스트.txt -l kor+eng 입력.pdf 출력_OCR완료.pdf
```

- `--force-ocr`: 기존 텍스트 레이어가 있어도 다시 OCR
- `--deskew`: 기울어진 스캔 보정
- `--rotate-pages`: 돌아간 페이지 자동 보정
- `--pdf-renderer sandwich`: hOCR 파싱 오류를 피하기 위해 Tesseract의 텍스트 PDF 방식을 사용
- `--sidecar`: 엑셀 추출용 OCR 텍스트 파일을 별도로 저장
- `-l kor+eng`: 한국어와 영어 OCR

`unpaper`가 설치된 PC에서는 `--clean` 옵션도 자동으로 추가합니다. `unpaper`가 없으면 `--clean` 없이 실행합니다.

## 설치 확인

설치 후 새 명령 프롬프트에서 아래 명령이 동작해야 합니다.

```text
ocrmypdf --version
tesseract --version
tesseract --list-langs
```

`tesseract --list-langs` 결과에 `kor`와 `eng`가 보여야 합니다.

## 설치했는데 프로그램이 못 찾는 경우

`check_ocr_setup.bat`을 실행해 확인하세요.

만약 `NOT FOUND`가 나오면 `ocrmypdf_path.txt` 파일에 `ocrmypdf.exe` 전체 경로를 한 줄로 입력합니다.

예:

```text
C:\Users\사용자\AppData\Roaming\Python\Python312\Scripts\ocrmypdf.exe
```

또는 실행 파일이 없고 패키지만 설치된 경우에는 Python 명령을 넣어도 됩니다.

```text
C:\Users\사용자\AppData\Local\Python\pythoncore-3.14-64\python.exe -m ocrmypdf
```

저장 후 `start_ocr_server.bat`을 다시 실행하면 됩니다.

## 참고

현재 프로그램은 `ocrmypdf` 명령을 자동으로 찾아 사용합니다. 설치 후에는 `start_ocr_server.bat`을 다시 실행하면 됩니다.
