"""
Office Document Copy Utility
=============================
현재 실행 중인 Office 앱(Acrobat, Excel, PowerPoint, Word)에서
열려있는 문서를 리스팅하고, 선택한 문서를 원하는 위치에 복사 저장합니다.

사용법:
    python office_copy.py
"""

import os
import shutil
import sys
import tkinter as tk
from tkinter import filedialog

import pythoncom
import win32com.client as win32

from InquirerPy import inquirer

# ──────────────────────────────────────────────
# PDF Save Flags
# ──────────────────────────────────────────────
PDSaveFull = 0x0001
PDSaveCopy = 0x0002
PDSaveCollectGarbage = 0x0020


# ──────────────────────────────────────────────
# 문서 리스팅 함수들
# ──────────────────────────────────────────────

def list_pdf_docs():
    """Acrobat에서 열려있는 PDF 문서 목록을 반환합니다."""
    docs = []
    try:
        pythoncom.CoInitialize()
        acro_app = win32.Dispatch("AcroExch.App")
        num = acro_app.GetNumAVDocs()
        for i in range(num):
            avdoc = acro_app.GetAVDoc(i)
            if avdoc:
                pddoc = avdoc.GetPDDoc()
                title = avdoc.GetTitle() or f"PDF 문서 {i + 1}"
                docs.append({
                    "app": "pdf",
                    "name": title,
                    "avdoc": avdoc,
                    "pddoc": pddoc,
                })
    except Exception:
        pass
    finally:
        pythoncom.CoUninitialize()
    return docs


def list_excel_docs():
    """Excel에서 열려있는 워크북/시트 목록을 반환합니다."""
    docs = []
    try:
        excel = win32.GetActiveObject("Excel.Application")
        for i in range(1, excel.Workbooks.Count + 1):
            wb = excel.Workbooks(i)
            sheet = wb.ActiveSheet
            name = wb.Name
            if sheet:
                name = f"{wb.Name} ({sheet.Name})"
            docs.append({
                "app": "excel",
                "name": name,
                "workbook": wb,
                "sheet": sheet,
            })
    except Exception:
        pass
    return docs


def list_powerpoint_docs():
    """PowerPoint에서 열려있는 프레젠테이션 목록을 반환합니다."""
    docs = []
    try:
        ppt = win32.GetActiveObject("PowerPoint.Application")
        for i in range(1, ppt.Presentations.Count + 1):
            pres = ppt.Presentations(i)
            docs.append({
                "app": "powerpoint",
                "name": pres.Name,
                "presentation": pres,
            })
    except Exception:
        pass
    return docs


def list_word_docs():
    """Word에서 열려있는 문서 목록을 반환합니다."""
    docs = []
    try:
        word = win32.GetActiveObject("Word.Application")
        for i in range(1, word.Documents.Count + 1):
            doc = word.Documents(i)
            docs.append({
                "app": "word",
                "name": doc.Name,
                "document": doc,
            })
    except Exception:
        pass
    return docs


def list_all_open_docs():
    """모든 Office 앱에서 열려있는 문서를 통합 리스팅합니다."""
    docs = []
    docs.extend(list_pdf_docs())
    docs.extend(list_excel_docs())
    docs.extend(list_powerpoint_docs())
    docs.extend(list_word_docs())
    return docs


# ──────────────────────────────────────────────
# 복사 저장 함수들
# ──────────────────────────────────────────────

def copy_pdf(doc_info, save_path):
    """Acrobat PDF를 복사 저장합니다."""
    save_path = os.path.abspath(save_path)
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    if os.path.exists(save_path):
        os.remove(save_path)

    pythoncom.CoInitialize()
    try:
        pddoc = doc_info["pddoc"]
        flags = PDSaveFull | PDSaveCopy | PDSaveCollectGarbage
        if not pddoc.Save(flags, save_path):
            raise RuntimeError("PDF 복사 저장 실패")
    finally:
        pythoncom.CoUninitialize()

    return save_path


def copy_excel(doc_info, save_path):
    """Excel 활성 시트를 새 워크북으로 복사 저장합니다."""
    save_path = os.path.abspath(save_path)
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    sheet = doc_info["sheet"]
    excel = sheet.Application

    # 시트를 새 워크북으로 복사
    sheet.Copy()
    new_wb = excel.ActiveWorkbook

    try:
        new_wb.SaveAs(save_path, FileFormat=51)
    finally:
        new_wb.Close(SaveChanges=False)

    return save_path


def copy_powerpoint(doc_info, save_path):
    """PowerPoint 프레젠테이션을 복사 저장합니다."""
    save_path = os.path.abspath(save_path)
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    if os.path.exists(save_path):
        os.remove(save_path)

    doc_info["presentation"].SaveCopyAs(save_path)
    return save_path


def copy_word(doc_info, save_path):
    """Word 문서를 복사 저장합니다."""
    save_path = os.path.abspath(save_path)
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    if os.path.exists(save_path):
        os.remove(save_path)

    doc = doc_info["document"]

    # 1차: SaveCopyAs 시도
    try:
        doc.SaveCopyAs(save_path)
        return save_path
    except Exception:
        pass

    # 2차: 이미 저장된 문서라면 파일 복사
    if doc.Path:
        if not doc.Saved:
            doc.Save()
        shutil.copy2(doc.FullName, save_path)
        return save_path

    # 3차: 저장 불가
    raise RuntimeError(
        "현재 Word 문서는 아직 저장된 적이 없고, "
        "이 환경에서는 SaveCopyAs를 사용할 수 없습니다. "
        "먼저 원본 문서를 한 번 저장한 뒤 다시 실행해 주세요."
    )


# 앱별 복사 함수 매핑
_COPY_FUNCS = {
    "pdf": copy_pdf,
    "excel": copy_excel,
    "powerpoint": copy_powerpoint,
    "word": copy_word,
}

# 앱별 기본 확장자
_DEFAULT_EXT = {
    "pdf": ".pdf",
    "excel": ".xlsx",
    "powerpoint": ".pptx",
    "word": ".docx",
}

# 앱별 표시 라벨
_APP_LABELS = {
    "pdf": "PDF",
    "excel": "Excel",
    "powerpoint": "PowerPoint",
    "word": "Word",
}


# ──────────────────────────────────────────────
# 폴더 선택 다이얼로그
# ──────────────────────────────────────────────

def ask_save_folder():
    """tkinter 폴더 선택 다이얼로그를 열어 저장 폴더를 반환합니다."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="저장할 폴더를 선택하세요")
    root.destroy()
    return folder if folder else None


# ──────────────────────────────────────────────
# 메인 플로우
# ──────────────────────────────────────────────

def main():
    print("열려있는 Office 문서를 검색 중...")
    docs = list_all_open_docs()

    if not docs:
        print("현재 열려있는 Office 문서를 찾을 수 없습니다.")
        print("Acrobat, Excel, PowerPoint, Word 중 하나 이상을 실행하고 문서를 열어주세요.")
        sys.exit(1)

    # 문서 선택
    choices = []
    for doc in docs:
        label = _APP_LABELS.get(doc["app"], doc["app"])
        choices.append({"name": f"[{label}] {doc['name']}", "value": doc})

    selected = inquirer.select(
        message="복사할 문서를 선택하세요:",
        choices=choices,
    ).execute()

    if selected is None:
        print("취소되었습니다.")
        sys.exit(0)

    # 폴더 선택
    print("저장할 폴더를 선택해주세요... (다이얼로그 창 확인)")
    folder = ask_save_folder()
    if not folder:
        print("폴더 선택이 취소되었습니다.")
        sys.exit(0)

    # 파일명 입력 (기본값: 원본 파일명)
    default_name = selected["name"]
    ext = _DEFAULT_EXT.get(selected["app"], "")
    if ext and not default_name.lower().endswith(ext):
        default_name = os.path.splitext(default_name)[0] + ext

    filename = inquirer.text(
        message="저장할 파일명을 입력하세요:",
        default=default_name,
    ).execute()

    if not filename:
        print("파일명이 입력되지 않았습니다.")
        sys.exit(0)

    save_path = os.path.join(folder, filename)

    # 복사 저장 실행
    copy_func = _COPY_FUNCS[selected["app"]]
    try:
        result = copy_func(selected, save_path)
        print(f"저장 완료: {result}")
    except Exception as e:
        print(f"저장 실패: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
