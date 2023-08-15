import os
import subprocess
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

base_dir = Path(__file__).parent
excel_path = base_dir / "chat_history.xlsx"

HEADER_ROW_NUMBER = 2
ROW_HEIGHT = 18


def is_open_output_excel() -> bool:
    """
    excel_pathが開かれているかどうかを返す
    :return: excel_pathが開かれているかどうか
    """
    if os.name == "nt":
        # Windows
        try:
            with excel_path.open("r+b"):
                pass
            return False
        except IOError:
            return True
    elif os.name == "posix":
        # unix系
        if excel_path.exists():
            result = subprocess.run(["lsof", str(excel_path)], stdout=subprocess.PIPE)
            return bool(result.stdout)
        else:
            return False


def load_or_create_workbook() -> tuple[openpyxl.Workbook, bool]:
    """
    ワークブックを読み込み、そのオブジェクトを返すとともに新規作成したかどうかを返す
    :returns: ワークブックのオブジェクト, 新規作成したかどうか
    """

    # ワークブックの読み込み
    if excel_path.exists():
        wb = openpyxl.load_workbook(excel_path)
        return wb, False
    else:
        wb = openpyxl.Workbook()
        return wb, True


def create_worksheet(title: str, wb, is_new: bool):
    """
    ワークシートを作成する
    :param title: ワークシートのタイトル
    :param wb: 対象のワークブックのオブジェクト
    :param is_new: ブックを新規作成したかどうか
    :return: ワークシートのオブジェクト
    """

    title = trim_invalid_chars(title)
    if is_new:
        ws = wb.active
        ws.title = title
    else:
        ws = wb.create_sheet(title)  # 同じ名前がある場合、末尾に数字が付与される

    wb.move_sheet(ws, offset=-len(wb.worksheets) + 1)
    wb.active = ws

    return ws


def trim_invalid_chars(string: str) -> str:
    """
    エクセルのシート名で使えない文字列を削除
    :param string: 対象の文字列
    :return: 使えない文字列を削除した文字列
    """

    invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
    for char in invalid_chars:
        string = string.replace(char, '')
    return string


def header_formatting(ws):
    """
    ヘッダーの書式設定を行う
    :param ws: ワークシートオブジェクト
    """

    # ヘッダーのフォント設定とヘッダーオブジェクトを変数に格納
    ws["A1"].font = Font(name="Meiryo", size=11, bold=True)
    header_a, header_b = ws[f"A{HEADER_ROW_NUMBER}"], ws[f"B{HEADER_ROW_NUMBER}"]

    # ヘッダーの書き込み
    ws["A1"].value = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    header_a.value, header_b.value = "ロール", "発言内容"

    # ヘッダーのフォント変更
    font_style = Font(name="Meiryo", size=11, bold=True, color="FFFFFF")
    header_a.font, header_b.font = font_style, font_style

    # ヘッダーの色を緑色にする
    excel_green = PatternFill(fill_type='solid', fgColor='217346')
    header_a.fill, header_b.fill = excel_green, excel_green

    # 列の幅を調整
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 168


def write_chat_history(ws, chat_log: list[dict]):
    """
    チャットの履歴を書き込み、書式設定する
    :param chat_log: チャットのログ
    :param ws: ワークシートオブジェクト
    """

    font_style = Font(name="Meiryo", size=10)
    assistant_style = PatternFill(fill_type='solid', fgColor='d9d9d9')

    # チャット内容の書き込み
    for i, content in enumerate(chat_log, 3):
        cell_a, cell_b = ws[f"A{i}"], ws[f"B{i}"]

        # ロールと発言内容を書き込み
        cell_a.value, cell_b.value = content["role"], content["content"]

        # セル内改行に合わせて表示を調整
        cell_b.alignment = Alignment(wrapText=True)

        # 行の高さを調整
        adjusted_row_height = len(content["content"].split("\n")) * ROW_HEIGHT
        ws.row_dimensions[i].height = adjusted_row_height

        # 書式設定
        cell_a.font, cell_b.font = font_style, font_style
        if content["role"] == "assistant":
            cell_a.fill, cell_b.fill = assistant_style, assistant_style


def open_workbook():
    """エクセルを開く"""
    if os.name == "nt":
        # windows
        os.system(f"start {excel_path}")
    elif os.name == "posix":
        # mac
        os.system(f"open {excel_path}")


def output_excel(chat_log: list[dict], chat_summary: str):
    """
    chat_history.xlsx にチャットの履歴を書き込むためのエントリポイント。
    :param chat_log: チャットのログ
    :param chat_summary: チャットの要約
    """

    wb, is_new_create_wb = load_or_create_workbook()
    ws = create_worksheet(chat_summary, wb, is_new_create_wb)

    header_formatting(ws)
    write_chat_history(ws, chat_log)
    wb.save(excel_path)
    wb.close()
    open_workbook()


if __name__ == "__main__":
    test_log = [{"role": "user", "content": "こんにちは"}, {"role": "assistant", "content": "hello"}]
    test_summary = "test"
    output_excel(test_log, test_summary)
