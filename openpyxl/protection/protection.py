""" openpyxlを使ったExcelのパスワード保護系のサンプル

公式ドキュメント
https://openpyxl.readthedocs.io/en/stable/protection.html

ここでは、

・ブックの保護
・シートの保護

を実装。

一方、

・読み取りパスワード
・書き込みパスワード

は、openpyxlでは実装できない。
"""
from itertools import chain
import pathlib

import openpyxl
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.styles import Protection

BASE_FILE_DIR = pathlib.Path(__file__).resolve().parents[0].joinpath('base_files')
RESULT_FILE_DIR = pathlib.Path(__file__).resolve().parents[0].joinpath('results')

# ディレクトリが無い場合は、親ディレクトリを含めて作成しておく
RESULT_FILE_DIR.mkdir(parents=True, exist_ok=True)

NO_PROTECTION_FILE = 'no_protection.xlsx'
BOOK_PROTECTION_FILE = 'book_protection.xlsx'
SHEET_PROTECTION_WITH_PASSWORD_FILE = 'sheet_protection_with_password.xlsx'
SHEET_PROTECTION_WITHOUT_PASSWORD_FILE = 'sheet_protection_without_password.xlsx'
ONE_SHEET_PROTECTION_WITHOUT_PASSWORD_FILE = 'one_sheet_protection_without_password.xlsx'

PASSWORD_FOR_BOOK = '54321'
PASSWORD_FOR_SHEET = '12345'


def protect_book_without_password():
    """ ブックをパスワード無しで保護する """
    wb = _load(NO_PROTECTION_FILE)

    # ブックを保護
    wb.security = WorkbookProtection()
    wb.security.lockStructure = True

    # 保存
    _save(wb, f'No_1_PROTECT_book_using_{NO_PROTECTION_FILE}')


def protect_book_with_password():
    """ ブックをパスワード付きで保護する """
    wb = _load(NO_PROTECTION_FILE)

    # ブックを保護
    wb.security = WorkbookProtection()
    wb.security.lockStructure = True
    wb.security.workbook_password = PASSWORD_FOR_BOOK

    # 保存
    _save(wb, f'No_2_PROTECT_book_using_{NO_PROTECTION_FILE}')


def protect_sheet_without_password():
    """ シートをパスワード無しで保護する """
    wb = _load(NO_PROTECTION_FILE)

    # 対象のワークシートオブジェクトを取得する
    ws = wb['Sheet1']

    # このシートのすべてのユーザーに許可する操作
    ws.protection.objects = True                # オブジェクトの編集
    ws.protection.scenarios = True              # シナリオの編集
    ws.protection.formatCells = True            # セルの書式設定
    ws.protection.formatColumns = True          # 列の書式設定
    ws.protection.formatRows = True             # 行の書式設定
    ws.protection.insertColumns = True          # 列の挿入
    ws.protection.insertRows = True             # 行の挿入
    ws.protection.insertHyperlinks = True       # ハイパーリンクの挿入
    ws.protection.deleteColumns = True          # 列の削除
    ws.protection.deleteRows = True             # 行の削除
    ws.protection.selectLockedCells = True      # ロックされたセルの選択
    ws.protection.selectUnlockedCells = True    # ロックされていないセルの選択
    ws.protection.sort = True                   # 並べ替え
    ws.protection.autoFilter = True             # フィルター
    ws.protection.pivotTables = True            # ピボットテーブルレポート

    # パスワード無しで保護
    ws.protection.enable()

    # 保存
    _save(wb, f'No_3_PROTECT_sheet_without_password_{NO_PROTECTION_FILE}')


def protect_sheet_with_unlock_cell():
    """ 一部のセルをロックせずに、シートをパスワード無しで保護する

    https://stackoverflow.com/questions/46877091/lock-some-cells-from-editing-in-python-openpyxl
    """
    wb = _load(NO_PROTECTION_FILE)
    ws = wb['Sheet1']

    # ロックを外したい(保護されない)セルを選ぶ
    unlock_cells = ws['A1:B3']

    # 取得したデータや型を見ると、行ごとにタプルでセルが入っている
    print(f'type: ({type(unlock_cells)}), values: {unlock_cells}')
    # => type: (<class 'tuple'>), values: ((<Cell 'Sheet1'.A1>, <Cell 'Sheet1'.B1>),
    #                                      (<Cell 'Sheet1'.A2>, <Cell 'Sheet1'.B2>),
    #                                      (<Cell 'Sheet1'.A3>, <Cell 'Sheet1'.B3>))

    # chain.from_iterable()でネストタプルを平坦にしてから処理 (使ってみたかっただけ)
    # 普通は for の2重ループで良いのかな
    for cell in chain.from_iterable(unlock_cells):

        # 念のための確認
        print(f'type: ({type(cell)}), values: {cell}')
        # => type: (<class 'openpyxl.cell.cell.Cell'>), values: <Cell 'Sheet1'.A1>

        # ロックを解除
        cell.protection = Protection(locked=False)

    # シートを保護
    ws.protection.enable()

    # 保存
    _save(wb, f'No_4_PROTECT_sheet_with_unlock_cell_using_{NO_PROTECTION_FILE}')


def protect_sheet_with_password():
    """ シートをパスワード付きで保護する """
    wb = _load(NO_PROTECTION_FILE)
    ws = wb['Sheet1']

    # パスワードをセット
    ws.protection.password = PASSWORD_FOR_SHEET

    # シートを保護
    ws.protection.enable()

    _save(wb, f'No_5_PROTECT_sheet_with_password_using_{NO_PROTECTION_FILE}')


def unprotect_book():
    """ ブックの保護を解除する """
    wb = _load(BOOK_PROTECTION_FILE)

    # 保護したときのパスワードをセット
    wb.security.workbook_password = PASSWORD_FOR_BOOK

    # ブックの保護を解除
    # wb.security.lock_structureでも良い：Aliasが設定されている
    wb.security.lockStructure = False

    _save(wb, f'No_6_UNPROTECT_{BOOK_PROTECTION_FILE}')


def unprotect_sheet_without_password():
    """ パスワードのないシートの保護を解除する """
    wb = _load(SHEET_PROTECTION_WITHOUT_PASSWORD_FILE)
    ws = wb['Sheet1']

    # シートの保護を解除
    ws.protection.disable()

    _save(wb, f'No_7_UNPROTECT_{SHEET_PROTECTION_WITHOUT_PASSWORD_FILE}')


def unprotect_sheet_with_password():
    """ パスワードのあるシートの保護を解除する """
    wb = _load(SHEET_PROTECTION_WITH_PASSWORD_FILE)
    ws = wb['Sheet1']

    # シートを保護したときのパスワードをセット
    ws.protection.password = PASSWORD_FOR_SHEET

    # シートの保護を解除
    ws.protection.disable()

    _save(wb, f'No_8_UNPROTECT_{SHEET_PROTECTION_WITH_PASSWORD_FILE}')


def _load(file_name):
    return openpyxl.load_workbook(BASE_FILE_DIR.joinpath(file_name))


def _save(workbook, file_name):
    workbook.save(RESULT_FILE_DIR.joinpath(file_name))


if __name__ == '__main__':
    # 保護系
    protect_book_with_password()

    protect_book_without_password()

    protect_sheet_with_unlock_cell()

    protect_sheet_without_password()

    protect_sheet_with_password()

    # 保護の解除系
    unprotect_book()

    unprotect_sheet_without_password()

    unprotect_sheet_with_password()
