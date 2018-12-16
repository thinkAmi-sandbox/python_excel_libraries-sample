import msoffcrypto
import pathlib


BASE_DIR = pathlib.Path(__file__).resolve().parents[0]
PASSWORD = '12345'


def unlock():
    for file in BASE_DIR.iterdir():
        # Excelファイルだけ対象
        if not file.is_file() or file.suffix not in ['.xlsx', '.xls']:
            continue

        with file.open(mode='rb') as locked:
            # xlsxファイルの場合、読み取りパスワード無しのファイルは例外が発生する
            # is_encrypted()には以下の記載がある
            #
            # https://github.com/nolze/msoffcrypto-tool/blob/v4.6.3/msoffcrypto/format/ooxml.py#L143
            # def is_encrypted(self):
            #     # olefile cannot process non password protected ooxml files.
            #     # Hence if it has reached here it must be password protected.
            #     return True
            try:
                office_file = msoffcrypto.OfficeFile(locked)
            except OSError:
                if file.suffix == '.xlsx':
                    continue
                raise

            # 読み取りパスワードが設定されているかをチェック(xlsxはチェックできないので、xls向け)
            if not office_file.is_encrypted():
                continue

            # パスワードをセット
            # xlsでパスワードが設定されていない場合、load_key()時にエラーが出るため、事前にチェックが必要
            #   File "python3.6/site-packages/msoffcrypto/format/xls97.py", line 479, in load_key
            #     # Skip to FilePass; TODO: Raise exception if not encrypted
            #     num, size = workbook.skip_to(recordNameNum['FilePass'])
            #   File "python3.6/site-packages/msoffcrypto/format/xls97.py", line 428, in skip_to
            #     raise Exception("Record not found")
            # Exception: Record not found
            office_file.load_key(password=PASSWORD)

            # 読み取りパスワード解除後のファイルは、拡張子の前に '_unlocked' を付けて保存する
            unlocked_file = BASE_DIR.joinpath(f'{file.stem}_unlocked{file.suffix}')
            with unlocked_file.open(mode='wb') as unlocked:
                # パスワードを解除
                office_file.decrypt(unlocked)


if __name__ == '__main__':
    unlock()
