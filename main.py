import openpyxl
from selenium import webdriver


# Excelファイル名
file_name = './URL.xlsx'

# シート名
sheet_name = 'URL'

# URLの開始行
row_start = 2

# URLの最終行
row_end = 6


if __name__ == '__main__':
    # Excelファイルをロードする
    workbook = openpyxl.load_workbook(file_name)
    worksheet = workbook[sheet_name]

    # Chromeブラウザを初期化する
    driver = webdriver.Chrome()

    for row in range(row_start, row_end + 1):
        # URLを読み込む
        url = worksheet[f'B{row}'].value

        # ページを開く
        driver.get(url)

        # タイトルを取得する
        title = driver.title

        # タイトルをExcelファイルに書き込む
        worksheet[f'C{row}'].value = title

    # Excelファイルを保存
    workbook.save(file_name)

    # Chromeブラウザを終了する
    driver.quit()
