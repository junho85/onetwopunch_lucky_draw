import xlrd
import random
import time


def load_applies():
    wb = xlrd.open_workbook("응모리스트샘플.xlsx")
    sheet = wb.sheet_by_index(0)

    result = []
    for i in range(1, sheet.nrows):
        (num, nickname) = sheet.row_values(i)
        if len(nickname) > 0:
            result.append({int(num): nickname})
    return result


applies = load_applies()
print(applies)

while True:
    random.shuffle(applies)
    print("탈락", applies.pop())
    print(applies)

    if len(applies) == 5:
        break

    time.sleep(0.1)

print("🥳🎁👏당첨을 축하합니다🎊🎉🎈")
print(applies)