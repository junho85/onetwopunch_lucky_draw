import xlrd
import random
import time


def load_applies_for_book():
    filename = "응모리스트샘플_책.xlsx"
    # filename = "2019 원투펀치 이벤트 도서이벤트.xlsx"
    wb = xlrd.open_workbook(filename)
    sheet = wb.sheet_by_index(0)

    result = []
    for i in range(1, sheet.nrows):
        nickname = sheet.row_values(i)[0]

        nickname = str(nickname)

        if len(nickname) > 0:
            result.append(nickname)
    return result


applies = load_applies_for_book()
print(applies)

wins = [] # 당첨자 목록

while True:
    random.shuffle(applies)
    print(applies)
    win_candidate = applies.pop(0)

    if len(wins) < 5:
        # 5명 까지 당첨
        print(f"🎊 당첨 {win_candidate}")
        wins.append(win_candidate)
    else:
        break

    time.sleep(0.1)

print()
print("🥳🎁👏당첨을 축하합니다🎊🎉🎈")
print(wins)

# 당첨자 목록 저장
with open("당첨결과_당첨자_도서.txt", "w") as file:
    for win in wins:
        file.write(f"{win}\n")
print(f"당첨자수:{len(wins)}")