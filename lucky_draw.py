import xlrd
import random
import time


def load_applies():
    wb = xlrd.open_workbook("응모리스트샘플.xlsx")
    sheet = wb.sheet_by_index(0)

    result = []
    for i in range(1, sheet.nrows):
        (num, nickname, companion) = sheet.row_values(i)
        if len(nickname) > 0:
            result.append((int(num), nickname, companion))
    return result


applies = load_applies()
print(applies)

wins = []
sum = 0
max_win = 60 # 동행인 포함 총 인원
win_count_with_companion = 0
win_count_without_companion = 0

while True:
    random.shuffle(applies)
    print(applies)
    win_candidate = applies.pop(0)

    if win_candidate[2].strip().upper() == "O":
        if max_win - sum >= 2:
            sum += 2
            win_count_with_companion += 1
        else:
            print(f"😥 남은 자리가 하나뿐이라 탈락 ㅠㅠ {win_candidate}")
            continue
    else:
        sum += 1
        win_count_without_companion += 1

    print(f"🎊 당첨 {win_candidate}")
    wins.append(win_candidate)

    if sum == max_win:
        break

    time.sleep(0.1)

print("🥳🎁👏당첨을 축하합니다🎊🎉🎈")
print(wins)
print(f"당첨자수:{len(wins)}, (동행인O:{win_count_with_companion}, 동행인X:{win_count_without_companion}, 총인원:{win_count_with_companion*2 + win_count_without_companion})")
