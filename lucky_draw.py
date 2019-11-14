import xlrd
import random
import time
import collections


def load_applies():
    filename = "응모리스트샘플.xlsx"
    wb = xlrd.open_workbook(filename)
    sheet = wb.sheet_by_index(0)

    result = []
    for i in range(1, sheet.nrows):
        (num, nickname, companion) = sheet.row_values(i)

        num = int(num)
        nickname = str(nickname)

        if len(nickname) > 0:
            result.append((num, nickname, companion))
    return result


applies = load_applies()
print(applies)

wins = []
sum = 0
max_win = 60 # 동행인 포함 총 인원
win_count_with_companion = 0
win_count_without_companion = 0

nicknames = []

while True:
    random.shuffle(applies)
    print(applies)
    win_candidate = applies.pop(0)
    nicknames.append(win_candidate[1])

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
with open("당첨결과.txt", "w") as file:
    for win in wins:
        file.write(f"{win[0]},{win[1]},{win[2]},\n")
print(f"당첨자수:{len(wins)}, (동행인 O:{win_count_with_companion}, 동행인 X:{win_count_without_companion}, 총인원:{win_count_with_companion*2 + win_count_without_companion})")

# 당첨자 중 중복된 닉네임 있는지 검사
nicknames_counter = collections.Counter(nicknames)
duplicates = {x : nicknames_counter[x] for x in nicknames_counter if nicknames_counter[x] >= 2}
for win in wins:
    nickname = win[1]
    if nickname in duplicates:
        print(f"{nickname}은 닉네임이 중복되니 주의! 당첨자는 {win} 입니다.")