import xlrd
import random
import time
import collections


def load_applies():
    filename = "응모리스트샘플_공개방송.xlsx"
    # filename = "2019 원투펀치 공개방송 신청.xlsx"
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

wins = [] # 당첨자 목록
wins2 = [] # 당첨후보 목록

sum = 0 # 총 참여 인원
max_win = 60 # 동행인 포함 총 인원

win_count_with_companion = 0 # 동행인 포함 당첨자수
win_count_without_companion = 0 # 동행인 미포함 당첨자수

nicknames = [] # 중복 닉네임 검사용 리스트

while True:
    random.shuffle(applies)
    print(applies)
    win_candidate = applies.pop(0)
    nicknames.append(win_candidate[1])

    # 동행인 있으면 인원수 +2, 없으면 +1
    if win_candidate[2].strip().upper() == "O":
        sum += 2
    else:
        sum += 1

    if sum <= max_win:
        # 60명 까지 당첨
        print(f"🎊 당첨 {win_candidate}")
        wins.append(win_candidate)

        # 동행인 포함 당첨수와 미포함 당첨수
        if win_candidate[2].strip().upper() == "O":
            win_count_with_companion += 1
        else:
            win_count_without_companion += 1
    else:
        # 60명 넘으면 후보. 후보는 인원수 100명 까지
        print(f"후보 {win_candidate}")
        wins2.append(win_candidate)

    # 100명 넘으면 끝
    if sum >= 100:
        break

    time.sleep(0.1)

print()
print("🥳🎁👏당첨을 축하합니다🎊🎉🎈")
print(wins)

# 당첨자 목록 저장
with open("당첨결과_공개방송_당첨자.txt", "w") as file:
    for win in wins:
        file.write(f"{win[0]},{win[1]},{win[2]},\n")
print(f"당첨자수:{len(wins)}, (동행인 O:{win_count_with_companion}, 동행인 X:{win_count_without_companion}, 총인원:{win_count_with_companion*2 + win_count_without_companion})")

# 당첨자후보 목록 저장
with open("당첨결과_공개방송_당첨자후보.txt", "w") as file:
    for win in wins2:
        file.write(f"{win[0]},{win[1]},{win[2]},\n")

# 당첨자 중 중복된 닉네임 있는지 검사
nicknames_counter = collections.Counter(nicknames)
duplicates = {x : nicknames_counter[x] for x in nicknames_counter if nicknames_counter[x] >= 2}
for win in wins:
    nickname = win[1]
    if nickname in duplicates:
        print(f"{nickname}은 닉네임이 중복되니 주의! 당첨자는 {win} 입니다.")