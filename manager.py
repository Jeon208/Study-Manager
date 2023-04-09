import openpyxl
# pip install openpyxl      (in cmd)

kor = input("[I] Your grades in Korean language\n>> ")
eng = input("[I] Your grades in English\n>> ")
mat = input("[I] Your grades in Math\n>> ")
soc = input("[I] Your grades in Social studies\n>> ")
sci = input("[I] Your grades in Science\n>> ")
# 과목 점수 입력받기.

if "." in kor:
    kor = float(kor)
else:
    kor = int(kor)

if "." in eng:
    eng = float(eng)
else:
    eng = int(eng)

if "." in mat:
    mat = float(mat)
else:
    mat = int(mat)

if "." in soc:
    soc = float(soc)
else:
    soc = int(soc)

if "." in sci:
    sci = float(sci)
else:
    sci = int(sci)
# 과목의 점수에 소숫점이 포함되는지 확인하기.

mean = (kor + eng + mat + soc + sci)/5
# mean 정의하기.


dict_sub = {
    "국어":kor, 
    "영어":eng,
    "수학":mat,
    "사회":soc,
    "과학":sci
}

sorted_dict = dict(sorted(dict_sub.items(), key = lambda x : x[1], reverse=True))
# 딕셔너리를 값에 따라 정리

list_sub = list(sorted_dict) # 과목 리스트화하기.
print(list_sub) # 확인하기(생략가능).


write_wb = openpyxl.Workbook()
# write_wb 정의하기.

# write_ws = write_wb.create_sheet('생성시트')
write_ws = write_wb.active
# write_ws 정의

for i in range(1,6):
    write_ws[f'A{i}'] = list_sub[i-1]

for i in range(1,6):
    write_ws[f'B{i}'] = dict_sub[list_sub[i-1]]
    if dict_sub[list_sub[i-1]] >= 100:
        write_ws[f'C{i}'] = "완벽합니다!!"
    elif dict_sub[list_sub[i-1]] > 90:
        write_ws[f'C{i}'] = "충분합니다!!"
    elif dict_sub[list_sub[i-1]] > 80:
        write_ws[f'C{i}'] = "조금만 더 분발합시다!!"
    elif dict_sub[list_sub[i-1]] > 70:
        write_ws[f'C{i}'] = "노력합시다!!"
    elif dict_sub[list_sub[i-1]] > 60:
        write_ws[f'C{i}'] = "포기하지 맙시다!!"
    elif dict_sub[list_sub[i-1]] > 50:
        write_ws[f'C{i}'] = "남들이 쉴 때 공부합시다!!"
    else:
        write_ws[f'C{i}'] = "공부가 인생의 전부는 아닙니다!!"
# 점수에 따른 메시지 저장.

write_ws['A7'] = '평균'
write_ws['B7'] = mean
# 평균 넣기.

if mean >= 100:
    write_ws['C7'] = "완벽합니다!!"
elif mean > 90:
    write_ws['C7'] = "충분합니다!!"
elif mean > 80:
    write_ws['C7'] = "조금만 더 분발합시다!!"
elif mean > 70:
    write_ws['C7'] = "노력합시다!!"
elif mean > 60:
    write_ws['C7'] = "포기하지 맙시다!!"
elif mean > 50:
    write_ws['C7'] = "남들이 쉴 때 공부합시다!!"
else:
    write_ws['C7'] = "공부가 인생의 전부는 아닙니다!!"
# 점수에 따른 메시지 저장.

write_wb.save("managed.xlsx")
# 저장하기