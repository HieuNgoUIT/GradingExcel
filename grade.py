from openpyxl import Workbook
import openpyxl
import os
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment

def read_answer_from_txt(file):
    answers = []
    with open(file) as f:
        for line in f:
            value = line.strip()
            assert isinstance(value, str)
            assert len(value) == 1
            answers.append(value)
    return answers

def read_user_value_from_excel(ws, cell, start, end):
    users_answers = []
    for i in range(start,end):
        try:
            value = ws[cell + str(i)].value.strip()
        except:
            value = ""
            print("something happened at", cell, i)
        assert isinstance(value, str)
        assert len(value) == 1 or len(value) == 0
        users_answers.append(value)
    return users_answers

def write_answers(ws, cell, start,end, answers):
    status = 0
    ws[cell + str(start-1)] = 'Đáp án GV'
    for i in range(start,end):
        ws[cell + str(i)] = answers[status]
        status += 1

def grade_lis_per_file(ws, answers):
    global lis_start
    global lis_end
    sum_row = lis_start-1
    rs_start = lis_start

    user_answers = read_user_value_from_excel(ws,lis_column, lis_start, lis_end)
    sum = 0
    for user, answer in zip(user_answers, answers):
        if user == answer:
            ws[lis_ans_num+ str(rs_start)] = 1
            ws[lis_ans_num+ str(rs_start)].alignment = Alignment(horizontal="center", vertical="center")
            sum += 1
        else:
            ws[lis_ans_num+ str(rs_start)] = 0
            ws[lis_ans_num+ str(rs_start)].alignment = Alignment(horizontal="center", vertical="center")
        rs_start += 1
    ws[lis_ans_num + str(sum_row)] = sum
    ws[lis_ans_num + str(sum_row)].alignment = Alignment(horizontal="center", vertical="center")

def grade_read_per_file(ws, answers):
    global read_start
    global read_end
    sum_row = read_start-1
    rs_start = read_start

    user_answers = read_user_value_from_excel(ws,read_column, read_start, read_end)
    sum = 0
    for user, answer in zip(user_answers, answers):
        if user == answer:
            ws[read_ans_num+ str(rs_start)] = 1
            ws[read_ans_num+ str(rs_start)].alignment = Alignment(horizontal="center", vertical="center")
            sum += 1
        else:
            ws[read_ans_num+ str(rs_start)] = 0
            ws[read_ans_num+ str(rs_start)].alignment = Alignment(horizontal="center", vertical="center")
        rs_start += 1
    ws[read_ans_num + str(sum_row)] = sum 
    ws[read_ans_num + str(sum_row)].alignment = Alignment(horizontal="center", vertical="center")

if __name__ == "__main__":
    lis_column = 'B'
    lis_ans = 'C'
    lis_ans_num = 'D'
    lis_start = 9
    lis_end = 15 #should + 1 compare to excel

    read_column = 'F'   
    read_ans = 'G'
    read_ans_num = 'H'
    read_start = 9
    read_end = 19 #should + 1 compare to excel
    files = os.listdir('hocsinh')

    for file in files:
        wb = openpyxl.load_workbook("/home/hieu/Desktop/grade/hocsinh/" + file)
        ws = wb.worksheets[0]

        lis_answer = read_answer_from_txt("lis.txt")
        read_answer = read_answer_from_txt("read.txt")

        write_answers(ws, lis_ans, lis_start, lis_end, lis_answer)
        write_answers(ws, read_ans, read_start, read_end, read_answer)

        grade_lis_per_file(ws, lis_answer)
        grade_read_per_file(ws, read_answer)

        wb.save("/home/hieu/Desktop/grade/result/" + file)
