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

def write_answers(ws, teacher_column, start,end, answers):
    status = 0
    ws[teacher_column + str(start-1)] = 'Đáp án GV'
    for i in range(start,end):
        ws[teacher_column + str(i)] = answers[status]
        status += 1

def grade_per_file(ws, answers, student_column, total_column, row_start, row_end):
    sum_row = row_start-1
    count = row_start

    user_answers = read_user_value_from_excel(ws,student_column, row_start, row_end)
    sum = 0
    for user, answer in zip(user_answers, answers):
        if user == answer:
            ws[total_column+ str(count)] = 1
            ws[total_column+ str(count)].alignment = Alignment(horizontal="center")
            sum += 1
        else:
            ws[total_column+ str(count)] = 0
            ws[total_column+ str(count)].alignment = Alignment(horizontal="center")
        count += 1
    ws[total_column + str(sum_row)] = sum 
    ws[total_column + str(sum_row)].alignment = Alignment(horizontal="center")

if __name__ == "__main__": 
    files = os.listdir('hocsinh')
    for file in files:
        wb = openpyxl.load_workbook("/home/hieu/GradingExcel/hocsinh/" + file)
        ws = wb.worksheets[0]

        lis_answer = read_answer_from_txt("lis.txt")
        read_answer = read_answer_from_txt("read.txt")
        x_read_answer = read_answer_from_txt("read2.txt")

        write_answers(ws, "C", 9, 15, lis_answer)
        write_answers(ws, "G", 9, 19, read_answer)
        write_answers(ws, "K", 9, 29, x_read_answer)

        grade_per_file(ws, lis_answer, student_column="B", total_column="D", row_start=9 , row_end=15) #rowend should +1 due to excel
        grade_per_file(ws, read_answer,  student_column="F", total_column="H", row_start=9 , row_end=19)
        grade_per_file(ws, x_read_answer,  student_column="J", total_column="L", row_start=9 , row_end=29)

        wb.save("/home/hieu/GradingExcel/result/" + file)
