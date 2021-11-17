from django.shortcuts import render, redirect
from django.http import HttpResponseRedirect
from pathlib import Path
import os
from django.conf import settings
from xlrd import open_workbook,cellname
from xlutils.copy import copy
from xlwt import Workbook
import xlrd

from .forms import *
from datetime import datetime, timedelta
import re
# Create your views here.

def index(request):
    files = FileUpload.objects.all().count()
    minutes = int(files) * 15
    hour = round(minutes / 60)

    if request.method == "POST":
        form = FileUploadForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            file = form.instance
            return redirect('view', id=file.id)
    else:
        form = FileUploadForm()
    context = {

        'form': form,
        'files': files,
        'd': minutes,
        'hour': hour

    }

    return render(request, "index.html", context)


def view(request, id=0):
    file = FileUpload.objects.get(pk=id)
    files = FileUpload.objects.all().count()

    minutes = int(files) * 15
    hour = round(minutes / 60)

    
    book = open_workbook(filename=file.excel_file.url[1:])
    book_template = open_workbook(filename="media/static/book_output_template.xls")
    wb = copy(book_template)
    s = wb.get_sheet(0)
    # wb = Workbook()
    # s = wb.add_sheet('PL1_Danh sach doi tuong tiem')
    sheet = book.sheet_by_index(0)
    count_p = len(range(8,sheet.nrows))

    for row_index in range(8, sheet.nrows):
        excel_name = sheet.cell(row_index,1).value
        excel_birthdate = str(sheet.cell(row_index,2).value)
        excel_sex = str(sheet.cell(row_index,3).value).lower()
        clean_birthdate = excel_birthdate.strip()
        clean_birthdate = clean_birthdate.replace(".", "/")
        split_birthdate = clean_birthdate.split("/")


        # chuyển đổi ngày tháng năm sinh
        if len(split_birthdate) == 3:
            #thêm kí tự cho đủ ngày tháng
            split_birthdate[0] = "0" + split_birthdate[0]
            split_birthdate[1] = "0" + split_birthdate[1]
            split_birthdate[2] = "1" + split_birthdate[2]
            split_birthdate[0] = split_birthdate[0][-2:]
            split_birthdate[1] = split_birthdate[1][-2:]
            split_birthdate[2] = split_birthdate[2][-4:]
            # đổi vị trí nếu chuỗi giữa lớn hơn 12

            if int(split_birthdate[1]) > 12:
                temp = split_birthdate[1]
                split_birthdate[1] = split_birthdate[0]
                split_birthdate[0] = temp
            text_birthdate = "/".join(split_birthdate)
        elif len(split_birthdate) == 2:
            python_date = datetime(*xlrd.xldate_as_tuple(int(split_birthdate[0]), 0))
            text_birthdate = python_date.strftime("%d/%m/%Y")
        else:
            text_birthdate = "01/01/1990"

        # chuyển đổi giới tính
        male_list = ['0', 'nam']
        female_list = ['1', 'nữ', 'nu']

        if excel_sex in male_list:
            text_sex = 0
        elif excel_sex in female_list:
            text_sex = 1
        else:
            if "thị" in str(excel_name).lower():
                text_sex = 1
            else:
                text_sex = 0


        # chuyển đổi họ và tên
        clean_name = excel_name.title()
        s.write(row_index, 2, text_birthdate)
        s.write(row_index, 1, clean_name)
        s.write(row_index, 3, text_sex)

    new_file_name = 'media/static/output/' + str(id) + '.xls'
    wb.save(new_file_name)
    file.output = "/" + new_file_name
    file.save()

    context = {
        'file': file,
        'files': files,
        'd': minutes,
        'hour': hour,
        'count_p': count_p
    }

    return render(request, 'view.html', context)