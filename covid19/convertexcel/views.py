from django.shortcuts import render, redirect
from django.http import HttpResponseRedirect

from pathlib import Path
import os
from django.conf import settings

from xlrd import open_workbook,cellname
from xlutils.copy import copy
import xlrd


from .forms import *
from datetime import datetime

import re
# Create your views here.


def index(request):
    if request.method == "POST":
        form = FileUploadForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            file = form.instance
            return redirect('view', id=file.id)
    else:
        form = FileUploadForm()
    context = {

        'form': form

    }

    return render(request, "index.html", context)



def view(request, id=0):
    file = FileUpload.objects.get(pk=id)


    book = open_workbook(filename = file.excel_file.url[1:])
    wb = copy(book)
    s = wb.get_sheet(0)
    sheet = book.sheet_by_index(0)

    for row_index in range(8, sheet.nrows):
        excel_name = sheet.cell(row_index,1).value
        excel_birthdate = str(sheet.cell(row_index,2).value)
        clean_birthdate = excel_birthdate.strip()
        clean_birthdate = clean_birthdate.replace(".", "/")
        split_birthdate = clean_birthdate.split("/")

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
            text_birthdate = str(python_date.strftime("%d")) +"/" + str(python_date.month) + "/" + str(python_date.year)

        clean_name = excel_name.title()
        s.write(row_index, 2, text_birthdate)
        s.write(row_index, 1, clean_name)

    new_file_name = 'media/static/output/' + str(id) + '.xls'
    wb.save(new_file_name)
    file.output = "/" + new_file_name
    file.save()

    context = {
        'file': file
    }

    return render(request, 'view.html', context)