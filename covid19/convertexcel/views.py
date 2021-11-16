from django.shortcuts import render, redirect
from django.http import HttpResponseRedirect

from pathlib import Path
import os
from django.conf import settings

from xlrd import open_workbook,cellname
from xlutils.copy import copy


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
        excel_date = sheet.cell(row_index,2).value
        string_date = str(excel_date)
        if len(string_date) == 10:

            clean_string = string_date.strip()
            clean_string = clean_string.replace(".", "/")
            x = clean_string.split("/")
            s.write(row_index, 2, clean_string)

            if int(x[1]) > 12:
                temp_day = x[1]
                x[1] = x[0]
                x[0] = temp_day

            for i in x:
            

                print(i)
        else:
            print("F")
    

    new_file_name = 'media/static/output/' + str(id) + '.xls'
    wb.save(new_file_name)
    file.output = "/" + new_file_name
    file.save()

    context = {
        'file': file
    }

    return render(request, 'view.html', context)