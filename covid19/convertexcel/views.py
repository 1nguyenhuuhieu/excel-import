from django.shortcuts import render, redirect
from django.http import HttpResponseRedirect
from pathlib import Path
import os
from django.conf import settings
from xlrd import open_workbook,cellname
from xlutils.copy import copy
from xlwt import Workbook
import xlrd
from django.template.defaultfilters import slugify
import random


from .forms import *
from datetime import datetime, timedelta
import re
# Create your views here.

MAXA = {
    "thi": 17329,
    "tho": 17332,
    "thanh": 17335,
    "binh":	17338,
    "tam": 17341,
    "dinh":	17344,
    "hung":	17347,
    "cam": 17350,
    "duc": 17353,
    "tuong": 17356,
    "hoa": 17357,
    "tao": 17359,
    "vinh":	17362,
    "lang":	17365,
    "hoi":  17368,
    "thach": 17371,
    "phuc": 17374,
    "long": 17377,
    "khai": 17380,
    "linh": 17383,
    "cao": 17386
}


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

    temp_group = 8
    temp_tinh = 0
    temp_huyen = 0
    temp_xa = 0

    for row_index in range(8, sheet.nrows):
        excel_name = sheet.cell(row_index,1).value
        excel_birthdate = str(sheet.cell(row_index,2).value)
        excel_sex = str(sheet.cell(row_index,3).value).lower()
        excel_group = sheet.cell(row_index,4).value
        try:
            excel_phone = int(sheet.cell(row_index, 6).value)
        except:
            excel_phone = str(sheet.cell(row_index, 6).value)
        try:
            excel_ccnd = int(sheet.cell(row_index, 7).value)
            excel_ccnd = str(excel_ccnd)
        except:
            excel_ccnd = str(sheet.cell(row_index, 7).value)

        excel_MBH = str(sheet.cell(row_index, 8).value)
        excel_tinh = sheet.cell(row_index, 10).value
        excel_huyen = sheet.cell(row_index, 12).value
        excel_xa = sheet.cell(row_index, 14).value
        
        excel_donvicongtac = sheet.cell(row_index, 5).value
        excel_chitiet = sheet.cell(row_index, 15).value

    # chuy???n ?????i ng??y th??ng n??m sinh
        clean_birthdate = excel_birthdate.strip()
        clean_birthdate = clean_birthdate.replace(".", "/")
        split_birthdate = clean_birthdate.split("/")
        if len(split_birthdate) == 3:
            #th??m k?? t??? cho ????? ng??y th??ng
            split_birthdate[0] = "0" + split_birthdate[0]
            split_birthdate[1] = "0" + split_birthdate[1]
            split_birthdate[2] = "1" + split_birthdate[2]
            split_birthdate[0] = split_birthdate[0][-2:]
            split_birthdate[1] = split_birthdate[1][-2:]
            split_birthdate[2] = split_birthdate[2][-4:]
            # ?????i v??? tr?? n???u chu???i gi???a l???n h??n 12

            if int(split_birthdate[1]) > 12:
                temp = split_birthdate[1]
                split_birthdate[1] = split_birthdate[0]
                split_birthdate[0] = temp
            text_birthdate = "/".join(split_birthdate)
        elif len(split_birthdate) == 2:
            python_date = datetime(*xlrd.xldate_as_tuple(int(split_birthdate[0]), 0))
            text_birthdate = python_date.strftime("%d/%m/%Y")
        else:
            if split_birthdate:
                try:
                    year = int(split_birthdate[0])
                    if year in range(1920,2010):
                        text_birthdate = "01/01/" + str(year)
                except:
                    text_birthdate = "01/01/1990"
            else:
                text_birthdate = "01/01/1990"

        # chuy???n ?????i gi???i t??nh
        male_list = ['0', 'nam']
        female_list = ['1', 'n???', 'nu']

        if excel_sex in male_list:
            text_sex = 0
        elif excel_sex in female_list:
            text_sex = 1
        else:
            if "th???" in str(excel_name).lower():
                text_sex = 1
            else:
                text_sex = 0


        # chuy???n ?????i h??? v?? t??n
        clean_name = excel_name.title()

        # chuy???n ?????i m?? nh??m
        if excel_group:
            text_group = int(excel_group)
            temp_group = text_group
        else:
            text_group = int(temp_group)

        # chuy???n ?????i s??? ??i???n tho???i
        if excel_phone:
            if type(excel_phone) == "str":
                clean_phone = excel_phone.strip()
                clean_phone = clean_phone.replace(".", "")
                clean_phone = clean_phone.replace(",", "")
                clean_phone = clean_phone.replace(" ", "")
                clean_phone = clean_phone.replace("o", "0")
                clean_phone = clean_phone.replace("O", "0")
            else:
                clean_phone = str(excel_phone)
            string_phone = "0000" + clean_phone
            text_phone = string_phone[-10:]
        else:
            text_phone = "0999999999"

        # chuy???n ?????i s??? cccd
        if excel_ccnd:
            if len(str(excel_ccnd)) == 12 :
                text_ccnd = int(excel_ccnd)
            elif len(str(excel_ccnd)) == 9:
                text_ccnd = int(excel_ccnd)
            elif len(str(excel_ccnd)) == 11:
                text_ccnd = int(excel_ccnd)
                text_ccnd = "0" + str(text_ccnd)
            else:
                text_ccnd = "040" + str(text_sex) + text_birthdate[-2:]  + "123456"
                text_ccnd = text_ccnd[-12:]
        else:
            text_ccnd = "040" + str(text_sex) + text_birthdate[-2:]  + str(random.randint(100000, 999999))
            text_ccnd = text_ccnd[-12:]

        # chuy???n ?????i m?? th??? b???o hi???m
        if excel_MBH:
            text_MBH = "00000000000" + excel_MBH
            text_MBH = text_MBH[-15:]
        else:
            text_MBH = ""

        # t???nh huy???n x??


        if excel_tinh:
            text_tinh = int(excel_tinh)
            temp_tinh = text_tinh
        else:
            text_tinh = temp_tinh

        if excel_huyen:
            text_huyen = int(excel_huyen)
            temp_huyen = text_huyen
        else:
            text_huyen = temp_huyen
            
        if excel_xa:
            text_xa = str(excel_xa)
            text_xa = text_xa.replace("??", "d")
            text_xa = slugify(text_xa)
            text_xa = text_xa.replace("-", "")
            text_xa = text_xa.replace("xa", "")
            text_xa = text_xa.replace("thitrananhson", "thi")
            text_xa = text_xa.replace("thitran", "thi")
            text_xa = text_xa.replace("son", "")
            text_xa = text_xa.replace("tran", "")
            try:
                text_xa = MAXA[text_xa]
            except:
                if len(text_xa) == 5:
                    text_xa = str(text_xa)
                else:
                    text_xa = MAXA["thi"]
        else:
            text_xa = "17329"


        s.write(row_index, 2, text_birthdate)
        s.write(row_index, 1, clean_name)
        s.write(row_index, 3, text_sex)
        s.write(row_index, 4, text_group)
        s.write(row_index, 6, text_phone)
        s.write(row_index, 7, text_ccnd)
        s.write(row_index, 8, text_MBH)
        s.write(row_index, 10, text_tinh)
        s.write(row_index, 12, text_huyen)
        s.write(row_index, 14, text_xa)
        s.write(row_index, 5, excel_donvicongtac)
        s.write(row_index, 15, excel_chitiet)

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