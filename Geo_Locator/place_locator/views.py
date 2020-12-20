#!/usr/bin/python
# -*- coding: utf-8 -*-


import xlwt
import pandas as pd
from tablib import Dataset
from geopy.geocoders import Nominatim
from django.shortcuts import render
from django.http import HttpResponse
from django.contrib import messages
from django.shortcuts import redirect
from django.utils.datastructures import MultiValueDictKeyError



# Create your views here.

def index(request):
    if request.method == 'POST':
        if 'submit' in request.POST:
            dataset = Dataset()
            try:
                file_data = request.FILES['filename']  # selected file name
            except MultiValueDictKeyError:
                return render(request, 'place_locator/index.html', {})
            if not file_data.name.endswith('xlsx'):
                data = pd.read_csv(file_data)  # Read CSV file
                messages.info(request, 'wrong format')
                imported_data = dataset.load(data)
            else:
                imported_data = dataset.load(file_data.read(),
                        format='xlsx')  # Read excel file
            global fileobj

            def fileobj():
                return imported_data

        elif 'download' in request.POST:

            return redirect('download')
    return render(request, 'place_locator/index.html', {})  # Rendering to index file


def download(request):
    data = fileobj()
    longitude = []  # longitude collection
    latitude = []  # latitude collection
    for address in data:
        obj = Nominatim(user_agent='place')
        try:
            values = obj.geocode(address[1])
            longitude.append(values.longitude)
            latitude.append(values.latitude)
        except AttributeError:
            longitude.append('null/wrong address')
            latitude.append('null/wrong address')
    data.append_col(longitude, header='latitude')
    data.append_col(latitude, header='longitude')

    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="place_locator.xls"'  # Creating Excel file
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('co-ordinator')  # Creating spreed sheet
    font_style = xlwt.XFStyle()
    font_style.font.bold = True

    headers = data.headers
    row_num = 0

    # Adding header to file

    for col_num in range(len(headers)):
        ws.write(row_num, col_num, headers[col_num], font_style)  # at 0 row 0 column

    font_style = xlwt.XFStyle()
    col_num = 0

    # Adding row data to particular column

    for col_name in headers:
        for row_num in range(1, len(data['Sl.no']) + 1):
            ws.write(row_num, col_num, data[col_name][row_num - 1],
                     font_style)
        col_num += 1
    wb.save(response)
    return response
