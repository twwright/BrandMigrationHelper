from openpyxl import Workbook, load_workbook
from openpyxl.writer.excel import save_virtual_workbook
from django.http import HttpResponse

# SERVICES

# CREATE NEW SERVICES WORKBOOK

def services(data_file):

    servbook = Workbook()
    servsheet = servbook.active

    servsheet["A1"] = "Name"
    servsheet["B1"] = "Category"
    servsheet["C1"] = "SubCategory"
    servsheet["D1"] = "SKU"
    servsheet["E1"] = "Barcode"
    servsheet["F1"] = "Duration"
    servsheet["G1"] = "RecoveryTime"
    servsheet["H1"] = "ServiceType"
    servsheet["I1"] = "AllowCustomersToBookOnline"
    servsheet["J1"] = "Description"
    servsheet["K1"] = "RequiresTwoStaffMembers"
    servsheet["L1"] = "Price"
    servsheet["M1"] = "Addon"
    servsheet["N1"] = "SpecialEvent"

    # OPEN AND CLEAN WORKBOOK TO COPY FROM

    copyservbook = load_workbook(data_file)
    copyservsheet = copyservbook.active
    copyservsheet.delete_rows(idx=1, amount=2)

    # COPY CELLS FROM PREVIOUS WORKBOOK
    allrows = copyservsheet.max_row

    for idx in range(1, allrows + 1):
        a = copyservsheet.cell(row=idx, column=1)
        servsheet.cell(row=idx + 1, column=1).value = a.value

        b = copyservsheet.cell(row=idx, column=2)
        servsheet.cell(row=idx + 1, column=2).value = b.value
        if b.value == "Add-ons":
           servsheet.cell(row=idx + 1, column=13).value = 1
        else:
           servsheet.cell(row=idx + 1, column=13).value = 0

        c = copyservsheet.cell(row=idx, column=3)
        servsheet.cell(row=idx + 1, column=3).value = c.value

        d = copyservsheet.cell(row=idx, column=4)
        servsheet.cell(row=idx + 1, column=10).value = d.value

        g = copyservsheet.cell(row=idx, column=11)
        if g.value == None:
           servsheet.cell(row=idx + 1, column=7).value = 0
        else:
          servsheet.cell(row=idx + 1, column=7).value = g.value

        h = copyservsheet.cell(row=idx, column=8)
        servsheet.cell(row=idx + 1, column=11).value = h.value

        i = copyservsheet.cell(row=idx, column=9)
        servsheet.cell(row=idx + 1, column=6).value = i.value

        k = copyservsheet.cell(row=idx, column=8)
        if k.value == 1:
            servsheet.cell(row=idx + 1, column=11).value = k.value
        else:
           servsheet.cell(row=idx + 1, column=11).value = 0

        l = copyservsheet.cell(row=idx, column=12)
        servsheet.cell(row=idx + 1, column=12).value = l.value

        servsheet.cell(row=idx + 1, column=14).value = 0

    servsheet.delete_rows(idx=allrows, amount=2)

    # SAVE WORKBOOK
    response = HttpResponse(content=save_virtual_workbook(servbook), content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename=services-output.xlsx'
    return response