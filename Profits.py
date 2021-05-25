from tkinter import *
from tkinter import Tk, Button, filedialog, Label, messagebox
import xlrd
import xlsxwriter


medicine = {}

def openfile(x: str):
    if x == 'b1':
        global book
        book = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xls *.xlsx'),("All Files", "* ")])
        if book is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return
        print(book)
        return book
    if x == 'b2':
        global book2
        book2 = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xls *.xlsx'),("All Files", "* ")])
        if book2 is None:  # asksaveasfile return `None` if dialog closed with "cancel".
            return
        print(book2)
        return book2

def saveFile():
    try:
        new_book = xlrd.open_workbook(book)
        new_book2 = xlrd.open_workbook(book2)
        read_file(new_book, new_book2)
        print(medicine)
        write_to_file(medicine)

    except:
        messagebox.showinfo(title='Error', message='Try Again, An Error Occurred')


def read_file(new_book, new_book2):
    # Count Stmnts for Required Data
    valid_count = 0
    app_req_count = 0
    intr_req_count = 0

    # Count Stmts for Missing Data
    missing_app_count = 0
    missing_intr_count = 0
    missing_oe_intr_count = 0

    sheet = new_book.sheets()[0]
    sheet2 = new_book2.sheets()[0]

    # Prices Row, Each row from item information will be appended here
    prices = []
    count = -1
    # Read the item information and append to list prices
    for row_num, row in enumerate(sheet2.get_rows()):
        if row_num > 5:
            count += 1
            price_rate = sheet2.row(row_num)[6]
            quantity = sheet2.cell(row_num, 2)
            prices.append(row[1:])
            if (float(quantity.value) != 0.0) and (float(price_rate.value) != 0.0):
                answer = float(price_rate.value) / float(quantity.value)
                answer = str(round(answer, 2))
                prices[count].append(answer)
            else:
                prices[count].append(0.0)

    # create medicine dictionary with quanity, total value as key
    for col_num, row in enumerate(sheet.get_rows()):
        if col_num > 5:
            if row[1].value not in medicine.keys():
                medicine.setdefault(row[1].value, []).extend([row[2].value, row[3].value])

            else:
                medicine.setdefault(row[1].value, []).extend([row[2].value, row[3].value])

    # merge prices with dictionary, where relevant
    for i in range(len(prices)):
        medicine_name = prices[i][0].value
        item_price = (prices[i][-1])
        if medicine_name in medicine.keys():
            medicine.setdefault(medicine_name, []).extend([float(prices[i][-1])])

    return medicine


def write_to_file(d:dict)->None:
    row = 0
    col = 0
    workbook = xlsxwriter.Workbook('Final.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    worksheet.write('A1', 'Medicine Names', bold)
    worksheet.write('B1', 'QTY SALES', bold)
    worksheet.write('C1', 'TOTAL AMT', bold)
    worksheet.write('D1', 'CP', bold)
    worksheet.write('E1', 'RP', bold)
    worksheet.write('F1', 'PROFIT DIFF', bold)
    worksheet.write('G1', 'TOTAL PROF', bold)
    worksheet.write('H1', 'PROF %', bold)

    for key in d.keys():
        col = 0
        row += 1
        worksheet.write(row, col, key)
        for item in d[key]:
            worksheet.write(row, col + 1, item)
            col += 1

        if row > 1:
            worksheet.write_formula('E{}'.format(row), '=IF(B{}=0, "", C{}/B{})'.format(row,row,row))
            worksheet.write_formula('F{}'.format(row), '=E{} - D{}'.format(row, row))
            worksheet.write_formula('G{}'.format(row), '=(C{} - (B{} * D{}))'.format(row, row, row))
            worksheet.write_formula('H{}'.format(row), '=IF(D{}=0, "", (((E{}/D{})-1) *100))'.format(row, row, row))

    worksheet.write("A{}".format((row + 2)), "Total", bold)

    worksheet.write_formula("C{}".format((row + 2)), "=SUM(C2:C{})".format(row + 1))
    worksheet.write_formula("E{}".format((row + 2)), "=SUM(E2:E{})".format(row + 1))
    worksheet.write_formula("F{}".format((row + 2)), "=SUM(F2:F{})".format(row + 1))
    worksheet.write_formula("G{}".format((row + 2)), "=SUM(G2:G{})".format(row + 1))

    workbook.close()
    print('Completed.')
    messagebox.showinfo(title='Success', message='File Successfully Saved!')


if __name__ == '__main__':

    root = Tk()
    root.title('Profits Excel Sheet')
    root.geometry('{}x{}'.format(480, 280))
    l = Label(root, text="Yourchemist")
    l.config(width=200, font=("Times New Roman", 20))
    l.place(relx=0.50, rely=0.2, anchor=CENTER)
    #button widget
    b1 = Button(root, text = "Retail Prices", command=lambda :openfile('b1'),  height=1, width=15)
    b1.place(x=100, y=120)
    b2 = Button(root, text="Purchase Prices",command=lambda :openfile('b2'), height=1, width=15)
    b2.place(x=280, y=120)
    b3 = Button(root, text="Profits", command=lambda:saveFile(), height=1, width=15)
    b3.place(relx=0.5, rely=0.75, anchor=CENTER)

    root.mainloop()
