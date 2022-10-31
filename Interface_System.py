import tkinter as tk
import tkinter.ttk
from tkinter.ttk import Treeview
from openpyxl import load_workbook

window = tk.Tk()
window.title("명부")
window.geometry('1200x600')
window.resizable(False, False)
wb = load_workbook("Purchase_Sale.xlsx")

def matrix_from_sheet(sheet):
    matrix = []
    for row in wb[sheet].rows:
        row_value = []
        for cell in row:
            row_value.append(cell.value)
        matrix.append(row_value)
    return matrix

# Page 1
def event1_1():
    frame1.pack_forget()
    btn1_1.pack_forget()
    btn1_2.pack_forget()

    frame2_1.pack(expand=True, fill='both')
    treeview_buy.pack()
    btn2_1_1.pack(side=tk.LEFT, anchor='sw')
    btn2_1_2.pack(side=tk.RIGHT, anchor='se')
    btn2_1_3.pack(side=tk.RIGHT, anchor='se')

def event1_2():
    frame1.pack_forget()
    btn1_1.pack_forget()
    btn1_2.pack_forget()

    frame2_2.pack(expand=True, fill='both')
    treeview_sell.pack()
    btn2_2_1.pack(side=tk.LEFT, anchor='sw')
    btn2_2_2.pack(side=tk.RIGHT, anchor='se')
    btn2_2_3.pack(side=tk.RIGHT, anchor='se')

frame1 = tk.Frame(window)
btn1_1 = tk.Button(frame1, text="방 보러 온 사람", command=event1_1)
btn1_2 = tk.Button(frame1, text="방 내놓으러 온 사람", command=event1_2)
frame1.pack(expand=True, fill='both')
btn1_1.pack(expand=True, fill='both', side=tk.LEFT)
btn1_2.pack(expand=True, fill='both', side=tk.RIGHT)


# Page 2_1
def event2_1_1():
    frame2_1.pack_forget()
    btn2_1_1.pack_forget()
    btn2_1_2.pack_forget()
    btn2_1_3.pack_forget()

    frame1.pack(expand=True, fill='both')
    btn1_1.pack(expand=True, fill='both', side=tk.LEFT)
    btn1_2.pack(expand=True, fill='both', side=tk.RIGHT)

def event2_1_2():
    frame2_1.pack_forget()
    btn2_1_1.pack_forget()
    btn2_1_2.pack_forget()
    btn2_1_3.pack_forget()

    frame3_1.pack()
    lb1_1.pack()
    entry1_1.pack()
    lb1_2.pack()
    entry1_2.pack()
    lb1_3.pack()
    entry1_3.pack()
    lb1_4.pack()
    entry1_4.pack()
    lb1_5.pack()
    entry1_5.pack()
    lb1_6.pack()
    entry1_6.pack()
    lb1_7.pack()
    entry1_7.pack()
    lb1_8.pack()
    entry1_8.pack()
    lb1_9.pack()
    entry1_9.pack()
    lb1_10.pack()
    entry1_10.pack()
    btn3_1_1.pack()
    btn3_1_2.pack()

def event2_1_3():
    selected_index = treeview_buy.selection()
    wb["Purchase_Quest"].delete_rows(int(selected_index[0]))
    treeview_buy.delete(*treeview_buy.get_children())
    list_changed = matrix_from_sheet("Purchase_Quest")
    for p in range(len(list_changed)):
        treeview_buy.insert("", 'end', text=str(p), values=list_changed[p], iid=str(p))

    wb.save("Purchase_sale.xlsx")


frame2_1 = tk.Frame(window)

treeview_buy = tk.ttk.Treeview(frame2_1, columns=["#1", "#2", "#3", "#4", "#5", "#6", "#7", "#8", "#9", "#10"], displaycolumns=["#1", "#2", "#3", "#4", "#5", "#6", "#7", "#8", "#9", "#10"], height=20, selectmode="browse")

treeview_buy.column("#0", width=90, anchor="center")
treeview_buy.heading("#0", text="index", anchor="center")

treeview_buy.column("#1", width=90, anchor="center")
treeview_buy.heading("#1", text="희망 지역", anchor="center")

treeview_buy.column("#2", width=90, anchor="center")
treeview_buy.heading("#2", text="전세/월세", anchor="center")

treeview_buy.column("#3", width=90, anchor="center")
treeview_buy.heading("#3", text="희망 가격", anchor="center")

treeview_buy.column("#4", width=70, anchor="center")
treeview_buy.heading("#4", text="희망 층수", anchor="center")

treeview_buy.column("#5", width=70, anchor="center")
treeview_buy.heading("#5", text="희망 칸수", anchor="center")

treeview_buy.column("#6", width=70, anchor="center")
treeview_buy.heading("#6", text="희망 평수", anchor="center")

treeview_buy.column("#7", width=130, anchor="center")
treeview_buy.heading("#7", text="보유 옵션", anchor="center")

treeview_buy.column("#8", width=130, anchor="center")
treeview_buy.heading("#8", text="이사 가능 날짜", anchor="center")

treeview_buy.column("#9", width=90, anchor="center")
treeview_buy.heading("#9", text="특이 사항", anchor="center")

treeview_buy.column("#10", width=150, anchor="center")
treeview_buy.heading("#10", text="연락처", anchor="center")

treelist_buy = matrix_from_sheet("Purchase_Quest")
for i in range(len(treelist_buy)):
    treeview_buy.insert("", 'end', text=str(i), values=treelist_buy[i], iid=str(i))

btn2_1_1 = tk.Button(frame2_1, text="뒤로", command=event2_1_1)
btn2_1_2 = tk.Button(frame2_1, text="등록", command=event2_1_2)
btn2_1_3 = tk.Button(frame2_1, text="삭제", command=event2_1_3)


# Page2_2
def event2_2_1():
    frame2_2.pack_forget()
    btn2_2_1.pack_forget()
    btn2_2_2.pack_forget()
    btn2_2_3.pack_forget()

    frame1.pack(expand=True, fill='both')
    btn1_1.pack(expand=True, fill='both', side=tk.LEFT)
    btn1_2.pack(expand=True, fill='both', side=tk.RIGHT)

def event2_2_2():
    frame2_2.pack_forget()
    btn2_2_1.pack_forget()
    btn2_2_2.pack_forget()
    btn2_2_3.pack_forget()

    frame3_2.pack()
    lb2_1.pack()
    entry2_1.pack()
    lb2_2.pack()
    entry2_2.pack()
    lb2_3.pack()
    entry2_3.pack()
    lb2_4.pack()
    entry2_4.pack()
    lb2_5.pack()
    entry2_5.pack()
    lb2_6.pack()
    entry2_6.pack()
    lb2_7.pack()
    entry2_7.pack()
    lb2_8.pack()
    entry2_8.pack()
    lb2_9.pack()
    entry2_9.pack()
    lb2_10.pack()
    entry2_10.pack()
    btn3_2_1.pack()
    btn3_2_2.pack()

def event2_2_3():
    selected_index = treeview_sell.selection()
    wb["Sale_List"].delete_rows(int(selected_index[0]))
    treeview_sell.delete(*treeview_sell.get_children())
    list_changed = matrix_from_sheet("Sale_List")
    for p in range(len(list_changed)):
        treeview_sell.insert("", 'end', text=str(p), values=list_changed[p], iid=str(p))

    wb.save("Purchase_sale.xlsx")


frame2_2 = tk.Frame(window)

treeview_sell = tk.ttk.Treeview(frame2_2, columns=["#1", "#2", "#3", "#4", "#5", "#6", "#7", "#8", "#9", "#10"], displaycolumns=["#1", "#2", "#3", "#4", "#5", "#6", "#7", "#8", "#9", "#10"], height=20, selectmode='browse')

treeview_sell.column("#0", width=90, anchor="center")
treeview_sell.heading("#0", text="index", anchor="center")

treeview_sell.column("#1", width=90, anchor="center")
treeview_sell.heading("#1", text="매물 주소", anchor="center")

treeview_sell.column("#2", width=90, anchor="center")
treeview_sell.heading("#2", text="전세/월세", anchor="center")

treeview_sell.column("#3", width=90, anchor="center")
treeview_sell.heading("#3", text="매물 가격", anchor="center")

treeview_sell.column("#4", width=70, anchor="center")
treeview_sell.heading("#4", text="매물 층수", anchor="center")

treeview_sell.column("#5", width=70, anchor="center")
treeview_sell.heading("#5", text="매물 칸수", anchor="center")

treeview_sell.column("#6", width=70, anchor="center")
treeview_sell.heading("#6", text="매물 평수", anchor="center")

treeview_sell.column("#7", width=130, anchor="center")
treeview_sell.heading("#7", text="옵션", anchor="center")

treeview_sell.column("#8", width=130, anchor="center")
treeview_sell.heading("#8", text="이사 날짜", anchor="center")

treeview_sell.column("#9", width=90, anchor="center")
treeview_sell.heading("#9", text="특이 사항", anchor="center")

treeview_sell.column("#10", width=150, anchor="center")
treeview_sell.heading("#10", text="집주인/세입자 연락처", anchor="center")

treelist_sell = matrix_from_sheet("Sale_List")
for i in range(len(treelist_sell)):
    treeview_sell.insert("", 'end', text=str(i), values=treelist_sell[i], iid=str(i))

btn2_2_1 = tk.Button(frame2_2, text="뒤로", command=event2_2_1)
btn2_2_2 = tk.Button(frame2_2, text="등록", command=event2_2_2)
btn2_2_3 = tk.Button(frame2_2, text="삭제", command=event2_2_3)


# Page 3_1
def event3_1_1():
    frame3_1.pack_forget()
    lb1_1.pack_forget()
    entry1_1.pack_forget()
    lb1_2.pack_forget()
    entry1_2.pack_forget()
    lb1_3.pack_forget()
    entry1_3.pack_forget()
    lb1_4.pack_forget()
    entry1_4.pack_forget()
    lb1_5.pack_forget()
    entry1_5.pack_forget()
    lb1_6.pack_forget()
    entry1_6.pack_forget()
    lb1_7.pack_forget()
    entry1_7.pack_forget()
    lb1_8.pack_forget()
    entry1_8.pack_forget()
    lb1_9.pack_forget()
    entry1_9.pack_forget()
    lb1_10.pack_forget()
    entry1_10.pack_forget()
    btn3_1_1.pack_forget()
    btn3_1_2.pack_forget()

    frame2_1.pack(expand=True, fill='both')
    treeview_buy.pack()
    btn2_1_1.pack(side=tk.LEFT, anchor='sw')
    btn2_1_2.pack(side=tk.RIGHT, anchor='se')
    btn2_1_3.pack(side=tk.RIGHT, anchor='se')

def event3_1_2():
    input_items = [input_text1_1.get(), input_text1_2.get(), input_text1_3.get(), input_text1_4.get(), input_text1_5.get(), input_text1_6.get(), input_text1_7.get(), input_text1_8.get(), input_text1_9.get(), input_text1_10.get()]
    wb["Purchase_Quest"].append(input_items)
    treeview_buy.delete(*treeview_buy.get_children())
    list_changed = matrix_from_sheet("Purchase_Quest")
    for p in range(len(list_changed)):
        treeview_buy.insert("", 'end', text=str(p), values=list_changed[p], iid=str(p))

    wb.save("Purchase_sale.xlsx")

    frame3_1.pack_forget()
    lb1_1.pack_forget()
    entry1_1.pack_forget()
    lb1_2.pack_forget()
    entry1_2.pack_forget()
    lb1_3.pack_forget()
    entry1_3.pack_forget()
    lb1_4.pack_forget()
    entry1_4.pack_forget()
    lb1_5.pack_forget()
    entry1_5.pack_forget()
    lb1_6.pack_forget()
    entry1_6.pack_forget()
    lb1_7.pack_forget()
    entry1_7.pack_forget()
    lb1_8.pack_forget()
    entry1_8.pack_forget()
    lb1_9.pack_forget()
    entry1_9.pack_forget()
    lb1_10.pack_forget()
    entry1_10.pack_forget()
    btn3_1_1.pack_forget()
    btn3_1_2.pack_forget()

    frame2_1.pack(expand=True, fill='both')
    treeview_buy.pack()
    btn2_1_1.pack(side=tk.LEFT, anchor='sw')
    btn2_1_2.pack(side=tk.RIGHT, anchor='se')
    btn2_1_3.pack(side=tk.RIGHT, anchor='se')


input_text1_1 = tk.StringVar()
input_text1_2 = tk.StringVar()
input_text1_3 = tk.StringVar()
input_text1_4 = tk.StringVar()
input_text1_5 = tk.StringVar()
input_text1_6 = tk.StringVar()
input_text1_7 = tk.StringVar()
input_text1_8 = tk.StringVar()
input_text1_9 = tk.StringVar()
input_text1_10 = tk.StringVar()

frame3_1 = tk.Frame()
lb1_1 = tk.Label(frame3_1, text="희망 지역")
entry1_1 = tk.Entry(frame3_1, textvariable=input_text1_1)
lb1_2 = tk.Label(frame3_1, text="전세/월세")
entry1_2 = tk.Entry(frame3_1, textvariable=input_text1_2)
lb1_3 = tk.Label(frame3_1, text="희망 가격")
entry1_3 = tk.Entry(frame3_1, textvariable=input_text1_3)
lb1_4 = tk.Label(frame3_1, text="희망 층수")
entry1_4 = tk.Entry(frame3_1, textvariable=input_text1_4)
lb1_5 = tk.Label(frame3_1, text="희망 평수")
entry1_5 = tk.Entry(frame3_1, textvariable=input_text1_5)
lb1_6 = tk.Label(frame3_1, text="희망 칸수")
entry1_6 = tk.Entry(frame3_1, textvariable=input_text1_6)
lb1_7 = tk.Label(frame3_1, text="보유 옵션")
entry1_7 = tk.Entry(frame3_1, textvariable=input_text1_7)
lb1_8 = tk.Label(frame3_1, text="이사 가능 날짜")
entry1_8 = tk.Entry(frame3_1, textvariable=input_text1_8)
lb1_9 = tk.Label(frame3_1, text="특이사항")
entry1_9 = tk.Entry(frame3_1, textvariable=input_text1_9)
lb1_10 = tk.Label(frame3_1, text="연락처")
entry1_10 = tk.Entry(frame3_1, textvariable=input_text1_10)

btn3_1_1 = tk.Button(frame3_1, text="뒤로", command=event3_1_1)
btn3_1_2 = tk.Button(frame3_1, text="등록", command=event3_1_2)


# Page 3_2
def event3_2_1():
    frame3_2.pack_forget()
    lb2_1.pack_forget()
    entry2_1.pack_forget()
    lb2_2.pack_forget()
    entry2_2.pack_forget()
    lb2_3.pack_forget()
    entry2_3.pack_forget()
    lb2_4.pack_forget()
    entry2_4.pack_forget()
    lb2_5.pack_forget()
    entry2_5.pack_forget()
    lb2_6.pack_forget()
    entry2_6.pack_forget()
    lb2_7.pack_forget()
    entry2_7.pack_forget()
    lb2_8.pack_forget()
    entry2_8.pack_forget()
    lb2_9.pack_forget()
    entry2_9.pack_forget()
    lb2_10.pack_forget()
    entry2_10.pack_forget()
    btn3_2_1.pack_forget()
    btn3_2_2.pack_forget()

    frame2_2.pack(expand=True, fill='both')
    treeview_sell.pack()
    btn2_2_1.pack(side=tk.LEFT, anchor='sw')
    btn2_2_2.pack(side=tk.RIGHT, anchor='se')
    btn2_2_3.pack(side=tk.RIGHT, anchor='se')

def event3_2_2():
    input_items = [input_text2_1.get(), input_text2_2.get(), input_text2_3.get(), input_text2_4.get(), input_text2_5.get(), input_text2_6.get(), input_text2_7.get(), input_text2_8.get(), input_text2_9.get(), input_text2_10.get()]
    wb["Sale_List"].append(input_items)
    treeview_sell.delete(*treeview_sell.get_children())
    list_changed = matrix_from_sheet("Sale_List")
    for p in range(len(list_changed)):
        treeview_sell.insert("", 'end', text=str(p), values=list_changed[p], iid=str(p))
    wb.save("Purchase_sale.xlsx")


    frame3_2.pack_forget()
    lb2_1.pack_forget()
    entry2_1.pack_forget()
    lb2_2.pack_forget()
    entry2_2.pack_forget()
    lb2_3.pack_forget()
    entry2_3.pack_forget()
    lb2_4.pack_forget()
    entry2_4.pack_forget()
    lb2_5.pack_forget()
    entry2_5.pack_forget()
    lb2_6.pack_forget()
    entry2_6.pack_forget()
    lb2_7.pack_forget()
    entry2_7.pack_forget()
    lb2_8.pack_forget()
    entry2_8.pack_forget()
    lb2_9.pack_forget()
    entry2_9.pack_forget()
    lb2_10.pack_forget()
    entry2_10.pack_forget()
    btn3_2_1.pack_forget()
    btn3_2_2.pack_forget()

    frame2_2.pack(expand=True, fill='both')
    treeview_sell.pack()
    btn2_2_1.pack(side=tk.LEFT, anchor='sw')
    btn2_2_2.pack(side=tk.RIGHT, anchor='se')
    btn2_2_3.pack(side=tk.RIGHT, anchor='se')

input_text2_1 = tk.StringVar()
input_text2_2 = tk.StringVar()
input_text2_3 = tk.StringVar()
input_text2_4 = tk.StringVar()
input_text2_5 = tk.StringVar()
input_text2_6 = tk.StringVar()
input_text2_7 = tk.StringVar()
input_text2_8 = tk.StringVar()
input_text2_9 = tk.StringVar()
input_text2_10 = tk.StringVar()

frame3_2 = tk.Frame()
lb2_1 = tk.Label(frame3_2, text="매물 주소")
entry2_1 = tk.Entry(frame3_2, textvariable=input_text2_1)
lb2_2 = tk.Label(frame3_2, text="전세/월세")
entry2_2 = tk.Entry(frame3_2, textvariable=input_text2_2)
lb2_3 = tk.Label(frame3_2, text="매물 가격")
entry2_3 = tk.Entry(frame3_2, textvariable=input_text2_3)
lb2_4 = tk.Label(frame3_2, text="매물 층수")
entry2_4 = tk.Entry(frame3_2, textvariable=input_text2_4)
lb2_5 = tk.Label(frame3_2, text="매물 칸수")
entry2_5 = tk.Entry(frame3_2, textvariable=input_text2_5)
lb2_6 = tk.Label(frame3_2, text="매물 평수")
entry2_6 = tk.Entry(frame3_2, textvariable=input_text2_6)
lb2_7 = tk.Label(frame3_2, text="보유 옵션")
entry2_7 = tk.Entry(frame3_2, textvariable=input_text2_7)
lb2_8 = tk.Label(frame3_2, text="이사 가능 날짜")
entry2_8 = tk.Entry(frame3_2, textvariable=input_text2_8)
lb2_9 = tk.Label(frame3_2, text="특이사항")
entry2_9 = tk.Entry(frame3_2, textvariable=input_text2_9)
lb2_10 = tk.Label(frame3_2, text="집주인/세입자 연락처")
entry2_10 = tk.Entry(frame3_2, textvariable=input_text2_10)

btn3_2_1 = tk.Button(frame3_2, text="뒤로", command=event3_2_1)
btn3_2_2 = tk.Button(frame3_2, text="등록", command=event3_2_2)

window.mainloop()