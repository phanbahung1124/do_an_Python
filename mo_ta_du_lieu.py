import PySimpleGUI as sg
import xlwings as xw

# doc file excel
wb = xw.Book('D:/Pycharm/PycharmProjects/Temperature.xlsx')

# doc sheet excel
sht = wb.sheets('Sheet1')

# lay hang cuoi trong cot A cua sheet1
lr = sht.range("A" + str(sht.cells.last_cell.row)).end("up").row

# lay cot cuoi trong dong 1 cua sheet1
lc = sht["A1"].expand("right").last_cell.address.split("$")[1]

# tao list Year
list_year = sht.range(f"A2:A{lr}").value
li_year = list(set(list_year))

# tao list Month
list_month = sht.range(f"B2:B{lr}").value

# tao list cac Season
list_season = sht.range(f"C2:C{lr}").value

#lay cac gia tri duy nhat trong list season
li_season = list(set(list_season))
li_season.insert(0,"ALL")

# tao list Salinity
list_salinity = sht.range(f"D2:D{lr}").value

# tao list Temperature
list_temperature = sht.range(f"E2:E{lr}").value

# tao list CHLFa
list_CHlFa = sht.range(f"F2:F{lr}").value

#tao list cac Area
list_area = sht.range(f"G2:G{lr}").value

#lay cac gia tri duy nhat trong list area
li_area = list(set(list_area))
li_area.insert(0,"ALL")

# lay tieu de cua sheet1
data_headings = sht.range(f"A1:{lc}1").value

#lay gia tri trong sheet1 de hien thi len table
data_values_disp = sht.range(f"A2:{lc}{lr}").value

#tao giao dien
sg.theme("LightGrey6")

layout_frame = [
    [sg.Text("Year", size=(5,1)), sg.Input("1990",size = (11,1),key = "Year"),
     sg.Text("Season",size=(6,1)),sg.Combo(li_season,default_value = "ALL",key = "Season",expand_x = True)],

     [sg.Text("Month", size=(8,1)), sg.Input("1",size = (15,1),key = "Month"),
     sg.CalendarButton(button_text="Month", target="Start Month", format = "%d/%m/%Y"),
     sg.Text("Area",size=(4,1)),sg.Combo(li_area,default_value = "ALL",key = "Area",expand_x = True)],

     [sg.Text("Sal", size=(2,1)),sg.Input("0",size = (10,1),key = "Sal"),
      sg.Text("Te",size= (2,1)),sg.Input("0",size = (10,1),key = "Te"),
      sg.Text("CH",size= (2,1)),sg.Input("0",size = (10,1),key = "CH")],

      [sg.Button("Search",key="Search",size=(8,1))]
]

layout = [
 [sg.Frame("ADVANCED FILTER", layout_frame, size = (400,150))],

 [sg.Table(values=data_values_disp, headings=data_headings,
           background_color="#D1AAAA",
           num_rows= 20,
           max_col_width=20,
           justification="Center",
           text_color="Black",
           header_background_color="#9A3D3D",
           header_font=("Arial",11),
           key='_filestable_',
           expand_x=True,expand_y=True
           )]
]

window = sg.Window("ADVANCED FILTER", layout, size =(480,540))

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == "Exit":
        break