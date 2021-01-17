import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter
#from Tkinter import Tk  # from tkinter import Tk for Python 3.x
#from tkinter.filedialog import askopenfilename
import json

def load_df1():
    # Use a breakpoint in the code line below to debug your script.
    return pd.read_csv(list[0], delimiter=';')
def load_df2():
    # Use a breakpoint in the code line below to debug your script.
    return pd.read_csv(list[1], delimiter=';')
def load_df3():
    # Use a breakpoint in the code line below to debug your script.
    return pd.read_csv(list[2], delimiter=';')

if __name__ == '__main__':


    #Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    #filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    #print(filename)
    #tkinter.filedialog.askopenfilename()

    with open('json_data.json') as f:
        company_data = json.load(f)
    #company_data = json.loads(json_data.json)
    print(type(company_data['sample']))

    list = []
    for sam in company_data['sample']:
        print(sam['path'])
        list.append(sam['path'])

    print(type(list[0]))
    df1 = load_df1()
    df2 = load_df2()
    df3 = load_df3()

    number_rows = len(df1.index)
    print(number_rows)

    # Add some summary data using the new assign functionality in pandas 0.16
    df1 = df1.assign(Aver=lambda x: (x.node1_temp + x.node2_temp + x.node3_temp)/3)
    df1 = df1.assign(cooling_rate=lambda y: (abs(y.Aver.sub(df1.Aver.shift(1), fill_value=0))))
    df2 = df2.assign(Aver=lambda x: (x.node1_temp + x.node2_temp + x.node3_temp) / 3)
    df2 = df2.assign(cooling_rate=lambda y: (abs(y.Aver.sub(df2.Aver.shift(1), fill_value=0))))
    df3 = df3.assign(Aver=lambda x: (x.node1_temp + x.node2_temp + x.node3_temp) / 3)
    df3 = df3.assign(cooling_rate=lambda y: (abs(y.Aver.sub(df3.Aver.shift(1), fill_value=0))))
    df1.loc[0, 'cooling_rate'] = 0
    df2.loc[0, 'cooling_rate'] = 0
    df3.loc[0, 'cooling_rate'] = 0

    print(df1.loc[0,'Aver'])
    #df1['cooling_rate'] = abs(df1.Aver.sub(df1.Aver.shift(1), fill_value=0))
    df1.loc[0,'cooling_rate'] = 0
    #df1['cooling_rate'] = abs(df1.Aver.sub(df1.Aver.shift(), fill_value=0).astype(int))
    print(df1)

    #df1 = df1.assign(cooling_rate=lambda y: diff(y.Aver)/diff(y.time))

    df = pd.concat([df1, df2, df3],axis=1, names = ['a','b', 'c'])
    writer = pd.ExcelWriter(list[3], engine='xlsxwriter')

    df.to_excel(writer, sheet_name='sheet1')

    workbook = writer.book
    worksheet = writer.sheets['sheet1']
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        'name': 'dT/dt to t',
        'name_font': {'size': 12, 'bold': True},
        'categories': '=sheet1!$B$2:$B$21',
        'values': '=sheet1!$F$2:$F$21',
    })
    chart.set_x_axis({
        'name': 'Time t [s]',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
    })
    chart.set_y_axis({
        'name': 'Temperature T [C]',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
    })

    worksheet1 = writer.sheets['sheet1']
    chart1 = workbook.add_chart({'type': 'line'})
    chart1.add_series({
        'name': 'dT/dt to t',
        'name_font': {'size': 12, 'bold': True},
        'categories': '=sheet1!$H$2:$H$21',
        'values': '=sheet1!$L$2:$L$21',
    })
    chart1.set_x_axis({
        'name': 'Time t [s]',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
        # 'reverse': True
    })


    chart1.set_y_axis({
        'name': 'Temperature T [C]',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
        # 'reverse': True
    })

    worksheet2 = writer.sheets['sheet1']
    chart2 = workbook.add_chart({'type': 'line'})
    chart2.add_series({
        'name': 'dT/dt to t',
        'name_font': {'size': 12, 'bold': True},
        'categories': '=sheet1!$N$2:$N$21',
        'values': '=sheet1!$R$2:$R$21',
    })
    chart2.set_x_axis({
        'name': 'Time t [s]',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
    })
    chart2.set_y_axis({
        'name': 'Temperature T [C]',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
    })

    chart.add_series({
        'name': 'T to dT/dt',
        'values': '=sheet1!$G$2:$G$21',
        'categories': '=sheet1!$F$2:$F$21',
        'X2_axis': True,
        'y2_axis': True,

    })
    chart.set_x2_axis({
        'name': 'Temperature T [C]',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
        'visible': True,
        #'reverse': True

    })
    chart.set_y2_axis({
        'name': 'Cooling rate, °C s^-1',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
        'visible': True,
        'reverse': True
    })


    chart1.add_series({
        'name': 'T to dT/dt',
        'values': '=sheet1!$M$2:$M$21',
        'categories': '=sheet1!$L$2:$L$21',
        'x2_axis': True,
        'y2_axis': True

    })
    chart1.set_x2_axis({
        'name': 'Temperature T [C]',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
        'visible': True
    })
    chart1.set_y2_axis({
        'name': 'Cooling rate, °C s^-1',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
        'visible': True,
        'reverse': True
    })
    chart2.add_series({
        'name': 'T to dT/dt',
        'values': '=sheet1!$S$2:$S$21',
        'categories': '=sheet1!$R$2:$R$21',
        'x2_axis': True,
        'y2_axis': True

    })

    chart2.set_x2_axis({
        'name': 'Temperature T [C]',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
        'visible': True

    })
    chart2.set_y2_axis({
        'name': 'Cooling rate, °C s^-1',
        'name_font': {'size': 10, 'bold': True},
        'num_font': {'italic': True},
        'visible': True,
        'reverse': True
    })
    chart.set_style(45)
    chart1.set_style(46)
    chart2.set_style(47)

    worksheet.insert_chart(number_rows+1, 0, chart)
    worksheet1.insert_chart(2*number_rows, 0, chart1)
    worksheet2.insert_chart(3 * number_rows, 0, chart2)

    writer.save()

