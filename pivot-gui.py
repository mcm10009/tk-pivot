import importlib
import tkinter as tk
from tkinter import filedialog
import pandas as pd


def open_click():

        f = list(filedialog.askopenfilenames(parent=window, title='Select File'))
        if len(f) !=0:
                df = pd.read_excel(f[0], header=None)

                global imported_module
                comp_name = df.iloc[0,0].split(' ')[0].replace(',','')
                imported_module = importlib.import_module(comp_name)

                name = df.iloc[0,0] + ' ' + df.iloc[1,0] + ' ' + df.iloc[2,0]
                name = name.replace('  ',' ')
                cols = df.iloc[4,1:].tolist()
                data_set = df.drop([0,1,2,3,4]).reset_index(drop=True)
                data_set = data_set[data_set[1] != "Beginning Balance"]
                window.title(name)

                frame = []
                global frame_dict
                frame_dict = {}
                frame_dict_keys = []


                for idx, line in enumerate(data_set[1]):
                        if (pd.isna(line) == False) and (pd.isna(data_set.iloc[idx-1,1])):
                                frame_title = data_set.iloc[idx-1,0].strip()
                                frame_dict_keys.append(frame_title)
                                frame.append(data_set.iloc[idx,:])
                        elif pd.isna(line) == False:
                                frame.append(data_set.iloc[idx,:])
                        elif (idx != 0) and (pd.isna(data_set.iloc[idx-1,1]) == False):
                                temp_frame = pd.DataFrame(frame).drop(0,axis=1).reset_index(drop=True)
                                temp_frame.columns = cols
                                frame_dict[frame_title] = temp_frame
                                frame = []

                
                lst = sel.get().split(',')
                quit_button.destroy()

                new_quit_button = tk.Button(window, text='Exit', command=window.destroy)
                new_quit_button.grid(row=1, column=2, sticky='ew')

                if len(button_list) != 0:
                        for b in button_list:
                                b.destroy()

                num = int(len(frame_dict_keys)/3)
                num1 = 2*num
                
                
                for idx, item in enumerate(frame_dict_keys):
                        for stuff in lst:
                                if stuff in item:
                                        if idx <= num:
                                                button = tk.Button(window, text=item, borderwidth=1, command=lambda choice=item: button_click(choice))
                                                button.grid(row=idx+1+2,column=0, sticky='ew')
                                                button_list.append(button)
                                        elif idx <= num1:
                                                button = tk.Button(window, text=item, borderwidth=1, command=lambda choice=item: button_click(choice))
                                                button.grid(row=idx-num+2,column=1, sticky='ew')
                                                button_list.append(button)
                                        else:
                                                button = tk.Button(window, text=item, borderwidth=1, command=lambda choice=item: button_click(choice))
                                                button.grid(row=idx-num1+2,column=2, sticky='ew')
                                                button_list.append(button)


                global pivot_number
                my_label = tk.Label(window, fg='red', text='Max Number Of Columns: ')
                my_label.grid(row=1, column=0, sticky='ew')
                pivot_number = tk.Entry(window, fg='red')
                pivot_number.grid(row=1, column=1, sticky='ew')
                pivot_number.insert(0,10)
                return frame_dict, imported_module


def button_click(choice):

        data = frame_dict[choice]
        vendor_list_1 = imported_module.first_parser(data)
        vendor_list_2 = imported_module.second_parser(data)
        vendors = imported_module.third_parser(vendor_list_1,vendor_list_2)

        data['Vendors'] = vendors
        data['Amount'] = pd.to_numeric(data['Amount'])
        data['Month'] = pd.to_datetime(data['Date']).dt.month
        data['Month'] = data['Month'].apply(lambda x: imported_module.months[x])

        pvt = pd.pivot_table(data, values=['Amount'], index=['Vendors'], columns=['Month'], aggfunc=sum)
        pvt= pvt['Amount']

        new_col = []
        for index in range(1,13):
                for element in pvt.columns:
                        if imported_module.months[index] == element:
                                new_col.append(element)
                                break
        pvt.columns = new_col

        pvt['Mean'] = pvt.mean(axis=1).round(2)
        pvt = pvt.sort_values(by=['Mean'], ascending=False)

        pcnt = 0.25
        def highlight_mean_gt(s):
                x = round(s.mean(),2)
                gt_mean = s > (x+pcnt*x)
                return ['background-color: #FF7F7F' if v else '' for v in gt_mean]

        def highlight_mean_lt(s):
                x = round(s.mean(),2)
                lt_mean = s < (x - pcnt*x)
                return ['background-color: #7CC089' if v else '' for v in lt_mean]

        def color_null(s):
                null_values = s.isnull()
                return ['color: gray' if v else '' for v in null_values]

        def highlight_cols(s):
                color = '#EAEC81'
                return 'background-color: %s' % color

        sub = pvt.style\
        .apply(highlight_mean_gt,axis=1)\
        .apply(highlight_mean_lt,axis=1)\
        .apply(color_null,axis=1)\
        .applymap(highlight_cols, subset=pd.IndexSlice[:,['Mean']])\
        .highlight_null(null_color='#CBFEFE')\
        .format('${:,.2F}')

        pvt = pvt.reset_index()
        pvt = pvt.head(int(pivot_number.get()))

        def pivot_click():
                sub.to_excel('~/Desktop/test_pivot_table.xlsx')


        r,c = pvt.shape

        pop_up = tk.Toplevel()
        pop_up.attributes('-topmost', 'true')
        pop_up.title(choice + ' (' + str(r) + ' ' + 'Rows)')

        pop_up.rowconfigure(0, weight=0)

        canvas1 = tk.Canvas(pop_up, height=25, width=110*c)
        non_scroll_frame = tk.Frame(canvas1)

        canvas2 = tk.Canvas(pop_up, height=35*r, width=110*c)
        scroll_frame = tk.Frame(canvas2, bd=10, relief='ridge')

        scroll_y = tk.Scrollbar(pop_up, orient='vertical', command=canvas2.yview)
        scroll_x = tk.Scrollbar(pop_up, orient='horizontal', command=canvas2.xview)

        for idx, item in enumerate(pvt.columns):
                tk.Label(scroll_frame, text=item, font=('Lucia Grande', 13, 'bold'), borderwidth=1).grid(row=0,column=idx+1,  sticky='ew')

        for i in range(1,r+1):
                for j in range(1,c+1):
                        compare = pvt.iloc[i-1,c-1]
                        if i % 2 != 0:
                                if j == c:
                                        tk.Label(scroll_frame, bg='#EAEC81', text='${:,.2f}'.format(pvt.iloc[i-1,j-1]), borderwidth=1).grid(row=i,column=j, sticky='ew')
                                elif (j != 1) and (pd.isna(pvt.iloc[i-1,j-1])):
                                                tk.Label(scroll_frame, bg='lightgray', fg='black', text='x', borderwidth=1).grid(row=i,column=j, sticky='ew')
                                elif (j != 1) and (pvt.iloc[i-1,j-1] > compare + 0.25*compare):
                                        tk.Label(scroll_frame, bg='#FF7F7F', text='${:,.2f}'.format(pvt.iloc[i-1,j-1]), borderwidth=1).grid(row=i,column=j, sticky='ew')
                                elif (j != 1) and (pvt.iloc[i-1,j-1] < compare - 0.25*compare):
                                        tk.Label(scroll_frame, bg='#7CC089', text='${:,.2f}'.format(pvt.iloc[i-1,j-1]), borderwidth=1).grid(row=i,column=j, sticky='ew')
                                elif (j != 1):
                                        tk.Label(scroll_frame, bg='lightgray', text='${:,.2f}'.format(pvt.iloc[i-1,j-1]), borderwidth=1).grid(row=i,column=j, sticky='ew')
                                else:
                                        tk.Label(scroll_frame, bg='lightgray', text=pvt.iloc[i-1,j-1], borderwidth=1).grid(row=i,column=j, sticky='ew')
                        else:
                                if j == c:
                                        tk.Label(scroll_frame, bg='#EAEC81', text='${:,.2f}'.format(pvt.iloc[i-1,j-1]), borderwidth=1).grid(row=i,column=j, sticky='ew')
                                elif (j != 1) and (pd.isna(pvt.iloc[i-1,j-1])):
                                                tk.Label(scroll_frame, bg='white', fg='black', text='x', borderwidth=1).grid(row=i,column=j, sticky='ew')
                                elif (j != 1) and (pvt.iloc[i-1,j-1] > compare + 0.25*compare):
                                        tk.Label(scroll_frame, bg='#FF7F7F', text='${:,.2f}'.format(pvt.iloc[i-1,j-1]), borderwidth=1).grid(row=i,column=j, sticky='ew')
                                elif (j != 1) and (pvt.iloc[i-1,j-1] < compare - 0.25*compare):
                                        tk.Label(scroll_frame, bg='#7CC089', text='${:,.2f}'.format(pvt.iloc[i-1,j-1]), borderwidth=1).grid(row=i,column=j, sticky='ew')
                                elif (j != 1):
                                        tk.Label(scroll_frame, text='${:,.2f}'.format(pvt.iloc[i-1,j-1]), borderwidth=1).grid(row=i,column=j, sticky='ew')
                                else:
                                        tk.Label(scroll_frame, text=pvt.iloc[i-1,j-1], borderwidth=1).grid(row=i,column=j, sticky='ew')
                                        
        pivot_button = tk.Button(non_scroll_frame, text='Save Styled Pivot Table', borderwidth=1, command=pivot_click)
        pivot_button.pack(fill='both', expand=True, side='left')

        canvas1.create_window((0,0), anchor='nw', window=non_scroll_frame)
        canvas1.pack(fill='both', expand=True, side='top')
        
        canvas2.create_window((0, 0), anchor='nw', window=scroll_frame)
        canvas2.update_idletasks()
        canvas2.configure(scrollregion=canvas2.bbox('all'))
        canvas2.configure(yscrollcommand=scroll_y.set)
        canvas2.configure(xscrollcommand=scroll_x.set)
        scroll_y.pack(fill='y', side='right')
        scroll_x.pack(fill='x', side='bottom')
        canvas2.pack(fill='both', expand=True, side='bottom')


window = tk.Tk()
window.title('Pivot Demo')
#window.minsize(530,20)

lab = tk.Label(window, fg='red', text='Choose Sections: ')
lab.grid(row=0, column=0, sticky='ew')

sel = tk.Entry(window, fg='red')
sel.grid(row=0, column=1, sticky='ew')

global quit_button
global button_list
button_list = []

open_button = tk.Button(window, text='File Select', command=open_click)
open_button.grid(row=0, column=2, sticky='ew')

quit_button = tk.Button(window, text='Exit', command=window.destroy)
quit_button.grid(row=0, column=3, sticky='ew')

window.mainloop()
