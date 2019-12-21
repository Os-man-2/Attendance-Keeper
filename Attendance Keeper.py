from Tkinter import *
import xlrd
import ttk
import tkFileDialog
from xlwt import Workbook

class Tool(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.InitUI()

    def InitUI(self):

        self.pack()

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1) #Adjustment for the gui
        self.columnconfigure(2, weight=1) #Adjustment for the gui
        self.columnconfigure(3, weight=1) #Adjustment for the gui
        self.columnconfigure(4, weight=1)

        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1) #Adjustment for the gui
        self.rowconfigure(3, weight=1) #Adjustment for the gui
        self.rowconfigure(4, weight=1) #Adjustment for the gui
        self.rowconfigure(5, weight=1) #Adjustment for the gui
        self.rowconfigure(6, weight=1)

        self.label1 = Label(self, text='AttendanceKeeper v1.0', font=('', '20', 'bold')) #main label
        self.label1.grid(row=0, column=0, columnspan=5, sticky=NSEW)

        self.label2 = Label(self, text='Select student list Excel file:', font=('', '14', 'bold')) #Select student list Excel file Label
        self.label2.grid(row=1, column=0, columnspan=1, sticky=NSEW)

        self.label3 = Label(self, text='Select a Student:', anchor="w", font=('', '14', 'bold')) #Select a student Label
        self.label3.grid(row=2, column=0, sticky=NSEW)

        self.label4 = Label(self, text='Section:', font=('', '14', 'bold')) #Section Label
        self.label4.grid(row=2, column=2, sticky=NSEW)

        self.label5 = Label(self, text='Attended Students:', font=('', '12', 'bold'))#Attended students Label
        self.label5.grid(row=2, column=3, sticky=NSEW)

        self.label6 = Label(self, text='Please select file type:', anchor='w', font=('', '12', 'bold'))#Please select file type Label
        self.label6.grid(row=6, column=0, sticky=NSEW)

        self.label7 = Label(self, text='Please enter week:', font=('', '12', 'bold'))#Please enter week Label
        self.label7.grid(row=6, column=3, sticky=NSEW)

        self.import_list = Button(self, text="Import List", width=1, command=self.dialogue)#import_list button
        self.import_list.grid(row=1, column=2, sticky=NSEW)

        self.list_box1 = Listbox(self,selectmode='multiple') #left listbox
        self.list_box1.grid(row=3, rowspan=3, column=0, columnspan=2, sticky=NSEW)

        self.list_box2 = Listbox(self,selectmode='multiple') #right listbox
        self.list_box2.grid(row=3, rowspan=3, column=3,columnspan=3,sticky=NSEW)

        self.scrollbar1 = Scrollbar(self, orient=VERTICAL) #list_box1's scrollbar
        self.list_box1.config(yscrollcommand=self.scrollbar1.set)
        self.scrollbar1.config(command=self.list_box1.yview)
        self.scrollbar1.grid(row=3,rowspan=3,column=1,sticky=NSEW)

        self.scrollbar2 = Scrollbar(self,orient=VERTICAL) #list_box2's scrollbar
        self.list_box2.config(yscrollcommand=self.scrollbar2.set)
        self.scrollbar2.config(command=self.list_box2.yview)
        self.scrollbar2.grid(row=3,rowspan=3,column=5,sticky=E+N+S)

        self.entry1 = Entry(self)
        self.entry1.grid(row=6, column=4, sticky=NSEW)

        self.button1 = Button(self, text='Export as file',command = save.make_afile)#Export as file button
        self.button1.grid(row=6, column=5, sticky=NSEW)

        self.button2 = Button(self, text='Add =>',command=A_R.add)#Add button
        self.button2.grid(row=4, column=2, sticky=NSEW)

        self.string = StringVar()
        self.string_1 = StringVar()
        self.n = ttk.Combobox(self, textvariable=self.string)#Combobox
        self.n.bind("<<ComboboxSelected>>",self.load_models)
        self.com1 = ttk.Combobox(self, textvariable=self.string_1, width=10)
        self.n.grid(row=3, column=2, sticky=NSEW)
        self.com1.grid(row=6, column=1, sticky=NSEW)
        self.com1['value'] = ('csv',"txt","xlsx")
        self.com1.current(1) # to set the initial value

        self.remove = Button(self, text='<= Remove',command=A_R.remove)
        self.remove.grid(row=5, column=2, sticky=NSEW)


    def dialogue(self):

        self.filename = tkFileDialog.askopenfilename(initialdir="/", title="Select Student list Excel File",filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        self.excel = xlrd.open_workbook(self.filename)

        sheet = self.excel.sheet_by_index(0)  # or by the index it has in excel's sheet collection
        self.sheet = sheet
        r = sheet.row(0)  # returns all the CELLS of row 0,
        c = sheet.col_values(0)  # returns all the VALUES of row 0,

        self.data2 = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        self.datai = {
            'ENGR 102 01': [i[1] for i in self.data2 if i[3] == 'ENGR 102 01'],
            'ENGR 102 02': [i[1] for i in self.data2 if i[3] == 'ENGR 102 02'],
            'ENGR 102 03': [i[1] for i in self.data2 if i[3] == 'ENGR 102 03'],
            'ENGR 102 04': [i[1] for i in self.data2 if i[3] == 'ENGR 102 04'],
            'ENGR 102 05': [i[1] for i in self.data2 if i[3] == 'ENGR 102 05'],
            'ENGR 102 06': [i[1] for i in self.data2 if i[3] == 'ENGR 102 06'],
            'ENGR 102 07': [i[1] for i in self.data2 if i[3] == 'ENGR 102 07'],
            'ENGR 102 08': [i[1] for i in self.data2 if i[3] == 'ENGR 102 08'],
            'ENGR 102 09': [i[1] for i in self.data2 if i[3] == 'ENGR 102 09'],#All of the students seperated by their sections
            'ENGR 102 12': [i[1] for i in self.data2 if i[3] == 'ENGR 102 12'],
            'ENGR 102 13': [i[1] for i in self.data2 if i[3] == 'ENGR 102 13'],
            'ENGR 102 14': [i[1] for i in self.data2 if i[3] == 'ENGR 102 14'],
            'ENGR 102 15': [i[1] for i in self.data2 if i[3] == 'ENGR 102 15'],
            'ENGR 102 16': [i[1] for i in self.data2 if i[3] == 'ENGR 102 16'],
            'ENGR 102 18': [i[1] for i in self.data2 if i[3] == 'ENGR 102 18'],
            'ENGR 102 19': [i[1] for i in self.data2 if i[3] == 'ENGR 102 19']
        }

        self.data = []
        for i in xrange(sheet.nrows):
            if sheet.row_values(i)[-1] != "Section" and sheet.row_values(i)[-1] not in self.data:
                self.data.append(sheet.row_values(i)[-1])  #contains section names after filtering
        self.n["values"] = self.data
        self.n.current(5) #engr 102 01 is in 5th row
        self.choice()


    def choice(self):
        self.section_choice = self.n.get()
        for j in xrange(self.sheet.nrows):
            if self.sheet.row_values(j)[-1] == self.section_choice:
                names = self.sheet.row_values(j)
                names2 = names[1].split(" ")
                self.list_box1.insert(END,(names2[0],",",names2[1],",",int(names[0])))

    def load_models(self,*args):
        selection = self.n.selection_get() #Keeping the data
        self.list_box1.delete(0, END) #Gives us the ability to show students from whatever section , CLEARS THE OTHER SECTIONS students from it
        self.list_box2.delete(0,END) #When we select another section list_box2 will be clear


        for i, item in enumerate(self.datai[selection]): #Gives us the ability to show students from whatever section we want
            self.list_box1.insert(i, item)

class AddRemove:
    def add(self):
        global app
        index = app.list_box1.curselection() #It returns a list of item indexes
        for i in index:
            name_to_send = app.list_box1.get(i) #keeping the data
            app.list_box2.insert(END,name_to_send) #Gives us the ability to add any student to the listbox

    def remove(self):
        index = app.list_box2.curselection()
        for j in index:
            name_to_send2 = app.list_box2.get(j) #Gives us the ability to remove any selected student from the listbox
        app.list_box1.insert(END,name_to_send2)
        app.list_box2.delete(END,j)


class Composefile:
    def make_afile(self):
        global app
        if app.com1.get() == 'txt':
            app.values_in_listbox = app.list_box2.get(0, END) #Checking if what file type is chosen by the user and according to that entering or writing the attended students  into txt file
            app.dosya = open('tam.txt','w')
            for i in app.values_in_listbox:
                app.dosya.write(str(i[4]) + ' ' + i[0].encode('utf-8') + ' '+i[2].encode('utf-8') + '\n') #for turkish characthers errors
        elif app.com1.get() == 'xlsx': #Checking if what file type is chosen by the user and according to that entering or writing the attended students into xls (excel) file
            app.dosya = Workbook()
            app.sayfa = app.dosya.add_sheet('sayfa 1')
            index = 0
            app.values_in_listbox = app.list_box2.get(0, END)
            for i in app.values_in_listbox:
                app.sayfa.write(index,0, str(i[4]))
                app.sayfa.write(index,1, i[0] +' '+i[2])
                index+=1
            app.dosya.save('test.xls')

        elif app.com1.get() == 'csv':
            raise BaseException("File type is not supported ") #raising an exception if csv selected



def main():
    global app
    root = Tk()
    app = Tool(root)
    app.grid(sticky=NSEW)
    root.mainloop()

A_R=AddRemove()
save = Composefile()
main() #having the whole program worked

