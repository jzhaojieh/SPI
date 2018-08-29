import tkinter as tk                
from tkinter import filedialog
from tkinter.ttk import *
from tkinter import messagebox
from tkinter.simpledialog import askstring 
import copy, getpass, requests, os, shutil, re, itertools, adodbapi, win32api, io, datetime
from time import gmtime, strftime,localtime
from multiprocessing import Queue
from PIL import Image
from PIL import ImageTk
from tkinter import StringVar, Entry, Frame, Listbox, Scrollbar
from tkinter.constants import *
user =''
passwordi = ''
departmentDict = {}
employees = []
employdict = {}
gl_descriptions = []
gl_codes = {}
gl_inputs = {}
entryParts = {}
purchaseOrders = {}
supplierID = {}
commodityCodes = []
temp = []
temp2 = []
t3 = []
lineItems = {}
dirs = {}
approvalGroupsExpense = {1:[],2:[],3:[],4:[],5:[],6:[],9:[],10:[],13:[],14:[]}
approvalGroupsInventory = {1:[],2:[],3:[],6:[],10:[],11:[],12:[]}
purchaseLimitsExpense = {1:'',2:'',3:'',4:'',5:'',6:'',9:'',10:'',13:'',14:''}
purchaseLimitsInventory = {1:'',2:'',3:'',6:'',10:'',11:'',12:''}
managers = []
purchasing = ['huangjos']
financeApprov = []
itApprover = 'stanistreetlu'
# itApprover = 'huangjos'

user = getpass.getuser()
print(getpass.getuser().lower())
myhost = 'SRV57DB2'
mydb = 'PurchaseReq'
myuser = 'mcs_client'
mypass = 'binarydevelopments'

conn = adodbapi.connect("PROVIDER=SQLOLEDB;Data Source={0};Database={1}; \
             trusted_connection=yes;UID={2};PWD={3};".format('srv57db2','PurchaseReq','mcs_client','binarydevelopments'))
#conn = pypyodbc.connect('DSN=PurchaseReq')
cur = conn.cursor()

conn3 = adodbapi.connect("PROVIDER=SQLOLEDB;Data Source={0};Database={1}; \
             trusted_connection=yes;UID={2};PWD={3};".format('srv57db2','InternalEmailer','mcs_client','binarydevelopments'))
cur3 = conn3.cursor()
conn2 = adodbapi.connect("PROVIDER=SQLOLEDB;Data Source={0};Database={1}; \
             UID={2};PWD={3};".format('srv57db1','WM1coTest','mcs_client','binarydevelopments'))
cur2 = conn2.cursor()

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def getDicts():
    cur.execute("""select Group_1,Group_2,Group_3,Group_5,Group_6,Group_9,Group_10,Group_11,Group_12,Group_13,Group_14,Group_4 from Approval_groups""")
    for a in cur.fetchall():
        if a[0]:
            approvalGroupsExpense[1].append(a[0])
            approvalGroupsInventory[1].append(a[0])
        if a[1]:
            approvalGroupsExpense[2].append(a[1])
            approvalGroupsInventory[2].append(a[1])
        if a[2]:
            approvalGroupsExpense[3].append(a[2])
            approvalGroupsInventory[3].append(a[2])
        if a[3]:
            approvalGroupsExpense[5].append(a[3])
        if a[4]:
            approvalGroupsExpense[6].append(a[4])
            approvalGroupsInventory[6].append(a[4])
        if a[5]:
            approvalGroupsExpense[9].append(a[5])
        if a[6]:
            approvalGroupsExpense[10].append(a[6])
            approvalGroupsInventory[10].append(a[6])
        if a[9]:
            approvalGroupsExpense[13].append(a[9])
        if a[10]:
            approvalGroupsExpense[14].append(a[10])
        if a[7]:
            approvalGroupsInventory[11].append(a[7])
        if a[8]:
            approvalGroupsInventory[12].append(a[8])
        if a[11]:
            approvalGroupsExpense[4].append(a[11])
    cur.execute("""select id, poreq_expense,poreq_inventory from signatory""")
    for a in cur.fetchall():
        if 'unlimited' in a[1]:
            purchaseLimitsExpense[a[0]] = float('inf')
        elif a[1] != '-' and 'unlimited':
            purchaseLimitsExpense[a[0]] = a[1]
        if 'unlimited' in a[2]:
            purchaseLimitsInventory[a[0]] = float('inf')
        elif a[2] != '-' and 'unlimited':
            purchaseLimitsInventory[a[0]] = a[2]
    a = list(itertools.chain(*approvalGroupsExpense.values())) + list(itertools.chain(*approvalGroupsInventory.values()))
    b = (list(set(a)))
    c = []
    global managers
    for elem in b:
        lname,fname = elem[elem.index(' ') + 1:].strip(),elem[:elem.index(' ')].strip()
        cur.execute("""Select userid from Employees2 where last_name = '%s' and first_name = '%s'"""%(lname,fname))
        for a in cur.fetchall():
            c.append(a[0])
    managers = c
    global commodityCodes, supplierID,employees,departmentDict,gl_descriptions,gl_codes,employdict
    #creates list with commodity codes and descriptions 
    cur2.execute("""select CommodityCodeId, CommodityCodeDescription from CommodityCodes""")
    for a in cur2.fetchall():
        commodityCodes.append(a[0])
    #creates dictionary with Supplier Names to Supplier IDs
    cur2.execute("""select SupplierID,SupplierName from Suppliers where 
      SupplierId like 'S%' and 
      SupplierName not like '%not use%' and 
      Address <> '' and 
      Address <> 'staff'
    order by
      SupplierName """)
    for a in cur2.fetchall():
        supplierID[a[1]] = a[0]

    #pulls employee names and adds them to global list
    cur.execute("""SELECT last_name FROM Employees2""")
    for a in cur.fetchall():
        for field in a:
            temp.append(field)

    #creates list with employee names
    count = 0
    cur.execute("""SELECT first_name FROM Employees2""")
    for a in cur.fetchall():
        for field in a:
            temp[count] += ", " + field
        count += 1
    employees = copy.deepcopy(temp)
    #adds department descriptions to global list
    cur.execute("""SELECT dept_description FROM Departments2""")
    for d in cur.fetchall():
        for field in d:
            temp2.append(field)
    for i in range(0, len(temp2)):
        departmentDict[i+1] = temp2[i]

    #adds gl descriptions to global list
    cur.execute("""SELECT gldescrip FROM GLCodes2""")
    for a in cur.fetchall():
        for field in a:
            t3.append(field)
    gl_descriptions = copy.deepcopy(t3)

    #addes employees to finance approval list
    cur.execute("""Select userid from employees2 where finance = 'y'""")
    for a in cur.fetchall():
        financeApprov.append(a[0])

    #adds employees to purchasing list
    cur.execute("""Select userid from employees2 where purchasing = 'y'""")
    for a in cur.fetchall():
        purchasing.append(a[0])

    #employee dict maps employee name to employee number
    for i in range(len(employees)):
       employdict[employees[i]] = i+1
getDicts()
managers.append(getpass.getuser().lower())


class MyApp(tk.Tk):

    def __init__(self):
        tk.Tk.__init__(self)
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand = True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.frames = {}
        #creates frames of with given name and size in frames dict
        for F,geometry in zip((StartPage, FormPage, ApprovalPage, PurchasePage, LoginPage, StatusPage), 
                                ('410x300', '1620x690', '1400x600','1200x650','300x300','1550x550')):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames['%s'%page_name] = (frame, geometry)
            frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame("StartPage")

    def show_frame(self, page_name):
        #brings the chosen frame to the front
        frame, geometry = self.frames[page_name]
        self.update_idletasks()
        self.geometry(geometry)
        frame.tkraise()


class StatusPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        #creates labels and buttons 
        label1 = tk.Label(self, text = "Status Page",font= ('Arial 15 bold'))
        label1.grid(row = 0, column = 1, columnspan = 7)
        button1 = tk.Button(self, text = "Go to Start Page",
                         command=lambda: controller.show_frame("StartPage"))
        button1.grid(row = 6, column = 1)
        button2 = tk.Button(self, text = "Load Purchase Orders", command = lambda: self.loadPO())
        button2.grid(row = 6, column = 2)

        #creates stringvar and assigns to the combobox
        self.dropVar1 = tk.StringVar()
        self.ccb1 = Combobox(self, textvariable = self.dropVar1, values = sorted(employees), state = 'readonly')
        self.ccb1.grid(row = 7, column = 2)
        self.dropVar1.set("Select Employee")
        #if employee is not in purchasing, list only contains his/her name
        cur.execute("""select last_name, first_name from Employees2 where userid = '%s'"""%(getpass.getuser().lower()))
        for a in cur.fetchall():
            fname,lname = a[1],a[0]
        if getpass.getuser().lower() not in purchasing:
            self.ccb1['values'] = [lname+', '+fname]

        #creates treeview to populate with items 
        tv = Treeview(self, height = 20)
        tv['columns'] = ('ID','Status','Manager Approved', 'Require IT Signoff','IT Approved','FIsign','FIsigned','Date Raised', 'Last Updated', 'Additional Info')
        tv.heading("#0", text = "Item", anchor = 'w')
        tv.column("#0", anchor = "w", width = 50)
        tv.grid(row = 1, column = 1, columnspan = 5, rowspan = 5)
        tv.heading("#1", text = "ID")
        tv.column("#1", anchor = "center",minwidth = 50, width = 50)
        tv.heading("#2", text = "Status")
        tv.column("#2", anchor = "center",minwidth = 50, width = 150)
        tv.heading("#3", text = "Manager Approved")
        tv.column("#3", anchor = "center",minwidth = 50, width = 100)
        tv.heading("#4", text = "Require IT Signoff")
        tv.column("#4", anchor = "center",minwidth = 50, width = 120)
        tv.heading("#5", text = "IT Approved")
        tv.column("#5", anchor = "center",minwidth = 50, width = 100)
        tv.heading("#6", text = "Require Finance Sign")
        tv.column("#6", anchor = "center",minwidth = 50, width = 150)
        tv.heading("#7", text = "Finance Signed")
        tv.column("#7", anchor = "center",minwidth = 50, width = 120)
        tv.heading("#8", text = "Date Raised")
        tv.column("#8", anchor = "w",minwidth = 50, width = 200)
        tv.heading("#9", text = "Last Updated")
        tv.column("#9", anchor = "w",minwidth = 50, width = 200)
        tv.heading("#10", text = "Additional Info")
        tv.column("#10", anchor = "w",minwidth = 50, width = 300)
        self.treeview = tv
        #allows widgets to expand with window size 
        for i in range(11):
            self.grid_rowconfigure(i,weight = 1)
            self.grid_columnconfigure(i,weight = 1)    

    def loadPO(self):
        self.treeview.delete(*self.treeview.get_children())
        fname,lname = self.dropVar1.get()[self.dropVar1.get().index(',') + 1:].strip(),self.dropVar1.get()[:self.dropVar1.get().index(',')].strip()
        #grabs userid given name chosen in combobox
        cur.execute("""SELECT userid from Employees2 WHERE last_name = '%s' and first_name = '%s'""" % (lname,fname))
        for a in cur.fetchall():
            userID = a[0]
        #selects purchaseorders from db and adds to TV
        cur.execute("""SELECT id, status,manager_approved,require_IT_signoff,IT_approved,FIsign,FIsigned,date_raised,last_updated from PurchaseOrders WHERE raisedby_employees_id = '%s'""" % (userID))
        count = 1
        for a in cur.fetchall():
            a = list(a)
            self.treeview.insert('','end',text=count, values = a)
            count += 1


class LoginPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text = "User:")
        label2 = tk.Label(self, text = "Password:")
        e1 = tk.Entry(self)
        e1.insert(0,user)
        e2 = tk.Entry(self)
        label.pack()
        e1.pack()
        button1 = tk.Button(self, text = "Submit", command = lambda: self.checkUser(e1.get(), e2.get()))
        button1.pack()


    def checkUser(self, user, password):
        print(user, password)
        self.controller.show_frame("StartPage")


class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        self.controller = controller
        #creates labels for buttons on page 
        label1 = tk.Label(self, text = "Create a new Purchase Order",font= ('Arial 12 bold'))
        label2 = tk.Label(self, text = "Approve Purchase Orders",font= ('Arial 12 bold'))
        label3 = tk.Label(self, text = "Purchased approved Purchase Orders",font= ('Arial 12 bold'))
        label4 = tk.Label(self, text = "Check the status of your Purchase Orders",font= ('Arial 12 bold'))
        button1 = tk.Button(self, text = "Form Page",
                         command=lambda: controller.show_frame("FormPage"))
        button2 = tk.Button(self, text = "Approval Page",
                        command=lambda: controller.show_frame("ApprovalPage"))
        #if user not in manager list, disable button
        if getpass.getuser().lower() not in managers:
            button2.configure(state='disabled')
        button3 = tk.Button(self, text = "Purchase Page",
                        command=lambda: controller.show_frame("PurchasePage"))
        #if user not in puchasing list, disable button
        if getpass.getuser().lower() not in purchasing:
            button3.configure(state='disabled')
        button5 = tk.Button(self, text = "Log Out",
                        command =lambda: controller.show_frame("LoginPage"))
        button4 = tk.Button(self, text = "Status Page",
                        command =lambda: controller.show_frame("StatusPage"))
        button1.grid(row = 1, column = 0)
        label1.grid(row = 0, column = 0)
        button2.grid(row = 4, column = 0)
        label2.grid(row = 3, column = 0)
        button3.grid(row = 6, column = 0)
        label3.grid(row = 5, column = 0)
        button4.grid(row = 8, column = 0)
        label4.grid(row = 7, column = 0)
        button5.grid(row = 9, column = 0)
        for i in range(20):
            self.grid_rowconfigure(i,weight = 1)
            self.grid_columnconfigure(i,weight = 1)


class FormPage(tk.Frame):

    def saveFile(self):
        global entryParts, gl_inputs, gl_codes, lineItems
        if len(lineItems.keys()) < 1:
            messagebox.showinfo("Error", "Please add Line Items")
            return 
        individualinfo = []
        individualinfo.append(self.e2.get())
        individualinfo.append(self.e3.get())
        individualinfo.append(self.e4.get())
        individualinfo.append(self.e5.get())
        individualinfo.append(self.e8.get())
        individualinfo.append(self.e9.get())
        individualinfo.append(self.e10.get())
        individualinfo.append(self.e11.get())
        individualinfo.append(self.e16.get())
        individualinfo.append(self.e17.get())
        individualinfo.append(self.e18.get())
        individualinfo.append(self.e21.get())
        individualinfo.append(self.e22.get())
        individualinfo.append(format(self.dropVar.get()))
        individualinfo.append(format(self.dropVar2.get()))
        individualinfo.append(format(self.dropVar3.get()))
        individualinfo.append(format(self.dropVar4.get()))
        individualinfo.append(format(self.dropVar5.get()))
        individualinfo.append(format(self.dropVar6.get()))
        individualinfo.append(format(self.dropVar7.get()))

        supplierInfo = []
        supplierInfo.append(self.e8.get())
        supplierInfo.append(self.e9.get())
        supplierInfo.append(self.e10.get())
        supplierInfo.append(self.e11.get())
        supplierInfo.append(self.e16.get())
        supplierInfo.append(self.e17.get())
        supplierInfo.append(self.e18.get())

        #calculates total price of order
        temptot = 0
        for key in sorted(lineItems.keys()):
            temptot += float(lineItems[key][2]) * float(lineItems[key][3])

        global employees, supplierID
        #calculates data to insert into PurchaseOrders table depending if a supplier is chosen 
        if self.dropVar8.get() in supplierID.keys():
            cur2.execute("""select Address, City, Region, PostalCode, PhoneNumber, EmailAddress, Website
                            from Suppliers where SupplierName = ? """ , [str(self.dropVar8.get().rstrip())])
            for a in cur2.fetchall():
                sup_address = a[0].replace('\n','').replace('\r','')
                sup_city = a[1]
                sup_region = a[2]
                sup_postalcode = a[3]
                sup_phone = a[4]
                sup_email = a[5]
                sup_webref = a[6]
            suppID = supplierID[self.dropVar8.get()]
        else:
            suppID = self.dropVar8.get()
            sup_address = supplierInfo[0]
            sup_city = supplierInfo[1]
            sup_region = supplierInfo[2]
            sup_postalcode = supplierInfo[3]
            sup_phone = supplierInfo[5]
            sup_email = supplierInfo[4]
            sup_webref = supplierInfo[6]
        #determines the appropriate manager for the PO
        fname,lname = self.dropVar7.get()[self.dropVar7.get().index(',') + 1:].strip(),self.dropVar7.get()[:self.dropVar7.get().index(',')].strip()
        cur.execute("""Select userid,manager from Employees2 where last_name = '%s' and first_name = '%s'"""%(lname,fname))
        for a in cur.fetchall():
            employID = a[0]
            manager = a[1]
        groupnum,mID = None,None
        manAprov, ITaprov, SMTaprov, ITsign, SMTsign, FIsign, FIsigned = 'n','n','n','n','n','n','n'
        if getpass.getuser().lower() == itApprover:
            ITaprov = 'y'
        #if order is a CIR
        if self.dropVar3.get() == 'Yes':
            search = False
            #PO raiser is an approver
            if fname + ' ' + lname in list(itertools.chain(*approvalGroupsInventory.values())):
                for group in approvalGroupsInventory.keys():
                    if (fname + ' ' + lname) in approvalGroupsInventory[group]:
                        tgroupnum = group
                #checks if they can approve the po raised
                if temptot < float(purchaseLimitsInventory[tgroupnum]):
                    search = True
                    manAprov = 'y'
                    mID = employID
            #else finds the group of their manager
            mlname,mfname = manager[:manager.find(',')], manager[manager.find(',')+2:]
            for group in approvalGroupsInventory.keys():
                if (mfname + ' ' + mlname) in approvalGroupsInventory[group]:
                    groupnum = group
            while not search:
                if groupnum == None or temptot > float(purchaseLimitsInventory[groupnum]):
                    cur.execute("""Select manager from Employees2 where last_name = '%s' and first_name = '%s'""" % (mlname,mfname))
                    for a in cur.fetchall():
                        manager = a[0]
                        mlname,mfname = manager[:manager.find(',')], manager[manager.find(',')+2:]
                    for group in approvalGroupsInventory.keys():
                        if (mfname + ' ' + mlname) in approvalGroupsInventory[group]:
                            groupnum = group
                elif temptot <= float(purchaseLimitsInventory[groupnum]):
                    cur.execute("""Select userid from Employees2 where last_name = '%s' and first_name = '%s'""" % (mlname,mfname))
                    for a in cur.fetchall():
                        mID = a[0]
                    search = True
        #same logic for expense
        else:
            search = False
            if fname + ' ' + lname in list(itertools.chain(*approvalGroupsExpense.values())):
                for group in approvalGroupsExpense.keys():
                    if (fname + ' ' + lname) in approvalGroupsExpense[group]:
                        tgroupnum = group
                if temptot < float(purchaseLimitsExpense[tgroupnum]):
                    search = True
                    manAprov = 'y'
                    mID = employID
            mlname,mfname = manager[:manager.find(',')], manager[manager.find(',')+2:]
            for group in approvalGroupsExpense.keys():
                if (mfname + ' ' + mlname) in approvalGroupsExpense[group]:
                    groupnum = group
            while not search:
                if groupnum == None or temptot > float(purchaseLimitsExpense[groupnum]):
                    cur.execute("""Select manager from Employees2 where last_name = '%s' and first_name = '%s'""" % (mlname,mfname))
                    for a in cur.fetchall():
                        manager = a[0]
                        mlname,mfname = manager[:manager.find(',')], manager[manager.find(',')+2:]
                    for group in approvalGroupsExpense.keys():
                        if (mfname + ' ' + mlname) in approvalGroupsExpense[group]:
                            groupnum = group
                elif temptot <= float(purchaseLimitsExpense[groupnum]):
                    cur.execute("""Select userid from Employees2 where last_name = '%s' and first_name = '%s'""" % (mlname,mfname))
                    for a in cur.fetchall():
                        mID = a[0]
                    search = True
        print(mID)
        if self.dropVar3.get() == 'Yes': FIsign = 'y'
        poOutput = []
        dateRaised, dateRequired = str(individualinfo[0]), str(individualinfo[1])
        lastUpdate= str(strftime("%Y-%m-%d %H:%M:%S", localtime()))
        additionalInfo = self.e22.get()
        outsideEU, quoteRef = individualinfo[16], individualinfo[11]
        #makes sure commodity code is chosen if shipped from outside eu
        if outsideEU not in self.optionList:
            messagebox.showinfo("Error", "Select where items are shipped from")
            return
        purCur, stat = individualinfo[-3],"waiting"
        if purCur not in self.currencyList:
            messagebox.showinfo("Error", "Select appropriate currency")
            return
        temptot = float(temptot)
        CIR = individualinfo[-5]
        for key in lineItems.keys():
            if 'y' in lineItems[key][6]:
                ITsign = 'y'
            if 'y' in lineItems[key][7]:
                SMTsign = 'y' 
        if CIR == 'Yes':
            CIRnum = self.e25.get()
        else:
            CIRnum = None
        #inserts data into PO db 
        sql = """INSERT INTO PurchaseOrders(suppliers_id, date_raised, last_updated, raisedby_employees_id, required_date, require_IT_signoff,
             require_SMT_signoff, manager_approved, IT_approved, SMT_approved, status, total_price, currency, addition_info, CIR, supplier_email,
             supplier_phone, supplier_webref,supplier_address, supplier_city, supplier_region, supplier_postalcode, outsideEU, quoteRef,manager,FIsign,
             FIsigned, CIR_num) 
             VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""" 
        vals = (suppID,dateRaised,lastUpdate,employID,dateRequired,ITsign,SMTsign,
                    manAprov, ITaprov, SMTaprov, stat, temptot,purCur , additionalInfo, 
                    CIR, sup_email, sup_phone, sup_webref,sup_address, sup_city, sup_region, sup_postalcode,
                    outsideEU, quoteRef,mID,FIsign,FIsigned,CIRnum)

        cur.execute(sql, vals)
        #conn.commit()

        #gets the purchase order number that was just inserted
        purID = None
        cur.execute("""SELECT id FROM PurchaseOrders WHERE last_updated = ? """ ,[lastUpdate])
        for row in cur.fetchall():
            purID = row[0]
        # gets each line item info and inserts into db
        # lineItems maps item number to the items information
        gl_codes2 = {y.rstrip():x for x,y in gl_codes.items()}
        partNos = ''
        for key in sorted(lineItems.keys()):
            lineItem = key
            gl_ID = lineItems[key][4]
            partNo = lineItems[key][0]
            partNos += str(partNo) + ','
            msR = lineItems[key][5]
            itEquip = lineItems[key][-3]
            newEquip = lineItems[key][-2]
            descrip = lineItems[key][1]
            qty = lineItems[key][2]
            unitPrice = lineItems[key][3]
            currency = individualinfo[-3]
            commodityCode = lineItems[key][-1]
            if 'y' in itEquip and ITaprov == 'n':
                itapprov = 'n'
                cur.execute("""Update PurchaseOrders set require_IT_signoff = 'y' where id = %i"""%(purID))
                #conn.commit()
            else:
                itapprov = 'y'
            if individualinfo[16] == 'Yes':
                #raises error if incorrect commodity code chosen 
                if commodityCode not in commodityCodes:
                    messagebox.showinfo("Invalid Code", "Please get code validated before submitting")
                    cur.execute("""Delete from PurchaseOrders where id = %i"""%(purID))
                    conn.rollback()
                    return
                    raise ValueError('get commodity code validated')
                    
            commodityCode = commodityCode[:commodityCode.find('/')]
            sql = """ INSERT INTO LineItems(purchase_order_id, line_item, allowable_GLcodes_id,
                        partno, material_safety_ref, IT_equipment,new_equipment_register_checked,
                        description, qty, currency, unitprice, commoditycode,it_approved)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?) """
            vals = (purID, lineItem, gl_ID, partNo, msR,itEquip, newEquip, descrip, qty, currency,
                     unitPrice, commodityCode, itapprov)
            cur.execute(sql, vals)
            #conn.commit()
        #grabs the quote reference that was attached and uploads it 
        url = 'http://srv57mcsapps/upload_purchasereq'
        cur.execute("""Select selurl from Settings""")
        for a in cur.fetchall():
            url = a[0]
        global dirs
        #makes temporary folder to hold files
        for i in list(dirs.keys()):
            if not os.path.exists('/'.join(dirs[i])+'/Temp'):
                os.mkdir('/'.join(dirs[i])+'/Temp')
        #copies files to temp folder and renames with the purchase id
        for i in (list(self.quoterefa.keys())):
            for filename in os.listdir('/'.join(dirs[i])):
                fname = self.quoterefa[i]
                if filename == fname.split('/')[-1]:
                    newname = str(purID)+filename
                    shutil.copy('/'.join(dirs[i]) + '/' + filename, '/'.join(dirs[i]) + '/Temp' )
                    os.rename('/'.join(dirs[i]) + '/Temp/' + filename, '/'.join(dirs[i]) + '/Temp/' + newname)
                    self.files[i] = '/'.join(dirs[i]) + '/Temp/' + newname
        #uploads files
        for i in list(self.files.keys()):
            with open(self.files[i],'rb') as im:
                r = requests.post(url, files={'file':im})
        #removes temporary folders 
        for i in (list(self.quoterefa.keys())):
            if os.path.exists('/'.join(dirs[i])+'/Temp'):
                shutil.rmtree('/'.join(dirs[i])+'/Temp')
        filenames = {}
        #grabs just the file name 
        for i in (list(self.quoterefa.keys())):
            filenames[i] = (str(purID) + str(self.quoterefa[i].split('/')[-1]))
        #inserts appropriate info to quoteref db 
        for i in list(self.files.keys()):
            sql = """INSERT INTO QuoteRefs (purchase_order_id, url, comment) VALUES (?,?,?)"""
            vals = (purID, filenames[i],self.quoteComments[i])
            cur.execute(sql, vals)
            #conn.commit()
        messagebox.showinfo("Submission successful!", "Your PO ID is %i" %(purID))
        conn.commit()

        #grabs names and emails to send out messages
        cur.execute("""Select last_name, first_name, Email from employees2 where userid = '%s' """ % (mID))
        for a in cur.fetchall():
            mfname = a[1]
            mlname = a[0]
            memail = a[2]
        fname,lname = self.dropVar7.get()[self.dropVar7.get().index(',') + 1:].strip(),self.dropVar7.get()[:self.dropVar7.get().index(',')].strip()
        cur.execute("""Select email from employees2 where last_name = '%s' and first_name = '%s'""" %(lname, fname))
        for a in cur.fetchall():
            eemail = a[0]
        cur.execute("""Select email, first_name from employees2 where userid = '%s'"""%(itApprover))
        for a in cur.fetchall():
            itemail = a[0]
            itfname = a[1]
        cur.execute("""Select email, first_name from employees2 where userid = '%s'"""%(financeApprov[0]))
        for a in cur.fetchall():
            finemail = a[0]
            fifname = a[1]
        messagetext = """<p>Hi %s,</p>
            <p>A PurchaseOrder has been raised which requires 
            manager approval.</p><table border="1"><tbody><tr>
            <td><strong>POID:</strong></td>
            <td><font color="Red">%i</font></td>
            </tr><tr><td><strong>Raised by:</strong></td><td>%s</td>
            </tr><tr><td><strong>Required By:</strong></td><td>%s</td>
            </tr><tr><td><strong>Supplier Name:</strong></td><td>%s</td>
            </tr><tr><td><strong>Part No:</strong></td><td>%s</td>
            </tr><td><strong>Total Cost:</strong></td><td>%f</td>
            </tr><tr></tbody></table>
            <p>Please log into the Applicationsystem to approve this PurchaseOrder. If the order is not approved in 2 bussiness days it will continue up the approval chain.</p>
            Thanks,<br>%s<br>
            """ % (mfname, purID, self.dropVar7.get(), dateRequired, self.dropVar8.get(),partNos.rstrip(','), temptot, fname + ' ' + lname)
        messagetext2 = """<p>Hi %s,</p>
            <p>A PurchaseOrder has been raised which requires 
            IT approval.</p><table border="1"><tbody><tr>
            <td><strong>POID:</strong></td>
            <td><font color="Red">%i</font></td>
            </tr><tr><td><strong>Raised by:</strong></td><td>%s</td>
            </tr><tr><td><strong>Required By:</strong></td><td>%s</td>
            </tr><tr><td><strong>Supplier Name:</strong></td><td>%s</td>
            </tr><tr><td><strong>Part No:</strong></td><td>%s</td>
            </tr><td><strong>Total Cost:</strong></td><td>%f</td>
            </tr><tr></tbody></table>
            <p>Please log into the Applicationsystem to approve this PurchaseOrder. If the order is not approved in 2 bussiness days it will continue up the approval chain.</p>
            Thanks,<br>%s<br>
            """ % (itfname, purID, self.dropVar7.get(), dateRequired, self.dropVar8.get(),partNos.rstrip(','), temptot, fname + ' ' + lname)
        messagetext3 = """<p>Hi %s,</p>
            <p>A PurchaseOrder has been raised which requires 
            finance approval.</p><table border="1"><tbody><tr>
            <td><strong>POID:</strong></td>
            <td><font color="Red">%i</font></td>
            </tr><tr><td><strong>Raised by:</strong></td><td>%s</td>
            </tr><tr><td><strong>Required By:</strong></td><td>%s</td>
            </tr><tr><td><strong>Supplier Name:</strong></td><td>%s</td>
            </tr><tr><td><strong>Part No:</strong></td><td>%s</td>
            </tr><td><strong>Total Cost:</strong></td><td>%f</td>
            </tr><tr></tbody></table>
            <p>Please log into the Applicationsystem to approve this PurchaseOrder. If the order is not approved in 2 bussiness days it will continue up the approval chain.</p>
            Thanks,<br>%s<br>
            """ % (fifname, purID, self.dropVar7.get(), dateRequired, self.dropVar8.get(),partNos.rstrip(','), temptot, fname + ' ' + lname)

        sql = """insert into EmailMessages
        (storedatetime, application_name,processed,processed_datetime,HTML_message,subject,
          send_to_address,from_address,cc_address,bcc_address,reply_to_address,message_text)
        values
        (getDate(),'PurchaseReq','N', 0,'Y','Purchase Order %i requires your approval',
        '%s', 'Purchase Order Admin<noreply@spilasers.com>', '','','%s', '%s')
        """ % (purID,memail,eemail,messagetext)
        sql2 = """insert into EmailMessages
        (storedatetime, application_name,processed,processed_datetime,HTML_message,subject,
          send_to_address,from_address,cc_address,bcc_address,reply_to_address,message_text)
        values
        (getDate(),'PurchaseReq','N', 0,'Y','Purchase Order %i requires your approval',
        '%s', 'Purchase Order Admin<noreply@spilasers.com>', '','','%s', '%s')
        """ % (purID,itemail,eemail,messagetext2)
        sql3 = """insert into EmailMessages
        (storedatetime, application_name,processed,processed_datetime,HTML_message,subject,
          send_to_address,from_address,cc_address,bcc_address,reply_to_address,message_text)
        values
        (getDate(),'PurchaseReq','N', 0,'Y','Purchase Order %i requires your approval',
        '%s', 'Purchase Order Admin<noreply@spilasers.com>', '','','%s', '%s')
        """ % (purID,finemail,eemail,messagetext3)
        print(memail)
        #cur3.execute(sql)
        if ITsign == 'y':
            print(itemail)
            #cur3.execute(sql2)
        if FIsign == 'y':
            print(finemail)
            cur3.execute(sql3)
        conn3.commit()
        self.reset()

    def addItem(self):
        #adds lineItem information to itemInfo
        if self.dropVar4.get() not in self.optionList:
            messagebox.showinfo("Error", "Select where items are shipped from")
            return 
        global entryParts, gl_descriptions, gl_inputs, lineItems
        itemInfo = []
        #inserts individual line item information from Entryparts to iteminfo
        keys = sorted(entryParts.keys())
        for i in keys:
            itemInfo.append(entryParts[i].get())
        itemInfo.append(self.dropVar9.get())
        if len(itemInfo[-2]) > 1 or len(itemInfo[-3]) > 1:
            messagebox.showinfo("Error", "Use 'y' or 'n' for last two entries")
            return 
        if len(self.e4.get()) == 0 or len(self.e5.get()) == 0:
            messagebox.showinfo("Error", "Choose work site/Cost Centre")
            return
        if '' in itemInfo:
            messagebox.showinfo("Error", "Fill in all line entries")
            return
        if gl_inputs[0].get() not in self.dm11['values']:
            messagebox.showinfo("Error", "Choose appropriate GL Description")
            return
        if 'Yes' in self.dropVar4.get() and itemInfo[-1] not in commodityCodes:
            messagebox.showinfo("Error", "Choose correct code or get code validated")
            return
        if 'No' in self.dropVar4.get() and itemInfo[-1] in commodityCodes:
            messagebox.showinfo("Error", "No Commodity Code Needed")
            self.dropVar9.set("Select Code/Description")
            return
        else:
            #inserts the GL description into the list and then inserts list into the Treeview
            lineItems[self.itemcount] = (itemInfo)
            cur.execute("""SELECT glcode from GLCodes2 Where gldescrip = '%s'"""%(gl_inputs[0].get()))
            for a in cur.fetchall():
                glCode = a[0]
            glCode = str(self.e4.get()) + '.' + glCode + '.' + str(self.e5.get())
            itemInfo.insert(4, glCode)
            self.treeview.insert('', 'end', text = self.itemcount, values = (itemInfo))
            self.itemcount += 1
            for i in range(8):
                if i != 4:
                    entryParts[i].delete(0,'end')
            gl_inputs[0].set("Select")
            self.dropVar9.set("Select Code/Description")

    def removeItem(self):
        #removes selected line item from table 
        global lineItems
        item = self.treeview.item(self.treeview.selection())
        num, vals = item['text'], item['values']
        if len(vals)>0:
            del lineItems[num]
        self.treeview.delete(self.treeview.selection())

    def updateGLCodes(self):
        #updates glcode combobox with allowable options given cost centre
        options = []
        cur.execute("""SELECT "%s" from GLCostCentres""" % (self.e5.get()))
        for a in cur.fetchall():
            if a[0]:
                options.append(a[0])
        self.dm11['values'] = options

    def costCentreChosen(self,*args):
        if (self.dropVar.get() == "Yes"):
            #inserts values for name, worksite, cost centre given userid
            self.e5.delete(0,"end")
            self.e4.delete(0,"end")
            name = ''
            cur.execute("""SELECT last_name, first_name from Employees2 WHERE userid = '%s'""" %(getpass.getuser().lower()))
            for a in cur.fetchall():
                name += a[0] + ', ' + a[1]
            self.dropVar7.set(name)
            cur.execute("""SELECT cost_centre from Employees2 WHERE userid = '%s'""" %(getpass.getuser().lower()))
            for a in cur.fetchall():
                deptID = a[0]
            cur.execute("""SELECT dept_description,work_site from Departments2 WHERE dept_num = %i""" %(deptID))
            for a in cur.fetchall():
                dept = a[0]
                work_site = a[1]
            self.e5.insert(0,deptID)
            self.dropVar6.set(dept)
            if work_site == 1:
                self.e4.insert(0,"01")
                self.dropVar2.set("Southampton")
            else:
                self.e4.insert(0,"02")
                self.dropVar2.set("Rugby")
        else:
            #if not for user, clears the fields
            self.dropVar7.set("Select")
            self.e5.delete(0,"end")
            self.dropVar5.set("Select")
            self.dropVar6.set("Select")
            self.dropVar2.set("Select")
            self.e4.delete(0,"end")

    def supplierChosen(self, index, value, op):
        if self.dropVar8.get() in supplierID.keys():
            #deletes the supplier fields and inserts data from db 
            self.e8.delete(0,"end")
            self.e9.delete(0,"end")
            self.e10.delete(0,"end")
            self.e11.delete(0,"end")
            self.e16.delete(0,"end")
            self.e17.delete(0,"end")
            self.e18.delete(0,"end")
            cur2.execute("""SELECT Address, City, Region, PostalCode, EmailAddress, PhoneNumber, Website FROM Suppliers WHERE SupplierName = '%s'"""%(self.dropVar8.get()))
            for a in cur2.fetchall():
                address, city, region, postalcode, email, phone,web =  a[0],a[1],a[2],a[3],a[4],a[5],a[6]
            self.e8.insert(0,address)
            self.e9.insert(0,city)
            self.e10.insert(0,region)
            self.e11.insert(0,postalcode)
            self.e16.insert(0,email)
            self.e17.insert(0,phone)
            self.e18.insert(0,web)

    def employeeChosen(self,*args):
        self.e4.delete(0,"end")
        self.e5.delete(0,"end")
        #updates worksite, cost centre for chosen employee
        fname,lname = self.dropVar7.get()[self.dropVar7.get().index(',') + 1:].strip(),self.dropVar7.get()[:self.dropVar7.get().index(',')].strip()
        cur.execute("""Select cost_centre from Employees2 where last_name = '%s' and first_name = '%s'"""%(lname,fname))
        for a in cur.fetchall():
            deptID = a[0]
        cur.execute("""Select dept_description,work_site from Departments2 where dept_num = %i""" % (deptID))
        for a in cur.fetchall():
            deptDescrip = a[0]
            worksite = a[1]
        if worksite == 1:
            self.e4.insert(0,"01")
            self.dropVar2.set("Southampton")
        else:
            self.e4.insert(0,"02")
            self.dropVar2.set("Rugby")
        self.e5.insert(0,deptID)
        self.dropVar6.set(deptDescrip)

    def workSiteChosen(self,*args):
        #fills in site# given work site chosen
        self.e4.delete(0,"end")
        if self.dropVar2.get() == 'Southampton':
            self.e4.insert(0,"01")
        else:
            self.e4.insert(0,"02")

    def deptChosen(self,*args):
        #fills in deptnum baseed on department 
        dept = self.dm6.get()
        cur.execute("""select dept_num from Departments2 where dept_description = '%s'""" % (dept))
        for a in cur.fetchall():
            dept_num = a[0]
        self.e5.delete(0,"end")
        self.e5.insert(0,dept_num)

    def reset(self):
        #resets values so another PO can be made
        global dirs
        self.itemcount = 1
        self.dropVar2.set("Select")
        self.dropVar7.set("Select")
        self.dropVar.set("Select")
        self.e2.delete(0,"end")
        self.e3.delete(0,"end")
        curdate = str(datetime.datetime.now())
        curdate = curdate[:curdate.index('.')]
        self.e2.insert(0,curdate)
        wantdate = str(datetime.datetime.now() + datetime.timedelta(days=5))
        wantdate = wantdate[:wantdate.index('.')]
        self.e3.insert(0,wantdate)
        self.e2.grid(row = 5, column = 2)
        self.e3.grid(row = 6, column = 2)
        self.e24.delete(0,"end")
        self.dropVar6.set("Select")
        self.dropVar3.set("No")
        self.dropVar8.set("Manually enter if not available")
        self.e8.delete(0,"end")
        self.e9.delete(0,"end")
        self.e10.delete(0,"end")
        self.e11.delete(0,"end")
        self.e16.delete(0,"end")
        self.e17.delete(0,"end")
        self.e18.delete(0,"end")
        self.dropVar4.set("Select")
        self.dropVar5.set("Select")
        self.e21.delete(0,"end")
        self.dropVar9.set("Select Code/Description")
        self.e22.delete(0,"end")
        self.e25.delete(0,"end")
        self.sVar.set("Select")
        self.treeview.delete(*self.treeview.get_children())
        self.treeview2.delete(*self.treeview2.get_children())            
        self.files= {}
        self.quoteComments = {}
        dirs = {}
        self.quoterefa = {}
        self.imCount = 1
        self.addedFiles.configure(state = 'normal')
        self.addedFiles.delete('1.0', END)
        self.addedFiles.configure(state = 'disabled')
        self.dm9.configure(state = 'disabled')
        self.dm11.configure(state = 'disabled')
    
    def ShipChosen(self, *args):
        #disables commodity code if not shipped from outside the EU
        if self.dropVar4.get() == 'No':
            self.dm9.configure(state = 'disabled')
        else:
            self.dm9.configure(state = 'readonly')

    def uGLCode(self, *args):
        #only allows glcodes to be chosen when cost centre has been chosen 
        if self.dropVar6.get() != 'Select':
            self.dm11.configure(state = 'readonly') 
        else:
            self.dm11.configure(state = 'disabled')

    def validate(self, action, index, value_if_allowed, prior_value, text, validation_type, trigger_type, widget_name):
        #only allows numbers/special chars
        if not text: return True
        try:
            for v in text:
                if v in '1234567890-: ':
                    return True
                else: 
                    ValueError
                    return False
        except ValueError:
            return False

    def validate2(self, action, index, value_if_allowed, prior_value, text, validation_type, trigger_type, widget_name):
        #only allows numbers
        if not text: return True
        try:
            for v in text:
                if v in '1234567890':
                    return True
                else: 
                    ValueError
                    return False
        except ValueError:
            return False

    def validate3(self, action, index, value_if_allowed, prior_value, text, validation_type, trigger_type, widget_name):
        #only allows for y/n
        if not text: return True
        try:
            for v in text:
                if v in 'yn':
                    return True
                else: 
                    ValueError
                    return False
        except ValueError:
            return False

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        self.controller = controller
        self.itemcount = 1

        label1 = tk.Label(self, text="Individual Information",font= ('Arial 15 bold'))
        label1.grid(row = 1, column = 2, columnspan = 3)
        label2 = tk.Label(self, text="Work Site")
        label2.grid(row = 4, column = 1)

        siteList = ["Southampton", "Rugby"]
        self.dropVar2 = tk.StringVar()
        self.dropVar2.set("Select")
        self.dm2 = Combobox(self, textvariable=self.dropVar2, values = siteList, state = 'readonly')
        self.dm2.grid(row = 4, column = 2)
        self.dm2.bind("<<ComboboxSelected>>",self.workSiteChosen)

        label3 = tk.Label(self, text = "For Employee:")
        label3.grid(row = 3, column = 1)
        self.dropVar7 = tk.StringVar()
        self.dropVar7.set("Select")
        self.dm7 = Combobox(self, textvariable = self.dropVar7, values = employees, state = 'readonly')

        self.dm7.grid(row = 3, column = 2)
        self.dm7.bind("<<ComboboxSelected>>",self.employeeChosen)

        label4 = tk.Label(self, text = "For Your Cost Centre? ")
        label4.grid(row = 2, column = 1)
        self.optionList = ["Yes","No"]
        self.dropVar = tk.StringVar()
        self.dropVar.set("Select")
        self.dropVar.trace('w',self.costCentreChosen)
        self.dm1 = Combobox(self,textvariable =  self.dropVar, values = self.optionList, state = 'readonly')
        self.dm1.grid(row = 2, column = 2)

        label5 = tk.Label(self, text = "Date Raised \n (YYYY-MM-DD HH:MM:SS)")
        label6 = tk.Label(self, text = "Date Required \n (YYYY-MM-DD HH:MM:SS)")
        label5.grid(row = 5, column = 1)
        label6.grid(row = 7, column = 1)
        vcmd = (self.register(self.validate),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        vcmd2 = (self.register(self.validate2),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        vcmd3 = (self.register(self.validate3),
                '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        self.e2 = tk.Entry(self, validate = 'key',validatecommand = vcmd)
        self.e3 = tk.Entry(self, validate = 'key',validatecommand = vcmd)
        self.e2.configure(relief = SUNKEN, borderwidth = 3)
        self.e3.configure(relief = SUNKEN, borderwidth = 3)
        curdate = str(datetime.datetime.now())
        curdate = curdate[:curdate.index('.')]
        self.e2.insert(0,curdate)
        wantdate = str(datetime.datetime.now() + datetime.timedelta(days=5))
        wantdate = wantdate[:wantdate.index('.')]
        self.e3.insert(0,wantdate)
        self.e2.grid(row = 5, column = 2)
        self.e3.grid(row = 7, column = 2)

        label7 = tk.Label(self, text = "For Site #")
        label7.grid(row = 2, column = 3)
        self.e4 = tk.Entry(self, width = 7, validate = 'key',validatecommand = vcmd2)
        self.e4.grid(row = 3, column = 3)
        self.e4.configure(relief = SUNKEN, borderwidth = 3)
        label8 = tk.Label(self, text = "For Cost Centre #")
        label8.grid(row = 2, column = 4)
        self.e5 = tk.Entry(self, width = 7, validate = 'key',validatecommand = vcmd2)
        self.e5.grid(row = 3, column = 4)
        self.e5.configure(relief = SUNKEN, borderwidth = 3)

        label33 = tk.Label(self, text = "Attach Quote Reference",font= ('Arial 15 bold'))
        label33.grid(row = 1, column = 5, columnspan = 2)
        button6 = tk.Button(self, text = "Attach", command = self.attachFile).grid(row = 3, column = 5)
        self.quoterefa = {}
        label34 = tk.Label(self, text = "Quote Reference Comment \n (Before Attaching)")
        label34.grid(row = 4, column = 6)
        self.e24 = tk.Entry(self)
        self.e24.grid(row = 5, column = 6)
        self.e24.configure(relief = SUNKEN, borderwidth = 3)
        
        label9 = tk.Label(self, text = "For Department")
        label9.grid(row = 4, column = 4)
        self.dropVar6 = tk.StringVar()
        self.dropVar6.set("Select")
        self.dm6 = Combobox(self, textvariable = self.dropVar6, values = sorted(list(departmentDict.values())), state = 'readonly')
        self.dm6.grid(row = 5, column = 4)
        self.dm6.bind("<<ComboboxSelected>>",self.deptChosen)
        self.dropVar6.trace('w',self.uGLCode)

        label10 = tk.Label(self, text = "Capital Investment Request?")
        label10.grid(row = 4, column = 3)
        self.dropVar3 = tk.StringVar()
        self.dropVar3.set("No")
        self.dm3 = Combobox(self, textvariable = self.dropVar3, values = self.optionList, state = 'readonly')
        self.dm3.grid(row = 5, column = 3)

        label11 = tk.Label(self, text = "Supplier Information",font= ('Arial 15 bold'))
        label11.grid(row = 1, column = 7, columnspan = 2) 
        label12 = tk.Label(self, text = "Supplier ID")
        label12.grid(row = 2, column = 7)
        self.dropVar8 = tk.StringVar()
        self.dropVar8.set("Manually enter if not available")
        self.dropVar8.trace('w',self.supplierChosen)
        a = sorted(supplierID.keys())
        self.dm8 = Combobox(self, textvariable = self.dropVar8, values = a, width = 26)
        self.dm8.grid(row = 2, column = 8)
        label14 = tk.Label(self, text = "Address:")
        label14.grid(row = 3, column = 7)
        self.e8 = tk.Entry(self)
        self.e8.grid(row = 3, column = 8)
        self.e9 = tk.Entry(self)
        self.e9.grid(row = 4, column = 8)
        self.e10 = tk.Entry(self)
        self.e10.grid(row = 5, column = 8)
        self.e11 = tk.Entry(self)
        self.e11.grid(row = 6, column = 8)
        self.e8.configure(relief = SUNKEN, borderwidth = 3)
        self.e9.configure(relief = SUNKEN, borderwidth = 3)
        self.e10.configure(relief = SUNKEN, borderwidth = 3)
        self.e11.configure(relief = SUNKEN, borderwidth = 3)

        label16 = tk.Label(self, text = "Email").grid(row = 7, column = 7)
        self.e16 = tk.Entry(self)
        self.e16.grid(row = 7, column = 8)
        self.e16.configure(relief = SUNKEN, borderwidth = 3)

        label17 = tk.Label(self, text = "Phone #").grid(row = 8, column = 7)
        self.e17 = tk.Entry(self)
        self.e17.grid(row = 8, column = 8)
        self.e17.configure(relief = SUNKEN, borderwidth = 3)

        label18 = tk.Label(self, text = "Web Reference").grid(row = 9, column = 7)
        self.e18 = tk.Entry(self)
        self.e18.grid(row = 9, column = 8)
        self.e18.configure(relief = SUNKEN, borderwidth = 3)

        label19 = tk.Label(self, text = "Shipped from outside EU?").grid(row = 10, column = 7)
        self.dropVar4 = tk.StringVar()
        self.dropVar4.set("Select")
        self.dm4 = Combobox(self, textvariable = self.dropVar4, values = self.optionList, state = 'readonly')
        self.dm4.grid(row = 10, column = 8)
        self.dropVar4.trace('w',self.ShipChosen)

        label20 = tk.Label(self, text = "Supplier Currency").grid(row = 11, column = 7)
        self.currencyList = ["British Pound", "US Dollar", "Euro"]
        self.dropVar5 = tk.StringVar()
        self.dropVar5.set("Select")
        self.dm5 =Combobox(self, textvariable = self.dropVar5, values = self.currencyList, state = 'readonly')
        self.dm5.grid(row = 11, column = 8)

        label21 = tk.Label(self, text = "Quote Reference:").grid(row = 2, column = 6)
        self.e21 = tk.Entry(self)
        self.e21.grid(row = 3, column = 6)
        self.e21.configure(relief = SUNKEN, borderwidth = 3)
        label36 = tk.Label(self, text = "Line Information",font= ('Arial 15 bold'))
        label36.grid(row = 8, column = 1, columnspan = 6)
        label22 = tk.Label(self, text = "Supplier Part No.").grid(row = 15, column = 1)
        label23 = tk.Label(self, text = "Supplier's Description").grid(row = 15, column = 2)
        label24 = tk.Label(self, text = "Qty.").grid(row = 15, column = 3)
        label25 = tk.Label(self, text = "Unit Price").grid(row = 15, column = 4)
        label26 = tk.Label(self, text = "GL Description").grid(row = 15, column = 5)
        label27 = tk.Label(self, text = "Agile MSDS Reference No.").grid(row = 15, column = 6)
        label28 = tk.Label(self, text = "IT Equipment?(y/n)").grid(row = 15, column = 7)
        label29 = tk.Label(self, text = "New Equipment?(y/n)").grid(row = 15, column = 8)
        label31 = tk.Label(self, text = "Commodity Code \n (If shipped from outside EU)").grid(row = 19, column = 5)
        imgp = resource_path("spi.ico")
        img = ImageTk.PhotoImage(Image.open(imgp))
        label36 = Label(self, image = img)
        label36.image = img
        label36.grid(row = 20, column = 0)
        global commodityCodes
        self.dropVar9 = tk.StringVar()
        self.dropVar9.set("Select Code/Description")
        self.dm9 = Combobox(self, textvariable = self.dropVar9, values = commodityCodes,  width = 25, state = 'disabled')
        self.dm9.grid(row = 20, column = 5)

        label30 = tk.Label(self, text = "Additional Info").grid(row = 6, column = 4)
        self.e22 = tk.Entry(self, width = 25)
        self.e22.grid(row = 7, column = 4)
        self.e22.configure(relief = SUNKEN, borderwidth = 3)
        label35 = tk.Label(self, text = "CIR # \n (If CIR attach CIR Form)")
        label35.grid(row = 6, column = 3)
        self.e25 = tk.Entry(self)
        self.e25.grid(row = 7,column = 3)
        self.e25.configure(relief = SUNKEN, borderwidth = 3)
        #creates text entries for each line item to add to treeview
        global entryParts, gl_descriptions, gl_inputs
        for row in range(1):
            for col in range(8):
                if col != 4:
                    if col == 2 or col == 3:
                        ent = tk.Entry(self, width = 15, validate = 'key',validatecommand = vcmd2)
                        ent.grid(row = 17 + row, column = col + 1)
                        ent.configure(relief = SUNKEN, borderwidth = 3)
                        entryParts[(col)] = ent
                    elif col == 6 or col == 7:
                        ent = tk.Entry(self, width = 15, validate = 'key',validatecommand = vcmd3)
                        ent.grid(row = 17 + row, column = col + 1)
                        ent.configure(relief = SUNKEN, borderwidth = 3)
                        entryParts[(col)] = ent
                    else:
                        ent = tk.Entry(self)
                        ent.grid(row = 17 + row, column = col + 1)
                        ent.configure(relief = SUNKEN, borderwidth = 3)
                        entryParts[(col)] = ent
                else:
                    self.sVar = tk.StringVar()
                    self.sVar.set("Select")
                    self.dm11 = Combobox(self, textvariable = self.sVar, values = sorted(gl_descriptions), postcommand = self.updateGLCodes, state = 'disabled')
                    self.dm11.grid(row = 17+row, column = 5)
                    gl_inputs[row] = self.sVar

        tv = Treeview(self)
        #sets up treeview for line items 
        tv['columns'] = ('Part No','Description', 'Qty', 'Unit Price', 'GL Description', 'MSDS No.', 'IT Equip', 'New Equip', 'Commodity Code')
        tv.heading("#0", text = "Item", anchor = 'w')
        tv.column("#0", anchor = "w", width = 40)
        tv.heading("#1", text = "Part No")
        tv.column("#1", anchor = "center",minwidth = 50, width = 100)
        tv.heading("#2", text = "Description")
        tv.column("#2", anchor = "center",minwidth = 50, width = 200)
        tv.heading("#3", text = "Qty")
        tv.column("#3", anchor = "center",minwidth = 50, width = 50)
        tv.heading("#4", text = "Unit Price")
        tv.column("#4", anchor = "center",minwidth = 50, width = 75)
        tv.heading("#5", text = "GL Description")
        tv.column("#5", anchor = "center",minwidth = 50, width = 200)
        tv.heading("#6", text = "MSDS No.")
        tv.column("#6", anchor = "center",minwidth = 50, width = 100)
        tv.heading("#7", text = "IT Equip.")
        tv.column("#7", anchor = "center",minwidth = 50, width = 75)
        tv.heading("#8", text = "New Equip.")
        tv.column("#8", anchor = "center",minwidth = 50, width = 80)
        tv.heading("#9", text = "Commodity Code")
        tv.column("#9", anchor = "center",minwidth = 50, width = 200)
        tv.grid(row = 9, column = 0, columnspan = 7, rowspan = 5)
        self.treeview = tv

        self.imCount = 1
        tv2 = Treeview(self)
        tv2['columns'] = ('File')
        tv2.heading("#0", text = "Item", anchor = 'w')
        tv2.column("#0", anchor = "w", width = 40)
        tv2.heading("#1", text = "File")
        tv2.column("#1", anchor = "center",minwidth = 50, width = 150)
        tv2.grid(row = 4, column = 5, rowspan = 4)
        self.treeview2 = tv2
        self.treeview2.bind("<Double-Button-1>", self.removeFile)
        button1 = tk.Button(self, text = "Go to Start Page",
                         command=lambda: controller.show_frame("StartPage"))
        button1.grid(row = 19, column = 1)
        button2 = tk.Button(self, text = "Save", command = self.saveFile).grid(row = 19, column = 2)
        button3 = tk.Button(self, text = "Add Item", command = self.addItem).grid(row = 19, column = 3)
        button4 = tk.Button(self, text = "Remove Item", command = self.removeItem).grid(row = 19, column = 4)
        self.files= {}
        self.quoteComments = {}
        self.addedFiles = tk.Text(self, width = 20, height = 8)
        #self.addedFiles.grid(row = 4, column = 5, rowspan = 4)
        self.addedFiles.configure(state = 'disabled',relief = SUNKEN, borderwidth = 3 )
        for i in range(22):
            self.grid_rowconfigure(i,weight = 1)
            self.grid_columnconfigure(i,weight = 1)

    def attachFile(self):
        global dirs
        #adds path to quoterefa, split path to dirs, comments to quoteComments 
        imagePath = tk.filedialog.askopenfilename()
        if len(str(imagePath)) > 1:
            self.quoterefa[self.imCount]= imagePath
            dirs[self.imCount] = imagePath.split('/')[:-1]
            self.quoteComments[self.imCount]= self.e24.get()
            self.treeview2.insert('', 'end', text = self.imCount, values = (imagePath.split('/')[-1]))
            self.imCount += 1
        #popup with file name back to user 
        messagebox.showinfo("File Upload", "Your file was: %s"%imagePath.split('/')[-1])
        print(self.quoterefa)
        print(dirs)
        self.e24.delete(0,"end")
        print(self.quoteComments)
    
    def removeFile(self, *args):
        global dirs
        item = self.treeview2.item(self.treeview2.selection())
        num, vals = item['text'], item['values']
        result = messagebox.askyesno(message = "Are you sure you want to remove this file")
        if result:
            self.quoterefa.pop(num)
            dirs.pop(num)
            self.quoteComments.pop(num)
            self.treeview2.delete(self.treeview2.selection())

        print(self.quoterefa)
        print(dirs)
        print(self.quoteComments)

class ApprovalPage(tk.Frame):

    def loadPO(self):
        print('loading Purchase Orders')
        #grabs the columns from Purchase Orders
        user = str(getpass.getuser().lower())
        count = 1
        #clears the treeview screen
        self.treeview.delete(*self.treeview.get_children())
        #if the PO needs approval, shows in the treeview
        if getpass.getuser().lower() == itApprover:
            cur.execute("""SELECT id, currency, total_price, manager_approved,require_IT_signoff,IT_approved,manager,FIsign, FIsigned FROM PurchaseOrders WHERE manager = '%s' or require_IT_signoff = 'y'"""%(getpass.getuser().lower()))
            for a in cur.fetchall():
                #if requires manager approval and manager is user 
                if a[3] != 'y' and a[6] == getpass.getuser().lower():
                    t = a[:6] + a[7:]
                    self.treeview.insert('', 'end', text= count, values = list(t))
                    count += 1
                #if requires it signoff and has not yet been signed 
                elif a[4] == 'y' and a[5] == 'n':
                    t = a[:6] + a[7:]
                    self.treeview.insert('', 'end', text= count, values = list(t))
                    count += 1
        elif getpass.getuser().lower() in financeApprov:
            cur.execute("""SELECT id, currency, total_price, manager_approved,require_IT_signoff,IT_approved,manager,FIsign, FIsigned FROM PurchaseOrders WHERE manager = '%s' or FIsign = 'y'"""%(getpass.getuser().lower()))
            for a in cur.fetchall():
                #if the PO requires finance signing and has not yet been signed
                if a[-2] == 'y' and a[-1] == 'n':
                    t = a[:6] + a[7:]
                    self.treeview.insert('', 'end', text= count, values = list(t))
                    count += 1
                #if has not been approved and user is the manager 
                elif a[3] == 'n' and a[6] == getpass.getuser().lower():
                    t = a[:6] + a[7:]
                    self.treeview.insert('', 'end', text= count, values = list(t))
                    count += 1
        else:
            sql ="""SELECT id, currency, total_price, manager_approved,require_IT_signoff,IT_approved,manager,FIsign, FIsigned FROM PurchaseOrders WHERE manager = '%s'""" % (getpass.getuser().lower())
            cur.execute(sql)
            for a in cur.fetchall():
                if a[3] != 'y' and a[6] == getpass.getuser().lower():
                    self.treeview.insert('', 'end', text= count, values = list(a))
                    self.treeValues[count] = a
                    count += 1
        count = 1
        print('done')

    def checkPO(self, purchaseOID):
        cur.execute("""SELECT manager_approved, it_approved FROM LineItems WHERE purchase_order_id = %i""" % (purchaseOID))
        approved = True
        approvedIT = True
        #checks to make sure each item has been approved 
        for a in cur.fetchall():
            if 'n' in a[0]:
                approved = False
            if 'n' in a[1]:
                approvedIT = False
        #if the entire order has been approved, update the appropriate columns in the purchaseorder db 
        if approved:
            sql = """ UPDATE PurchaseOrders SET manager_approved = 'y',status = 'approved' WHERE id = %i""" % (purchaseOID)
            cur.execute(sql)
            sql = """ UPDATE PurchaseOrders SET last_updated = getDate() WHERE id = %i""" % (purchaseOID)
            cur.execute(sql)
            conn.commit()
            #gets info to send notification to employee
            cur.execute("""select raisedby_employees_id, required_date, suppliers_id, total_price, date_raised from purchaseOrders where id = %i"""%(purchaseOID))
            for a in cur.fetchall():
                employId = a[0]
                reqDate = a[1]
                supId = a[2]
                tot = a[3]
                rDate = a[4]
            cur.execute("""Select last_name, first_name, Email from Employees2 where userid = '%s'"""%(a[0]))
            for a in cur.fetchall():
                fname = a[1]
                lname = a[0]
                eemail = a[2]

            messagetext = """<p>Hi %s,</p><p>A Purchase Order you raised has been approved by your manager.</p>
                <table border="1"><tbody><tr>
                <td><strong>POID:</strong></td>
                <td><font color="Red">%i</font></td>
                </tr><tr><td><strong>Raised On:</strong></td><td>%s</td>
                </tr><tr><td><strong>Required By:</strong></td><td>%s</td>
                </tr><tr><td><strong>Supplier Name:</strong></td><td>%s</td>
                </tr><td><strong>Total Cost:</strong></td><td>%f</td>
                </tr><tr></tbody></table></tbody></table>
                """ % (fname, purchaseOID,rDate, reqDate, supId, tot)
            sql = """insert into EmailMessages
            (storedatetime,application_name,processed,processed_datetime,HTML_message,subject,
              send_to_address,from_address,cc_address,bcc_address,reply_to_address,message_text)
            values
            (getDate(),'PurchaseReq','N',0,'Y','Purchase Order %i has been approved',
            '%s', 'Purchase Order Admin<noreply@spilasers.com>', '','','%s', '%s')
            """ %  (purchaseOID,eemail,'noreply@spilasers.com', messagetext)
            print(eemail)
            #cur3.execute(sql)
            #conn3.commit()
            self.treeview2.delete(*self.treeview2.get_children())
            self.loadPO()
        if approvedIT:
            cur.execute("""update PurchaseOrders set IT_approved = 'y', last_updated = getDate() where id = %i""" % (purchaseOID))
            conn.commit()
            self.treeview2.delete(*self.treeview2.get_children())
            self.loadPO()
        if getpass.getuser().lower() == itApprover:
            if approved and approvedIT:
                self.treeview2.delete(*self.treeview2.get_children())
                self.loadPO()
        elif approved:
            self.treeview2.delete(*self.treeview2.get_children())

    def showLineItems(self, event):
        #somehow show only line items for appropriate Dept. Manager 
        print("loading Line Items")
        #clears the treeview and adds line items from order to bottom TV
        self.treeview2.delete(*self.treeview2.get_children())
        item = self.treeview.item(self.treeview.selection())
        num, vals = item['text'], item['values']
        if len(vals) > 0:
            poID = vals[0]
            cur.execute("""SELECT partno, description, qty, unitprice, manager_approved, it_approved FROM LineItems WHERE purchase_order_id = %i""" % (poID))
            count = 1
            for a in cur.fetchall():
                self.treeview2.insert('', 'end', text =count, values = list(a) )
                count += 1
            count = 1
        print("done")

    def showLineItems2(self):
        print("loading Line Items")
        #same thing but can be called wihout event 
        self.treeview2.delete(*self.treeview2.get_children())
        item = self.treeview.item(self.treeview.selection())
        num, vals = item['text'], item['values']
        if len(vals) > 0:
            poID = vals[0]
            cur.execute("""SELECT partno, description, qty, unitprice, manager_approved,it_approved FROM LineItems WHERE purchase_order_id = %i""" % (poID))
            count = 1
            for a in cur.fetchall():
                self.treeview2.insert('', 'end', text =count, values = list(a))
                count += 1
            count = 1
        print("done")

    def approveItem(self):
        item = self.treeview2.item(self.treeview2.selection())
        num, vals = item['text'], item['values']
        pitem = self.treeview.item(self.treeview.selection())
        pnum, pvals = pitem['text'], pitem['values']
        if len(vals) > 0:
            result = messagebox.askyesno(message = 'Are you sure you want to approve?')
            if result:
                #updates the approval and last update if the item has not yet been approved 
                if 'n' in pvals[3]:
                    cur.execute("""Select manager from PurchaseOrders Where id = %i"""%(pvals[0]))
                    for a in cur.fetchall():
                        man = a[0]
                    if getpass.getuser().lower() == man:
                        sql = """UPDATE LineItems SET manager_approved = 'y' WHERE purchase_order_id = %i AND line_item = %i""" % (pvals[0], num)
                        cur.execute(sql)
                        sql = """ UPDATE PurchaseOrders SET last_updated = getDate() WHERE id = %i""" % (pvals[0])
                        cur.execute(sql)
                        conn.commit()
        self.showLineItems2()
        self.checkPO(pvals[0])

    def ITApprove(self):
        item = self.treeview.item(self.treeview.selection())
        num, vals = item['text'], item['values']
        item2 = self.treeview2.item(self.treeview2.selection())
        pnum, pvals = item2['text'], item2['values']
        if len(vals)>0:
            result = messagebox.askyesno(message = 'Are you sure you want to approve?')
            if result:
                if 'y' in vals[4] and 'n' in vals[5]:
                    cur.execute("""Update LineItems set it_approved = 'y' where purchase_order_id = %i and line_item = %i""" % (vals[0],pnum))
                    conn.commit()
        self.showLineItems2()
        self.checkPO(vals[0])

    def Fapprove(self):
        item = self.treeview.item(self.treeview.selection())
        num, vals = item['text'], item['values']
        item2 = self.treeview2.item(self.treeview2.selection())
        pnum, pvals = item2['text'], item2['values']
        if len(vals) > 0:
            result = messagebox.askyesno(message = 'Are you sure you want to approve?')
            if result:
                if 'y' in vals[-2] and 'n' in vals[-1]:
                    cur.execute("""Select CIR_num from PurchaseOrders where id = %i""" % vals[0])
                    cirnums = None
                    for a in cur.fetchall():
                        cirnums = a[0]
                    if cirnums == ' ' and vals[6] == 'y':
                        Cirnumss = None
                        while Cirnumss == None:
                            Cirnumss = askstring('Add CIR Number', 'Input CIR # to add')
                        cur.execute("""Update PurchaseOrders set FIsigned = 'y',CIR_num = %s, last_updated = getDate() where id = %i"""%(Cirnumss, vals[0]))
                        conn.commit()
                        messagebox.showinfo("CIR Code Added", "CIR Number successfully added")
                    cur.execute("""Update PurchaseOrders set FIsigned = 'y', last_updated = getDate() where id = %i"""%(vals[0]))
                    conn.commit()
        self.loadPO()
        self.showLineItems2()

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        self.controller = controller
        self.CreateUI()
        label = tk.Label(self, text = "Approval Page",font= ('Arial 15 bold'))
        label.grid(row = 0, column = 1, columnspan = 9)
        button1 = tk.Button(self, text = "Go to Start Page",
                         command=lambda: controller.show_frame("StartPage"))
        button1.grid(row = 3, column = 1)
        button2 = tk.Button(self, text = "Load Purchase Orders", command = lambda:self.loadPO())
        button2.grid(row = 3, column = 2)
        button3 = tk.Button(self, text = "Approve Item", command = self.approveItem)
        button3.grid(row = 3, column = 3)
        button4 = tk.Button(self, text = "IT Approve", command = self.ITApprove)
        button4.grid(row = 3, column = 4)
        button5 = tk.Button(self, text = "Finance Approve", command = self.Fapprove)
        button5.grid(row = 3, column = 5)
        # button6 = tk.Button(self, text = "Add CIR Code", command = self.addCIR)
        #button6.grid(row = 5, column = 5)

        if getpass.getuser().lower() != itApprover:
            button4.configure(state = 'disabled')
        if getpass.getuser().lower() not in financeApprov:
            button5.configure(state = 'disabled')
        for i in range(11):
            self.grid_rowconfigure(i,weight = 1)
            self.grid_columnconfigure(i,weight = 1)

    def CreateUI(self):
        tv = Treeview(self)
        #sets up the top treeview
        tv['columns'] = ('ID','Currency', 'Total Price', 'Manager Approved?', 'Requires IT Approval', 'IT Approved','Requires Finance Sign','Finance Signed')
        tv.heading("#0", text = "Item", anchor = 'w')
        tv.column("#0", anchor = "w", width = 50)
        tv.heading("#1", text = "ID")
        tv.column("#1", anchor = "center",minwidth = 50, width = 100)
        tv.heading("#2", text = "Currency")
        tv.column("#2", anchor = "center",minwidth = 50, width = 200)
        tv.heading("#3", text = "Total Price")
        tv.column("#3", anchor = "center",minwidth = 50, width = 200)
        tv.heading("#4", text = "Manager Approved")
        tv.column("#4", anchor = "center",minwidth = 50, width = 150)
        tv.column("#5", anchor = "center",minwidth = 50, width = 150)
        tv.heading("#5", text = "Req. IT Approval")
        tv.column("#6", anchor = "center",minwidth = 50, width = 150)
        tv.heading("#6", text = "IT Approved")
        tv.column("#7", anchor = "center",minwidth = 50, width = 150)
        tv.heading("#7", text = "Req. Finance Sign")
        tv.column("#8", anchor = "center",minwidth = 50, width = 165)
        tv.heading("#8", text = "Finance Signed")
        tv.grid(row = 1, column = 1, columnspan =9)
        tv2 = Treeview(self)
        #sets up the bottom treeview 
        tv2['columns'] = ('Part No.', 'Description', 'Quantity', 'Unit Price', 'Manager Approved','IT Approved')
        tv2.heading("#0", text = "Item", anchor = 'w')
        tv2.column("#0", anchor = "w", width = 50)
        tv2.heading("#1", text = "Part No.")
        tv2.column("#1", anchor = "center",minwidth = 50, width = 200)
        tv2.heading("#2", text = "Description")
        tv2.column("#2", anchor = "center",minwidth = 50, width = 200)
        tv2.column("#3", anchor = "center",minwidth = 50, width = 200)
        tv2.heading("#3", text = "Quantity")
        tv2.column("#4", anchor = "center",minwidth = 50, width = 200)
        tv2.heading("#4", text = "Unit Price")
        tv2.column("#5", anchor = "center",minwidth = 50, width = 150)
        tv2.heading("#5", text = "Manager Approved")
        tv2.column("#6", anchor = "center",minwidth = 50, width = 150)
        tv2.heading("#6", text = "IT Approved")
        tv2.grid(row = 2, column = 1, columnspan = 9)
        self.treeview = tv
        self.treeview2 = tv2
        self.treeview.bind("<Button-1>", self.showLineItems)
        self.treeValues = {}


class PurchasePage(tk.Frame):

    def loadPOA(self):
        global employees
        print('loading Purchase Orders')
        self.treeview.delete(*self.treeview.get_children())
        cur.execute("""SELECT raisedby_employees_id, id, required_date, suppliers_id,supplier_email,supplier_webref,
                        quoteRef,FIsign,FIsigned,require_IT_signoff,IT_approved, manager_approved, status FROM PurchaseOrders""")
        count = 1
        #adds purchase orders that have not been purchased to treeview
        sup = {y:x for x,y in supplierID.items()}
        for a in cur.fetchall():
            temp = [a[0]]   
            if a[-4] == 'n':
                if a[-2] == 'y' and ('purchased' not in a[-1]):
                    if (a[-6] == 'y' and a[-5] == 'y') or (a[-6] == 'n' or None):
                        temp.extend(a[1:-2])
                        temp[3] = temp[3].rstrip()
                        temp[3] = sup[temp[3]]
                        self.treeview.insert('', 'end', text= count, values = (temp))
                        count += 1
            else:
                if a[-3] == 'y' and a[-2] == 'y' and ('purchased' not in a[-1]):
                    if (a[-6] == 'y' and a[-5] == 'y') or (a[-6] == 'n' or None):
                        temp.extend(a[1:-2])
                        temp[3] = temp[3].rstrip()
                        temp[3] = sup[temp[3]]
                        self.treeview.insert('', 'end', text= count, values = (temp))
                        count += 1
        count = 1
        print('done')

    def showLineItemsA(self, event):
        item = self.treeview.item(self.treeview.selection())
        self.treeview2.delete(*self.treeview2.get_children())
        #adds line items when a purchase order has been clicked on 
        if len(item['values']) > 1:
            poID = (item['values'][1])
            cur.execute("""SELECT partno, description, qty, unitprice, allowable_GLcodes_id, commoditycode, purchased FROM LineItems WHERE purchase_order_id = %i""" % (poID))
            count = 1
            for a in cur.fetchall():
                self.treeview2.insert('', 'end', text = count, values = list(a))
                count += 1
            count = 1

    def showLineItemsA2(self):
        #same as above but without event input 
        self.treeview2.delete(*self.treeview2.get_children())
        item = self.treeview.item(self.treeview.selection())
        if len(item['values']) > 1:
            poID = (item['values'][1])
            cur.execute("""SELECT partno, description, qty, unitprice, allowable_GLcodes_id, commoditycode, purchased FROM LineItems WHERE purchase_order_id = %i""" % (poID))
            count = 1
            for a in cur.fetchall():
                self.treeview2.insert('', 'end', text = count, values = list(a))
                count += 1
            count = 1

    def completePurchase(self):
        item = self.treeview2.item(self.treeview2.selection())
        num, vals = item['text'], item['values']
        pitem = self.treeview.item(self.treeview.selection())
        pnum, pvals = pitem['text'], pitem['values']
        #gets the purchase order number 
        purchaseOID = pvals[1]
        if len(vals) > 0:
            result = messagebox.askyesno(message = 'Have the items been purchased?')
            if result:
                sql = """UPDATE LineItems SET purchased = 'y' WHERE purchase_order_id = %i AND line_item = %i""" % (purchaseOID, num)
                cur.execute(sql)
                sql = """ UPDATE PurchaseOrders SET last_updated = getDate() WHERE id = %i""" % (purchaseOID)
                cur.execute(sql)
                conn.commit()
        cur.execute("""SELECT url FROM QuoteRefs where purchase_order_id = %i""" % (purchaseOID))
        for a in cur.fetchall():
            file = a[0]
        try :
            os.remove(str(purchaseOID) + file)
        except:
            pass
        self.showLineItemsA2()
        self.checkPOA(purchaseOID)
        self.win.lift()

    def checkPOA(self, purchaseOID):
        cur.execute("""SELECT purchased FROM LineItems WHERE purchase_order_id = %i""" % (purchaseOID))
        allPurch = True
        #checks that all line items have been purchased
        for a in cur.fetchall():
            if 'n' in a[0]:
                allPurch = False
        #if they have, update the purchase order status and last update
        if allPurch:
            misNum = None
            while not misNum:
                misNum = askstring("Add MIS Num", "Enter MIS Number for the PO")
            sql = """ UPDATE PurchaseOrders SET status = 'purchased', MIS_Number = '%s', last_updated = getDate() WHERE id = %i""" % (misNum,purchaseOID)
            cur.execute(sql)
            conn.commit()
            cur.execute("""select raisedby_employees_id, required_date, suppliers_id, total_price, date_raised from purchaseOrders where id = %i"""%(purchaseOID))
            for a in cur.fetchall():
                employId = a[0]
                reqDate = a[1]
                supId = a[2]
                tot = a[3]
                rDate = a[4]
            cur.execute("""Select last_name, first_name, Email from Employees2 where userid = '%s'"""%(a[0]))
            for a in cur.fetchall():
                fname = a[1]
                lname = a[0]
                eemail = a[2]
            messagetext = """<p>Hi %s,</p><p>A Purchase Order you raised has been purchased.</p>
                <table border="1"><tbody><tr>
                <td><strong>POID:</strong></td>
                <td><font color="Red">%i</font></td>
                </tr><tr><td><strong>Raised On:</strong></td><td>%s</td>
                </tr><tr><td><strong>Required By:</strong></td><td>%s</td>
                </tr><tr><td><strong>Supplier Name:</strong></td><td>%s</td>
                </tr><td><strong>Total Cost:</strong></td><td>%f</td>
                </tr><tr></tbody></table></tbody></table>
                """ % (fname, purchaseOID,rDate, reqDate, supId, tot)
            sql = """insert into EmailMessages
            (storedatetime,application_name,processed,processed_datetime,HTML_message,subject,
              send_to_address,from_address,cc_address,bcc_address,reply_to_address,message_text)
            values
            (getDate(),'PurchaseReq','N',0,'Y','Purchase Order %i has been approved',
            '%s', 'Purchase Order Admin<noreply@spilasers.com>', '','','%s', '%s')
            """ %  (purchaseOID,eemail,'noreply@spilasers.com', messagetext)
            print(eemail)
            #cur3.execute(sql)
            #conn3.commit()
            self.treeview2.delete(*self.treeview2.get_children())
            self.loadPOA()
            self.win.destroy()

    def showPartInfo(self):
        item = self.treeview2.item(self.treeview2.selection())
        num, vals = item['text'], item['values']
        pitem = self.treeview.item(self.treeview.selection())
        pnum, pvals = pitem['text'], pitem['values']
        orderID = pvals[1]
        cur.execute("""SELECT required_date FROM PurchaseOrders WHERE id = %i """ % (orderID))
        for a in cur.fetchall():
            requireDate = a[0]
        date = ''
        date += str(requireDate.year) + '-'
        date += str(requireDate.month) + '-'
        date += str(requireDate.day)
        #gets the supplier name from the ID
        if isinstance(pvals[3], str):
            supId = pvals[3].rstrip()
        else: supId = pvals[3]
        for key in supplierID:
            if supplierID[key] == supId:
                supId = key
        t = """Required Date:  %s  Supplier Name:  %s  Supplier ID: %s  Part No.: %i  Qty:  %i  Unit Price:  %s  GL_Code: %s  Commodity Code: %i """ % (date, supId,pvals[3].rstrip(), vals[0], vals[2], vals[3], vals[4], vals[6])
        #creates text box to display information for user to copy/paste
        text = tk.Text(self, width = 50, height = 10)
        text.grid(row = 0, column = 12, columnspan = 10)
        text.insert("1.0", t)
        text.configure(state = 'disabled')

    def showLineInfo(self):
        self.win = tk.Toplevel()
        tv2 = Treeview(self.win, height = 25)
        tv2['columns'] = ('Part No', 'Description', 'Qty', 'Unit Price', 'GL Code', 'Commodity Code', 'Purchased')
        tv2.heading("#0", text = "Item", anchor = 'w')
        tv2.column("#0", anchor = "w", width = 50)
        tv2.heading("#1", text = "Part No")
        tv2.column("#1", anchor = "center",minwidth = 50, width = 100)
        tv2.heading("#2", text = "Description")
        tv2.column("#2", anchor = "center",minwidth = 50, width = 100)
        tv2.heading("#3", text = "Qty")
        tv2.column("#3", anchor = "center",minwidth = 50, width = 100)
        tv2.heading("#4", text = "Unit Price")
        tv2.column("#4", anchor = "center",minwidth = 50, width = 100)
        tv2.heading("#5", text = "GL Code")
        tv2.column("#5", anchor = "center",minwidth = 50, width = 200)
        tv2.heading("#6", text = "Commodity Code")
        tv2.column("#6", anchor = "center",minwidth = 50, width = 200)
        tv2.heading("#7", text = "Purchased")
        tv2.column("#7", anchor = "center",minwidth = 50, width = 100)
        tv2.grid(row = 6, column = 0, columnspan = 8)
        for i in range(11):
            self.win.grid_rowconfigure(i,weight = 1)
            self.win.grid_columnconfigure(i,weight = 1)
        self.treeview2 = tv2
        self.treeview2.bind("<Button-2>", self.OnClick2)
        self.showLineItemsA2()
        button3 = tk.Button(self.win, text = "Complete Purchase", command = self.completePurchase)
        button3.grid(row = 50, column = 3)

    def showQuoteRef(self):
        global dirs
        pitem = self.treeview.item(self.treeview.selection())
        pnum, pvals = pitem['text'], pitem['values']
        cur.execute("""SELECT url from QuoteRefs Where purchase_order_id = %i""" % (pvals[1]))
        pic = []
        for a in cur.fetchall():
            pic.append(a[0])
        self.dm1['values'] = pic

    def OnClick(self, event):
        item = self.treeview.identify('item',event.x,event.y)
        item = self.treeview.item(self.treeview.selection())
        vals = item['values']
        col = int(self.treeview.identify('column',event.x,event.y)[-1])-1
        print(vals)
        print(col)
        print(vals[col])
        r = tk.Tk()
        r.withdraw()
        r.clipboard_clear()
        if isinstance(vals[col], str):
            r.clipboard_append(vals[col].rstrip())
        else: r.clipboard_append(vals[col])

    def OnClick2(self, event):
        item = self.treeview2.identify('item',event.x,event.y)
        item = self.treeview2.item(self.treeview2.selection())
        vals = item['values']
        col = int(self.treeview2.identify('column',event.x,event.y)[-1])-1
        print(vals)
        print(col)
        print(vals[col])
        r = tk.Tk()
        r.withdraw()
        r.clipboard_clear()
        if isinstance(vals[col], str):
            r.clipboard_append(vals[col].rstrip())
        else: r.clipboard_append(vals[col])

    def showQuoteRefs(self, *args):
        file = self.sVar.get()
        print(file)
        url = 'http://srv57mcsapps/get_purchasereq?file='
        cur.execute("""Select upurl from Settings""")
        for a in cur.fetchall():
            url = a[0]
        url = url + file
        r = requests.get(url)
        with open(file,"wb") as f:
            f.write(r.content)
        os.startfile(file)
        self.sVar.set(" ")

    def addCode(self):
        pitem = self.treeview.item(self.treeview.selection())
        pnum, pvals = pitem['text'], pitem['values']
        code = self.e1.get()
        cur.execute("""Update PurchaseOrders Set MIS_Number = '%s' where id = %i""" % (code, int(pvals[1])))
        conn.commit()
        messagebox.showinfo("Purchase Order Updated", "MIS_Number added was: %s" % (code))
        self.e1.delete(0,'end')

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        self.controller = controller
        self.createUI()
        label = tk.Label(self, text = "Purchase Page",font= ('Arial 15 bold'))
        label.grid(row = 0, column = 1, columnspan = 8)
        button1 = tk.Button(self, text = "Go to Start Page",
                         command=lambda: controller.show_frame("StartPage"))
        button1.grid(row = 6, column = 1)
        button2 = tk.Button(self, text = "Load Purchase Orders", command = lambda:self.loadPOA())
        button2.grid(row = 6, column = 2)
        #button3 = tk.Button(self, text = "Complete Purchase", command = self.completePurchase)
        #button3.grid(row = 6, column = 3)
        button5 = tk.Button(self, text = "Show Line Info", command = self.showLineInfo)
        button5.grid(row = 6, column = 4)
        button6 = tk.Button(self, text = "Show Quote Reference", command = self.showQuoteRef)
        button6.grid(row = 6, column = 5)
        button7 = tk.Button(self, text = "Add MIS Code", command = self.addCode)
        #button7.grid(row=6,column=6)
        self.e1 = tk.Entry(self, relief = SUNKEN, borderwidth = 3)
        #self.e1.grid(row = 7,column=6)
        self.sVar = tk.StringVar()
        self.dm1 = Combobox(self, textvariable = self.sVar, values = [], width = 20)
        self.dm1.grid(row = 7, column = 5)
        self.dm1.bind("<<ComboboxSelected>>",self.showQuoteRefs)
        for i in range(11):
            self.grid_rowconfigure(i,weight = 1)
            self.grid_columnconfigure(i,weight = 1)

    def createUI(self):
        tv = Treeview(self, height = 25)
        #creates treeview of purchase orders
        tv['columns'] = ('Raised By','Order ID', 'Required Date','Supplier Name', 'Email','Web Ref','Quote Ref')
        tv.heading("#0", text = "Item", anchor = 'w')
        tv.column("#0", anchor = "w", width = 40)
        tv.heading("#1", text = "Raised By")
        tv.column("#1", anchor = "center",minwidth = 50, width = 200)
        tv.heading("#2", text = "Order ID")
        tv.column("#2", anchor = "center",minwidth = 50, width = 75)
        tv.heading("#3", text = "Required Date")
        tv.column("#3", anchor = "center",minwidth = 50, width = 200)
        tv.heading("#4", text = "Supplier Name")
        tv.column("#4", anchor = "w", minwidth = 50, width = 250)
        tv.heading("#5", text = "Email")
        tv.column("#5", anchor = "center",minwidth = 50, width = 100)
        tv.heading("#6", text = "Web Ref")
        tv.column("#6", anchor = "center",minwidth = 50, width = 200)
        tv.heading("#7", text = "Quote Ref")
        tv.column("#7", anchor = "center",minwidth = 50, width = 100)
        tv.grid(row = 1, column = 1, columnspan = 8, rowspan = 5)
        self.treeview = tv
        tv2 = Treeview(self)
        self.treeview.bind("<Button-1>", self.showLineItemsA)
        self.treeview.bind("<Button-2>", self.OnClick)

if __name__ == "__main__":
    app = MyApp()
    app.wm_title("Purchase Req")
    logo = resource_path("spi.ico")
    app.iconbitmap(logo)
    app.mainloop()

