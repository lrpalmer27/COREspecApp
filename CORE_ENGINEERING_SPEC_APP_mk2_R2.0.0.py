import os
import sys
import sqlite3
import win32com.client
import re
from PyQt5.QtCore import QSize, Qt, QRegExp, pyqtSignal
from PyQt5.QtGui import QIcon, QPixmap, QIntValidator, QRegExpValidator
from PyQt5.QtWidgets import *
import ctypes
import time #for stats only, cuz thats cool
from datetime import date
# import LDAModel as ml ##this is the accompanying LDA model ML script I built.
import LDAModel as ml ##logan built this file to accompany this.

global version
version = "mk2 R2.0.0" #External Revision #2. Internal Revision #0, Working Copy #0
global ReleaseDate
ReleaseDate = "June 05 2023"

#paywall parameters here:
MaxNumUses=200
expiry_date = date(2024, 7, 1) #Yr, month , date (no leading zeros)

txt_size=14

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"CORE Engr. Spec Builder {version}")
        self.setWindowIcon(QIcon(os.path.join(cwd,"CoreEngrSmallLogo.ico")))
        self.setMinimumSize(QSize(750,225)) 
        
        self.Structure() #this is tab1
        self.BuildTab2()
        self.BuildTab3()
        self.BuildTab4()
        # self.BuildTab5()
        self.QueryDB()
        self.BuildTab1_extras()
        self.BuildTab3_extras()
        self.showMaximized()
        ##end of main driving section!
        
    def Structure(self):
        #Define the tabs here!
        self.tabs=QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tab1=QWidget()
        self.tabs.addTab(self.tab1, 'Build')

        #some tab1 pre-formatting!
        global Tab1
        global MBox
        global MainGroupbox
        Tab1 = QGridLayout()
        MainGroupbox=QGroupBox("Build Draft Project Spec! Select Desired Sections!")
        MainGroupbox.setStyleSheet(f"font-size: {txt_size}px;")
        
        MBox=QFormLayout()
        MainGroupbox.setLayout(MBox)
        
        global primeboxes #for population of prime section checkboxes
        primeboxes=[]
        global mechboxes #for population of mech section checkboxes   
        mechboxes=[] 

    def BuildTab1_checkBoxes(self,DivisionNumber,SubDivisionNumber,SubDivisionDescriptions,Details):   
        global tree
        rows=0 #start at 0 always
        tree= QTreeWidget()
        tmp=DivisionNumber
        tmp0=SubDivisionNumber
        tmp1=SubDivisionDescriptions
        
        #Division (1st lvl) checkbox
        DivisionNumber=QTreeWidgetItem(tree)
        DivisionNumber.setText(0,f"Division {tmp} - {tmp0} - {tmp1}")
        DivisionNumber.setFlags(DivisionNumber.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        ALLCheckBoxes.append(DivisionNumber)
        
        if tmp==1 or tmp ==2: ##experimental
            primeboxes.append(DivisionNumber)
            
        if tmp==10 or tmp ==21 or tmp ==22 or tmp == 23 or tmp == 25: ##experimental
            mechboxes.append(DivisionNumber)
            
        rows+=1
        #tree.setHeaderLabel(f"Division Number {tmp} Items") ##header hidden next line anyways
        tree.header().hide()

        #Details (2nd lvl) checkbox
        for det in Details:
            tmp=det
            rows+=1
            det=QTreeWidgetItem(DivisionNumber)
            det.setText(0,f"{tmp}")
            det.setFlags(det.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            det.setCheckState(0, Qt.Unchecked)
            ALLCheckBoxes.append(det)
        
        #tree.setExpanded(QModelIndex,SubDivisionNumber)
        tree.expandAll()
        tree.setUniformRowHeights(True)
        
        global scaleFactor
        scaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0)/100 ##this deals with windows (only?) scaling factor - some laptops have default 2.5x scaling, so app doesnt behave w/out this 
        height = int(rows*22*scaleFactor) 
            
        tree.setFixedHeight(height)
        MBox.addWidget(tree)
            
    def BuildTab1_extras(self):
        #makes the whole thing scrollable!        
        scroll=QScrollArea()
        scroll.setWidget(MainGroupbox)
        scroll.setWidgetResizable(True)
        Tab1.addWidget(scroll,2,1,1,3)

        #autocheck some sections 
        prime_chk=QPushButton("Auto-check prime sections\n#1,2",self)
        prime_chk.setStyleSheet(f"font-size: {txt_size}px;") 
        prime_chk.setCheckable(True)
        prime_chk.clicked.connect(self.chk_prime_button)
        Tab1.addWidget(prime_chk,1,1,1,1)
        
        mech_chk=QPushButton("Auto-check mech sections\n#10,21,22,23,25",self)
        mech_chk.setStyleSheet(f"font-size: {txt_size}px;")
        mech_chk.setCheckable(True)
        mech_chk.clicked.connect(self.chk_mech_button)
        Tab1.addWidget(mech_chk,1,2,1,1)
        
        ## button for accessing LDA model
        OpenML=QPushButton("Intelligent Selection\nTopic Modelling",self)
        OpenML.setStyleSheet(f"font-size: {txt_size}px;")
        OpenML.setCheckable(True)
        OpenML.clicked.connect(self.initOpenML)
        Tab1.addWidget(OpenML,1,3,1,1)
        
        #add gotime buttons
        GoTime=QPushButton("Build Draft Specifications!\n(Single Word Document)",self)
        GoTime.setStyleSheet(f"font-size: {txt_size}px;")
        GoTime.setCheckable(True)
        GoTime.clicked.connect(self.gotime_button)
        Tab1.addWidget(GoTime,3,3,1,1)
        
        MultidocBuild=QPushButton("Build Draft Specifications!\n(Individual Word Documents)",self)
        MultidocBuild.setStyleSheet(f"font-size: {txt_size}px;")
        MultidocBuild.setCheckable(True)
        MultidocBuild.clicked.connect(self.MultidocBuildr)
        Tab1.addWidget(MultidocBuild,3,1,1,2)
            
        self.tab1.setLayout(Tab1)

    def initOpenML(self):
        print('Initializing AI Selection Window')
        # chk_these=MLwindow.go(app,cwd)
        ml.est_access()
        self.w=MLWindow()
        self.w.setAttribute(Qt.WA_DeleteOnClose)  # Ensures proper cleanup
        self.w.closed.connect(self.MLwindow_closed)  # Connect close signal to slot
        self.w.show()
        
    def MLwindow_closed(self):
        # print(MLapproved_chkboxs)
        ### format is: MLapproved_chkboxs=[[Div#,item,item,item],[Div#,item,item,item],[Div#,item,item,item]]
        tempo=[]
        for n in range(0,len(MLapproved_chkboxs)):
            for i in ALLCheckBoxes: 
                if i.text(0) == MLapproved_chkboxs[n][0]:
                    L=i.text(0)
                    h=i.parent()
                    tempo.append(i)
                elif i.parent() != None:
                    if i.parent().text(0) == MLapproved_chkboxs[n][0]: 
                        h=i.parent()
                        hh=i.parent().text(0)
                        tempo.append(i)
                        
            for b in tempo: 
                for o in range(1,len(MLapproved_chkboxs[n])): #erica experimental
                    if b.text(0) == MLapproved_chkboxs[n][o]:
                        b.setCheckState(0, Qt.Checked)
                    
        print("Auto-checked",len(MLapproved_chkboxs))

    def chk_prime_button(self):
        # print("check prime sections!")
        
        for i in primeboxes: 
            i.setCheckState(0, Qt.Checked)

    def chk_mech_button(self):
        # print("check mech sections now!")
        
        for i in mechboxes: 
            i.setCheckState(0, Qt.Checked)

    def MultidocBuildr(self):
        Checked_Boxes=["start"]
        #ALLCheckBoxes.reverse()
        for item in ALLCheckBoxes:
            if item.checkState(0) == 1 or item.checkState(0) == 2: 
                checkboxName=item.text(0)
                Checked_Boxes.append(checkboxName) 
        #print('These checkboxes are checked!')
        #print(Checked_Boxes)
        
        SaveFilesHere=str(QFileDialog.getExistingDirectory(self, "Select Where to Save Output Files"))
        if SaveFilesHere == None: 
            msg = QMessageBox()
            msg.setWindowTitle("Error") 
            msg.setText(f"No directory chosen to save output files, saving here:\n {cwd}")
            msg.setIcon(QMessageBox.Critical)
            msg.exec_()
            SaveFilesHere=cwd
        
        #now prepare 'em to be sent to re-builder class!
        DocumentLeadingRegex="Division [0-9]{1,2} - [0-9]{2}[ ]{1}[0-9]{2}[ ]{1}[0-9]{2}[.]{0,1}[0-9]{0,2}.*" 
        global send
        send=[]
        Send_subsections=[]
        Checked_Boxes.append("end!")
        for chkd in Checked_Boxes: 
            DivNum_SectionDesc = re.findall(DocumentLeadingRegex,chkd)
            if DivNum_SectionDesc != [] or chkd=="end!" or chkd=="start": ## enter if its formatted like: "Division ## - [0-9]{2}[ ]{1}[0-9]{2}[ ]{1}[0-9]{2}[.]{0,1}[0-9]{0,2}.*"
                if chkd =="start":
                    preamble() ##open word and gets shit ready
                    
                    
                elif chkd =="end!":
                    send.append(Send_subsections)
                    Rebuilder()
                    Final()
                    print("All Done!")
                
                elif send != [] and DivNum_SectionDesc != []: #meaning that this is the second division number leader -- 
                    ##we need to send along pervious complaition and reset lists
                    send.append(Send_subsections) ## compiles send list into: [Div# (as int) - subdivision# (as int), [details]] format
                    Rebuilder()
                    Final()
                    saveWnewname(SaveFilesHere,outputname=f'{(DivNum_SectionDesc[0])[:-4]}') #saves draft document as this name
                    #reset driver lists
                    send=[]
                    Send_subsections=[]
                    send.append(chkd[:-3])
                     
                elif DivNum_SectionDesc != []: ## meaning this is round one and we are adding the leader to the send list.
                    saveWnewname(SaveFilesHere,outputname=f'{(DivNum_SectionDesc[0])[:-4]}') #saves draft document as this name
                    send.append(chkd[:-3]) #this appends the division number list item
            
    
            else: 
                Send_subsections.append(chkd) #this appends anything else   


    def gotime_button(self):
        Checked_Boxes=["start"]
        #ALLCheckBoxes.reverse()
        for item in ALLCheckBoxes:
            if item.checkState(0) == 1 or item.checkState(0) == 2: 
                checkboxName=item.text(0)
                Checked_Boxes.append(checkboxName)
        
        #print('These checkboxes are checked!')
        #print(Checked_Boxes)
        
        SaveFileHere=str(QFileDialog.getExistingDirectory(self, "Select Where to Save Output File"))
        if SaveFileHere == None: 
            msg = QMessageBox()
            msg.setWindowTitle("Error") 
            msg.setText(f"No directory chosen to save output file, saving here:\n {cwd}")
            msg.setIcon(QMessageBox.Critical)
            msg.exec_()
            SaveFilesHere=cwd
        
        #now prepare 'em to be sent to re-builder class!
        DocumentLeadingRegex="Division [0-9]{1,2} - [0-9]{2}[ ]{1}[0-9]{2}[ ]{1}[0-9]{2}[.]{0,1}[0-9]{0,2}.*" 
        global send
        send=[]
        Send_subsections=[]
        Checked_Boxes.append("end!")
        for chkd in Checked_Boxes: 
            if re.findall(DocumentLeadingRegex,chkd) != [] or chkd=="end!" or chkd=="start": ## enter if its formatted like: "Division ## - [0-9]{2}[ ]{1}[0-9]{2}[ ]{1}[0-9]{2}[.]{0,1}[0-9]{0,2}.*"
                if chkd =="start":
                    preamble() ##open word and gets shit ready
                    saveWnewname(SaveFileHere) ##saves the output spec file with name OutputSpecDraftDocument
                
                elif chkd =="end!":
                    send.append(Send_subsections)
                    Rebuilder()
                    Final()
                    print("All Done!")
                
                elif send != []: #meaning that this is the second division number leader -- 
                    ##we need to send along pervious complaition and reset lists
                    send.append(Send_subsections) ## compiles send list into: [Div# (as int) - subdivision# (as int), [details]] format
                    #print("send to rebuilder!",send) ##call rebuilder on send now
                    Rebuilder()
                    #reset driver lists
                    send=[]
                    Send_subsections=[]
                    send.append(chkd[:-3])
                     
                elif send == []: ## meaning this is round one and we are adding the leader to the send list.
                    send.append(chkd[:-3]) #this appends the division number list item
    
            else: 
                Send_subsections.append(chkd) #this appends anything else
   
    def QueryDB(self):
        conn = sqlite3.connect(DB_Path) 
        cur=conn.cursor()

        #Get all division numbers in Database
        divisions_cur=cur.execute('''SELECT DISTINCT(Division_Number) FROM Components''')
        
        #get into a list so we can use cursor elsewhere
        divisions=[]
        global ALLCheckBoxes
        ALLCheckBoxes=[] #use this list as master list of all checkboxes!
        for i in divisions_cur:
            divisions.append(int(str(i[0])))
            
        for div in divisions:
            #divNum=int(str(div[0])) 
            #DivisionNumbers.append(divNum)
            subdivisions_cur= cur.execute(f'''SELECT DISTINCT(Subdivision_ID_Number)
                                FROM Components
                                WHERE Division_Number = {div}''')
            
            #get into a list so we can use cursor elsewhere
            subdivisions=[]
            for i in subdivisions_cur:
                subdivisions.append(str(i[0]))
            
            for SubDiv in subdivisions: 
                #SubDivNum=(str(subDiv[0]))
                #SubdivisionNumbers.append(SubDivNum)
                Descript_cur=cur.execute(f''' SELECT DISTINCT (Subdivision_Description)
                                                    FROM Components
                                                    WHERE Subdivision_ID_Number ="{SubDiv}"
                                                    LIMIT 1''')
                
                #Non scriptable :( --> this switches it into string
                for Desc in Descript_cur: 
                    SubDivisionDescriptions=(str(Desc[0]))
                    
                Deets=cur.execute(f'''SELECT DISTINCT(Details) 
                                FROM Components
                                WHERE Subdivision_ID_Number ="{SubDiv}"
                                ''')
                
                #Non scriptable :( --> this switches it into list of strings
                Details=[]
                for det in Deets: 
                    Details.append(str(det[0]))
                    
                '''#Troubleshooting
                print("Division Number:",div)
                print("Subdivision Number:",SubDiv,", Description:",SubDivisionDescriptions)
                print("Details:", Details)'''
                self.BuildTab1_checkBoxes(div,SubDiv,SubDivisionDescriptions,Details)     
        return 
               
    def BuildTab2(self):
        # Add some widgets to tab 2
        self.tab2=QWidget()
        self.tabs.addTab(self.tab2, 'Prepare DB')
        
        Tab2 = QVBoxLayout()
        Tab2GroupBox=QGroupBox("Prepare Spec Section Holding Database!")  
        Tab2GroupBox.setStyleSheet(f"font-size: {txt_size+4}px;")     
        Tab2BoxLayout=QFormLayout()
        Tab2GroupBox.setLayout(Tab2BoxLayout)
        
        How2BuildDB=QLabel("BASIC REQUIREMENTS:\n-All program specific files must live in one directory on your local* machine.\n-Imported specs must be saved in one directory in the format:\n\t'Directory' > Division# > (word docs here)\n\tWe will select the 'Directory' later.\n-Imported specs must be in any word file format (.doc/.docx is fine).\n-Please close all other programs when running this.\n\n*promote use of local machine so we dont get any server communication timeouts (not built to deal with this)\n\nSTEPS TO RUN: \n1. Click 'Select Division Holding Folder' button below. Navigate to the 'Directory' (see above for formatting).\n2. Click 'Build DB'\n3. Wait for program to run. This may take a while.\n4. After completion a list of items that had trouble saving correctly will be output. Save these manually through the 'Add Items to DB' tab (main menu bar)") ##Mega explaination how to build DB!
        How2BuildDB.setStyleSheet(f"font-size: {txt_size}px;")
        Tab2BoxLayout.addRow(How2BuildDB) 
        
        scl2=QScrollArea()
        scl2.setWidget(Tab2GroupBox)
        scl2.setWidgetResizable(True)
        Tab2.addWidget(scl2)
        
        global DivDirectory #need to drop dir in here later
        DivDirectory = None
        
        GetDir=QPushButton("Select Division Holding Folder",self)
        GetDir.setStyleSheet(f"font-size: {txt_size}px;")
        GetDir.setCheckable(True)
        GetDir.clicked.connect(self.GetDir)
        Tab2.addWidget(GetDir)
        
        PrepDB=QPushButton("Build DB!",self)
        PrepDB.setStyleSheet(f"font-size: {txt_size}px;")
        PrepDB.setCheckable(True)
        PrepDB.clicked.connect(self.GateKeeper)
        Tab2.addWidget(PrepDB)
        
        self.tab2.setLayout(Tab2)

    def BuildTab4(self):
        # Add some widgets to tab 2
        self.tab4=QWidget()
        self.tabs.addTab(self.tab4, 'Extras')
        Tab4 = QGridLayout()
        
        conn = sqlite3.connect(DB_Path) 
        cur=conn.cursor()
        
        cur.execute('SELECT NumFilesInDB FROM Stats')
        NumFilesInDB=cur.fetchone()
        if NumFilesInDB is not None: 
            NumFilesInDB=NumFilesInDB[0]
            cur.execute('SELECT NumFilesInRawSpecs FROM Stats LIMIT 1')
            NumFilesInRawSpecs=cur.fetchone()[0]
            cur.execute('SELECT MissingFromDB FROM FixManually LIMIT 1')
            MissingFromDB=cur.fetchone()[0]
            cur.execute('SELECT MissingFromRawSpecs FROM FixManually LIMIT 1')
            MissingFromRawSpecs=cur.fetchone()[0]
            cur.execute('SELECT NumMissedByRegex FROM FixManually LIMIT 1')
            NumMissedbyRegex=cur.fetchone()[0]
            cur.execute('SELECT errorOnRegex FROM FixManually LIMIT 1')
            ErroredonRegex=cur.fetchone()[0]
        else: 
            NumFilesInDB=NumFilesInRawSpecs=MissingFromDB=MissingFromRawSpecs=NumMissedbyRegex=ErroredonRegex='None'

        PrepDBDeets=QLabel(f'''There are {NumFilesInDB} files in the database \nThere are {NumFilesInRawSpecs} files in the raw specs folder
                           \nThe files missing from DB are:{MissingFromDB} \nThe files missing from RawSpecs are:{MissingFromRawSpecs}
                           \nThere were {NumMissedbyRegex} items missed by regex \nThe specs in RawSpecs missed by regex are: \n{ErroredonRegex}''') 
        PrepDBDeets.setStyleSheet(f"font-size: {txt_size}px;")
        Tab4.addWidget(PrepDBDeets,1,1,1,1)


        cur.execute("SELECT COUNT(*) FROM Components")
        TotalRows_inDB=cur.fetchone()[0]
        cur.close()
        conn.close()
        
        RowCountinDB=QLabel(f"There are {TotalRows_inDB} unique entries in the database to pull from!") 
        RowCountinDB.setStyleSheet(f"font-size: {txt_size}px;")
        Tab4.addWidget(RowCountinDB,1,2,1,1) 
        
        End_time=time.time()
        Startup_time = round((End_time-start_time),2) #yes I know there are things that happen after this, so it isnt a great understanding of startup time, but it'll be a good relative measure of startup time so long as we dont change the def'n across versions.
        
        RowCountinDB=QLabel(f"It took {Startup_time}s for this app to startup") 
        RowCountinDB.setStyleSheet(f"font-size: {txt_size}px;")
        Tab4.addWidget(RowCountinDB,1,3,1,1) 
        
        
        # RowCountinDB=QLabel("Use this spot to hold total db entries over time - show a graph here!") 
        # Tab4.addWidget(RowCountinDB,1,3,1,1) 
        
        
        label = QLabel(self)
        pixmap = QPixmap(os.path.join(cwd,"CoreEngrLogo.png"))
        label.resize(150, 150)
        label.setPixmap(pixmap)
        Tab4.addWidget(label,2,2,1,1)
        
        How2BuildDB=QLabel(f"DEVELOPPED FOR CORE ENGINEERING - 2023\nVERSION: {version}\nRELEASE DATE: {ReleaseDate}") 
        How2BuildDB.setStyleSheet(f"font-size: {txt_size}px;")
        Tab4.addWidget(How2BuildDB,2,1,1,1) 
        
        How2BuildDB=QLabel("IDEA: DAVE ENNIS FALL 2022\nIMPLEMENTATION: L.PALMER 2022-2023 - CONTACT: logan@palmers.ca\nPROJECT DESCRIPTION: www.lrpalmer.com/core") 
        How2BuildDB.setStyleSheet(f"font-size: {txt_size}px;")
        Tab4.addWidget(How2BuildDB,2,3,1,1) 

        self.tab4.setLayout(Tab4)
    

    def BuildTab3(self):
        # Add some widgets to tab 2
        self.tab3=QWidget()
        self.tabs.addTab(self.tab3, 'Add Items to DB')
        global Tab3
        Tab3 = QGridLayout()
        
        ##Set all fields here!
        Tab3.addWidget(QLabel("DIVISION NUMBER"),1,1,1,1) 
        self.AddDivNumber=QLineEdit(self)
        self.AddDivNumber.setPlaceholderText("1")
        self.AddDivNumber.setValidator(QIntValidator())
        Tab3.addWidget(self.AddDivNumber,1,2,1,1)
        
        Tab3.addWidget(QLabel("SECTION NUMBER"),1,3,1,1)
        self.AddSectionNumber=QLineEdit(self)
        SectionNumberRegex_validator = QRegExp("[0-9]{2}[ ]{1}[0-9]{2}[ ]{1}[0-9]{2}[.]{0,1}[0-9]{0,2}")
        SectionNumbervalidator = QRegExpValidator(SectionNumberRegex_validator)
        self.AddSectionNumber.setValidator(SectionNumbervalidator)
        self.AddSectionNumber.setPlaceholderText("00 00 00 or 11 11 11.11")
        Tab3.addWidget(self.AddSectionNumber,1,4,1,1)
        
        Tab3.addWidget(QLabel("SECTION DESCRIPTION"),1,5,1,1)
        self.AddSectionDescription=QLineEdit(self)
        self.AddSectionDescription.setPlaceholderText("Bid Depository Sections")
        Tab3.addWidget(self.AddSectionDescription,1,6,1,1)
        
        Tab3.addWidget(QLabel("DETAILS"),2,1,1,1)
        self.AddDetails=QTextEdit(self)
        self.AddDetails.setPlaceholderText("Detail one\nDetail two\nDetail three\n...")
        Tab3.addWidget(self.AddDetails,2,2,1,1)
        
        #Validate and add to db button
        ValidateAndAdd=QPushButton()
        ValidateAndAdd.setStyleSheet(f"font-size: {txt_size}px;")
        ValidateAndAdd.setText("Validate and Add to Master DB")
        ValidateAndAdd.clicked.connect(self.ValidateAndAdd_button)
        Tab3.addWidget(ValidateAndAdd,3,1,1,4)
        
        DELETEButton=QPushButton()
        DELETEButton.setStyleSheet(f"font-size: {txt_size}px;")
        DELETEButton.setText("DELETE THE SELECTED ROW")
        DELETEButton.clicked.connect(self.DELETEButton_clicked)
        Tab3.addWidget(DELETEButton,3,5,1,3)
        
        ##make a scrollbar to select where we want to add this custom item!
        global Where2Go_gbox
        Where2Go_gbox=QGroupBox("INSERT ABOVE")
        Where2Go_gbox.setStyleSheet(f"font-size: {txt_size+4}px;")
        Tab3.addWidget(Where2Go_gbox,2,3,1,6)
        Where2Go_gbox_lay=QFormLayout()
        #Where2Go_gbox.setLayout(Where2Go_gbox_lay)
        
        #get db data for this for loop
        conn = sqlite3.connect(DB_Path) 
        cur=conn.cursor()
        cur.execute('SELECT * FROM Components ORDER BY ID')
        
        rows=cur.fetchall()
        
        #Table here!
        table_widget = QTableWidget()
        table_widget.setColumnCount(5)
        table_widget.setRowCount(len(rows))
        table_widget.setHorizontalHeaderLabels(['Checkboxes', 'DIVISION #', 'SUBDIVISION #', 'SUBDIVISION DESC.', 'DETAILS'])
        table_widget.setColumnWidth(3,150)
        table_widget.setColumnWidth(4,300)
        table_widget.verticalHeader().setVisible(False)
        
        global Names_as_chkboxes
        Names_as_chkboxes=[] #drop checkbox objects in here later
        Names=[]
        for i in range(0,len(rows)):
            Names.append("Row "+str(i+1))
         
        for name in Names:
            row=Names.index(name) ##note this "row" index number is lagging one behind "name"
            self.BuildTab3_list(name,rows[row],row,table_widget)
            
        #table_widget.resizeColumnsToContents()
        
        Where2Go_gbox_lay.addWidget(table_widget)
        Where2Go_gbox.setLayout(Where2Go_gbox_lay)
    
        self.tab3.setLayout(Tab3)    

    def DELETEButton_clicked(self): 
        #validate not more than one is cheked!!        
        LocationChecked=[]
        for i in Names_as_chkboxes: 
            if i.isChecked():
                LocationChecked.append(i.text())    

        if len(LocationChecked)==1:
            RowName=LocationChecked[0] #must be "Row ######"
            RowNumber=int(RowName[4:])
            
        else: 
            #throw an error! only one check at a time
            msg = QMessageBox()
            msg.setWindowTitle("Error") ##
            msg.setText("Can only choose one! Restart app please.") #might be able to get it to set all to not-checked
            msg.setIcon(QMessageBox.Critical)
            msg.exec_()
            return
            #sys.exit() ## kills app
            
        msg=QMessageBox()
        #msg.setWindowTitle("DELETING A ROW CANNOT BE UNDONE!")
        del_or_nah=msg.question(self,'DELETING A ROW CANNOT BE UNDONE!',f"DELETE ROW #{RowNumber}?",msg.Yes|msg.No)
        #msg.setIcon(QMessageBox.Critical)
        #msg.exec_()
        
        if del_or_nah == msg.No: 
            return
        elif del_or_nah ==msg.Yes:
            self.deleteRowSQL(RowNumber)
        
    def deleteRowSQL(self,IDNo2Del):
        #print("delete row number", IDNo2Del)
        conn = sqlite3.connect(DB_Path) 
        cur=conn.cursor()
        
        try: 
            cur.execute('''PRAGMA foreign_keys=off''')
            cur.execute('''CREATE TABLE temp_table AS SELECT * FROM Components''')
            
            cur.execute(f'''DELETE FROM temp_table WHERE ID = {IDNo2Del}''')
            cur.execute(f'''UPDATE temp_table SET ID = ID-1 WHERE ID>{IDNo2Del}''')
            cur.execute('''DELETE FROM Components''')
            cur.execute('''INSERT INTO Components SELECT * FROM temp_table''')
            cur.execute('''DROP TABLE temp_table''')
            conn.commit()
            cur.close()
            conn.close()
            
        except: 
            print("somethign didnt work!?")
        
    def BuildTab3_list(self,name,deets,rowindex,table_widget):        
        #print(name, deets)
        #print(deets) #(1, '1', '01 00 00', 'BID Depository Sections.doc', 'GENERAL\r')
        Ntemp=name
        name = QCheckBox()
        name.setText(f"{Ntemp}")
        name.setCheckState(False)
        Names_as_chkboxes.append(name)
        table_widget.setCellWidget(rowindex, 0, name)
                
        for col in range(1,5):
            Cell_data=deets[col]
            item=QTableWidgetItem(Cell_data)
            item.setTextAlignment(Qt.AlignLeft) 
            table_widget.setItem(rowindex,col,item)
    
    def BuildTab3_extras(self):
        scroll=QScrollArea()
        scroll.setWidget(Where2Go_gbox)
        scroll.setWidgetResizable(True)
        Tab3.addWidget(scroll,2,3,1,6)
        
        self.tab3.setLayout(Tab3)

    def ValidateAndAdd_button(self):
        ## button at the bottom of tab3
        NewDivNumber=self.AddDivNumber.text()
        NewSectionNumber=self.AddSectionNumber.text()
        NewSectionDescription=self.AddSectionDescription.text()
        NewDetails=self.AddDetails.toPlainText()
        
        #validate not more than one is cheked!!        
        LocationChecked=[]
        for i in Names_as_chkboxes: 
            if i.isChecked():
                LocationChecked.append(i.text())    

        if len(LocationChecked)==1:
            RowName=LocationChecked[0] #must be "Row ######"
            RowNumber=int(RowName[4:])
            
        else: 
            #throw an error! only one check at a time
            msg = QMessageBox()
            msg.setWindowTitle("Error") ##
            msg.setText("Can only choose one place to insert! Verify only one box is checked!") #might be able to get it to set all to not-checked
            msg.setIcon(QMessageBox.Critical)
            msg.exec_()
            return
            #sys.exit() ## kills app
            
        ###now add into db!
        print("stuff",NewDivNumber,NewSectionNumber,NewSectionDescription,NewDetails) #this proves we have all the data in, as these variables
        print("rownumber",RowNumber)
        self.AddSingleToDB(RowNumber,NewDivNumber, NewSectionNumber, NewSectionDescription, NewDetails)
        
    def AddSingleToDB(self,IDNo,Division_Number, Subdivision_ID_Number, Subdivision_Description, Details):
        conn = sqlite3.connect(DB_Path) 
        cur=conn.cursor()

        try:
            cur.execute('''PRAGMA foreign_keys=off''')
            cur.execute('''CREATE TABLE temp_table AS SELECT * FROM Components
                        ''')
            PrevIDNo=IDNo-1
            cur.execute(f'''UPDATE temp_table SET ID = ID+1 WHERE ID>{PrevIDNo}
                        ''')
            cur.execute(rf'''INSERT INTO temp_table (ID, Division_Number, Subdivision_ID_Number, Subdivision_Description, Details)
                        VALUES ({IDNo}, '{Division_Number}', '{Subdivision_ID_Number}', '{Subdivision_Description}', '{Details}')
                        ''')
            cur.execute(''' DELETE FROM Components
                        ''')
            cur.execute('''INSERT INTO Components SELECT * FROM temp_table
                        ''')
            ##add line here to record date changes in a new table for future ref
            cur.execute('''DROP TABLE temp_table;
                        ''')
            cur.execute('''PRAGMA foreign_keys=on;
                        ''')
            conn.commit()
            cur.close()
            conn.close()
            
        except: 
            cur.execute('''DROP TABLE temp_table;
                        ''')

    def GetDir(self):
        global DivDirectory
        DivDirectory = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        print("folder", DivDirectory)
        
    def GateKeeper(self):
        if DivDirectory == None: 
            msg = QMessageBox()
            msg.setWindowTitle("Error") 
            msg.setText("Please Select a Directory!")
            msg.setIcon(QMessageBox.Critical)
            msg.exec_()
            self.GetDir()
            
        if DivDirectory == None:
            print("Killed program - need to select folder here!")
            sys.exit()
                
        print("build DB now!")
        BuildDB()

class MLWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Intelligent Selections")
        self.setWindowIcon(QIcon(os.path.join(cwd,"CoreEngrSmallLogo.ico")))
        self.setMinimumSize(QSize(750,225))
        self.tabs=QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tab1=QWidget()
        self.Query()
        self.Train()
        self.showMaximized()
        
    closed =pyqtSignal()
    
    def closeEvent(self, event):
        # Emit the custom signal when window is closed
        self.closed.emit()
        event.accept()
       
    def Train(self):
        self.tab2=QWidget()
        self.tabs.addTab(self.tab2, 'Train')
        
        Tab2Layout = QGridLayout()
        ExplainML=QLabel("This training tab allows you to initiate training the 'Pre-Trained' comprehension model. This app comes with a pre-trained model on the basic specs provided by the NL gov't.\nRe-train when you add custom specs such that the AI algorithm will consider these custom specs. Training takes about 1 minute.\nNothing is needed in order to train. It will pull data direct from the master database.")
        ExplainML.setStyleSheet(f"font-size: {txt_size}px;")
        Tab2Layout.addWidget(ExplainML,1,1,1,2)
        
        #Train button!
        train=QPushButton()
        train.setStyleSheet(f"font-size: {txt_size}px;")
        train.setText("Train!")
        train.clicked.connect(self.trainClicked)
        Tab2Layout.addWidget(train,3,1,1,2)
        
        self.tab2.setLayout(Tab2Layout)
    
    def trainClicked(self):
        print("Training LDA Model Now!")
        ml.train()
    
    def Query(self):
        # Add some widgets to tab 2
        self.tab5=QWidget()
        self.tabs.addTab(self.tab5, 'Query')
        Tab5Layout = QGridLayout()
        ExplainML=QLabel("The objective of this tab is to use a pre-trained machine learning model to intelligently pre-select sections that may be relevant to the project you are working on, based on an input of a few keywords input below;")
        ExplainML.setStyleSheet(f"font-size: {txt_size}px;")
        Tab5Layout.addWidget(ExplainML,1,1,1,2)
        
        global InputKeywords_GB_Layout
        InputKeywords_GB=QGroupBox("Click 'Enter' to create another Keyword input box, and click tab to drop down a level. Enter as many keywords as you'd like.")
        InputKeywords_GB.setStyleSheet(f"font-size: {txt_size+4}px;")
        InputKeywords_GB_Layout=QFormLayout()
        InputKeywords_GB_Layout.setAlignment(Qt.AlignTop)
        InputKeywords_GB.setLayout(InputKeywords_GB_Layout)
        
        #add keyword input blocks here
        Keyword1=QLabel("Keyword:")
        Keyword1.setStyleSheet(f"font-size: {txt_size}px;")
        self.KeywordInputBoxes_all=[]
        KeywordInputBox=QLineEdit(self)
        KeywordInputBox.setPlaceholderText("AHU")
        KeywordInputBox.returnPressed.connect(self.KeywordInputBoxPressed)
        self.KeywordInputBoxes_all.append(KeywordInputBox)
        InputKeywords_GB_Layout.addRow(Keyword1,KeywordInputBox)    

        scl_keyGB=QScrollArea()
        scl_keyGB.setWidget(InputKeywords_GB)
        scl_keyGB.setWidgetResizable(True)
        Tab5Layout.addWidget(scl_keyGB,2,1,1,2)
        
        #add input for number of related subsections and pct related we want
        n=QLabel("Output Number of Related Sections")
        n.setStyleSheet(f"font-size: {txt_size}px;")
        Tab5Layout.addWidget(n,3,1,1,2)
        numslider = QSlider(Qt.Horizontal)
        numslider.setMinimum(1)
        numslider.setMaximum(10)
        numslider.setValue(4)
        numslider.setSingleStep(1)
        numslider.setTickInterval(1)
        numslider.setTickPosition(QSlider.TicksBelow)
        Tab5Layout.addWidget(numslider,4,1,1,1)
        numslider.valueChanged.connect(self.nsliderval)
        self.nlabel = QLabel("Value: {}".format(numslider.value()))
        self.nlabel.setStyleSheet(f"font-size: {txt_size}px;")
        Tab5Layout.addWidget(self.nlabel,5,1,1,1)
        
        thr=QLabel("How closely related they are")
        thr.setStyleSheet(f"font-size: {txt_size}px;")
        Tab5Layout.addWidget(thr,3,2,1,2)
        pctSlider = QSlider(Qt.Horizontal)
        pctSlider.setMinimum(10)
        pctSlider.setMaximum(90)
        pctSlider.setValue(50)
        pctSlider.setTickInterval(5)
        pctSlider.setSingleStep(5)
        pctSlider.setTickPosition(QSlider.TicksBelow)
        Tab5Layout.addWidget(pctSlider,4,2,1,1)
        pctSlider.valueChanged.connect(self.pctsliderval)
        self.pctlabel = QLabel("Value: {}%".format(pctSlider.value()))
        self.pctlabel.setStyleSheet(f"font-size: {txt_size}px;")
        Tab5Layout.addWidget(self.pctlabel,5,2,1,1)
        
        MLGoButton=QPushButton()
        MLGoButton.setStyleSheet(f"font-size: {txt_size}px;")
        MLGoButton.setText("Auto-check relevant sections")
        MLGoButton.clicked.connect(self.MLGoButtonClicked)
        Tab5Layout.addWidget(MLGoButton,6,1,1,2)
        
        self.sliders=[numslider,pctSlider]
        
        self.tab5.setLayout(Tab5Layout)
    
    def nsliderval(self,value):
        self.nlabel.setText("Value: {}".format(value))
    
    def pctsliderval(self,value):
        self.pctlabel.setText("Value: {}%".format(value))
      
    def KeywordInputBoxPressed(self):
        Keyword=QLabel("Keyword:")
        Keyword.setStyleSheet(f"font-size: {txt_size}px;")
        KeywordInputBox=QLineEdit(self)
        KeywordInputBox.setPlaceholderText("AHU")
        KeywordInputBox.returnPressed.connect(self.KeywordInputBoxPressed)
        self.KeywordInputBoxes_all.append(KeywordInputBox)
        InputKeywords_GB_Layout.addRow(Keyword,KeywordInputBox) 

    def MLGoButtonClicked(self):
        ##check which keyword slots have keywords, and what are they?
        # NumKeywords=len(self.KeywordInputBoxes_all)
        keywrds_text=[]
        for i in self.KeywordInputBoxes_all:
            if i.text() != '':
                keywrds_text.append(i.text())
        
        # print("Keywords:",keywrds_text)
        
        Query_op_num=self.sliders[0].value()
        Query_op_thresh=self.sliders[1].value()/100
        
        #make a slider for these
        # Query_op_thresh=0.5
        # Query_op_num=3
        
        self.initQuery(keywrds_text,Query_op_num,Query_op_thresh)

    def initQuery(self,keywords,n,thresh):
        #n is number of related items each keyword can grab
        #thresh is pct related they need to be
        
        if len(keywords) != 0: 
            print("Querying Model...")
            Pre_Trained_LDAModel=ml.retrieve_fromDB()
            global relevantSubSectionNames
            relevantSubSectionNames=ml.test_performance(Pre_Trained_LDAModel,keywords,n,thresh)
            
            print(relevantSubSectionNames)
            self.MLkeepEm(relevantSubSectionNames)
        
        # return relevantSubSectionNames #Where does this pass it to?? How do we get it back to main?
    
        elif len(keywords) == 0: 
            msg = QMessageBox()
            msg.setWindowTitle("Error") 
            msg.setText(f"Please input some keywords!")
            msg.setIcon(QMessageBox.Critical)
            msg.exec_()
            return
    
    def MLkeepEm(self,relevantSubSectionNames):
        dialog = QDialog(self)
        # dialog.setGeometry()
        dialog.setMinimumSize(QSize(750,225))
        dialog.setWindowTitle("Approve Selections")
        layout = QGridLayout(dialog)
        label = QLabel("Please approve the following Intelligent Slections", dialog)
        layout.addWidget(label,1,1,1,3)
        
        #approve button 
        approve=QPushButton("Approve",self)
        approve.setStyleSheet(f"font-size: {txt_size}px;") 
        approve.setCheckable(True)
        approve.clicked.connect(self.approvechk)
        layout.addWidget(approve,3,2,1,1)
             
        ##get sections and subsections details here
        DNumber=[]
        SubDivisionNumber=[]
        SubDivisionDescription=[]
        Details=[]
        for l in relevantSubSectionNames:
            cur.execute(f"SELECT * FROM components WHERE Subdivision_Description = '{l}'")
            lp=cur.fetchall()

            DNumber.append(lp[0][1])
            SubDivisionNumber.append(lp[0][2])
            SubDivisionDescription.append(lp[0][3])
            t=[]
            for i in range(0,len(lp)):
                t.append(lp[i][4])
            Details.append(t)
        
        ##below here puts it into a qbox with a scrollbar and checkboxes.
        box=QGroupBox("Approve Intelligent Selections")
        box.setStyleSheet(f"font-size: {txt_size}px;")
        boxlay=QFormLayout()
        box.setLayout(boxlay)

        #loop this
        self.IntelligentBoxes=[]
        for i in range(0,len(DNumber)):
            rows=0 #start at 0 always
            ntree= QTreeWidget()
            tmp=DNumber[i]
            tmp0=SubDivisionNumber[i]
            tmp1=SubDivisionDescription[i]
            
            #Division (1st lvl) checkbox
            DivisionNumber=QTreeWidgetItem(ntree)
            DivisionNumber.setText(0,f"Division {tmp} - {tmp0} - {tmp1}")
            DivisionNumber.setFlags(DivisionNumber.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            DivisionNumber.setCheckState(0, Qt.Checked)
            self.IntelligentBoxes.append(DivisionNumber)
                
            rows+=1
            ntree.header().hide()

            #Details (2nd lvl) checkbox
            for det in Details[i]:
                tmp=det
                rows+=1
                det=QTreeWidgetItem(DivisionNumber)
                det.setText(0,f"{tmp}")
                det.setFlags(det.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                det.setCheckState(0, Qt.Checked)
                self.IntelligentBoxes.append(det)
            
            #tree.setExpanded(QModelIndex,SubDivisionNumber)
            ntree.expandAll()
            ntree.setUniformRowHeights(True)

            height = int(rows*22*scaleFactor) 
                
            ntree.setFixedHeight(height)
            boxlay.addWidget(ntree)
            
        scroll=QScrollArea()
        scroll.setWidget(box)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll,2,1,1,3)
        
        dialog.setLayout(layout)
        dialog.showMaximized()
        dialog.exec_()
        
    def approvechk(self):
        print("go")
        boxes=self.IntelligentBoxes
        
        global MLapproved_chkboxs
        MLapproved_chkboxs=[]
        t=[]
        divNumRegexr=r"Division [0-9]{2} [-]{1} [0-9]{2}[ ]{1}[0-9]{2}[ ]{1}[0-9]{2}[.]{0,1}[0-9]{0,2}\b"
        for i in range(0,len(boxes)): 
            if boxes[i].checkState(0) == 1 or boxes[i].checkState(0) == 2:
                bx=boxes[i].text(0)
                if re.match(divNumRegexr,bx) and len(t)!=0:
                    MLapproved_chkboxs.append(t)
                    t=[]
                    t.append(bx)
                else:
                    t.append(bx)
                    
            elif i == len(boxes)-1:
                MLapproved_chkboxs.append(t)
                
        # print(MLapproved_chkboxs)
        self.close()
        
class BuildDB:    
    def __init__(self):
        dir = DivDirectory
        Files=self.AllFilesInDIR(dir)  
        #print(Files)
        self.StripNSave2DB(Files)
        self.AddNumFilesStats(dir,Files)
    
    def IntoDB(self,DivisionNumber,SubdivisionNumber,SubDivisionDescription,SectionDetails): 
        conn = sqlite3.connect(DB_Path)
        cur=conn.cursor()
        cur.execute(f'''INSERT INTO Components (Division_Number, Subdivision_ID_Number, Subdivision_Description, Details)
                    VALUES('{DivisionNumber}','{SubdivisionNumber}','{SubDivisionDescription}','{SectionDetails.replace("'", "''")}')''')
        conn.commit()
        conn.close()
        return

    def AllFilesInDIR(self,path):
        All = []
        for file in os.listdir(path):
            d=os.path.join(path,file).replace("/","\\")
            #print('d',d,'file',file)
            if os.path.isdir(d):
                #print('yes')
                if file.startswith("DIVISION"):
                    #print('file',file)
                    for i in os.listdir(d):
                        #we are into each sub folder now
                        pth=os.path.join(d, i).replace("/","\\")
                        #print ('path',pth)
                        All.append(pth) #appends pathname of each word file to All
        
        convert = lambda text: int(text) if text.isdigit() else text 
        alphanum_key = lambda key: [ convert(c) for c in re.split('([0-9]+)', key) ]
    
        sortedAll=sorted(All, key = alphanum_key)
        return sortedAll

    def AddNumFilesStats(self,dir,Files):
        #get num files in folder
        NumSpecFiles=len(Files) #numeber of spec documents in the raw spec file folder
        #get into Division - ## format
        subdivIDexpr="[0-9]{2}[ ]{1}[0-9]{2}[ ]{1}[0-9]{2}[.]{0,1}[0-9]{0,2}"
        SpecFileNames=[]
        errorOnRegex=''
        for f in Files: 
            try: 
                SpecFileNames.append(re.findall(subdivIDexpr,f)[0])
            except:
                f.replace("/","\\")
                errorOnRegex+=f
                errorOnRegex+='\n'
               
        #get num files in DB
        conn = sqlite3.connect(DB_Path) 
        cur=conn.cursor()
        unique=cur.execute('''SELECT DISTINCT(Subdivision_ID_Number)
                               FROM Components
                               ''')
        UniqueSubdivisions=[]
        for i in unique:
            UniqueSubdivisions.append(str(i[0]))
        NumberOfUniqueSubdivisions=len(UniqueSubdivisions) #number of unique entries in the DB
        
        ##get differences
        MissingFromDB=sorted(list(set(SpecFileNames)-set(UniqueSubdivisions))) #gets list of items not in DB, but are in the raw spec list
        MissingFromRawSpecs=sorted(list(set(UniqueSubdivisions)-set(SpecFileNames))) #gets list of items in DB, but are not in raw spec list
        if len(MissingFromDB)==0: 
            MissingFromDB="Zero"
        if len(MissingFromRawSpecs)==0: 
            MissingFromRawSpecs="Zero"
        
        #add everything to DB - issues section
        cur.execute(f''' INSERT INTO Stats (NumFilesInDB, NumFilesInRawSpecs) 
                    VALUES ({NumSpecFiles},{NumberOfUniqueSubdivisions})
                    ''')
        conn.commit()
        
        cur.execute(rf'''INSERT INTO FixManually (MissingFromDB, MissingFromRawSpecs, errorOnRegex, NumMissedByRegex)
                    VALUES ('{MissingFromDB}', '{MissingFromRawSpecs}', '{errorOnRegex}', '{NumSpecFiles-len(SpecFileNames)}')
                    ''')
        conn.commit()

        
    def StripNSave2DB(self,Files):
        AllSubsections = []
        global cmpltmanually
        #strips word files & saves components in DB!
        word=win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        manual=[]
        
        SafeFiles=[]
        for i in Files: 
            if "~" not in i: 
                SafeFiles.append(i)

        for file in SafeFiles:
            try:
                #Get Div Number
                directory=(os.path.basename(os.path.dirname(file)))
                TwoDigNum_regex="([0-9]{1,2})" ##2 digit number
                DivNo = re.findall(TwoDigNum_regex,directory)
                DivisionNumber=int(DivNo[0])
                #get subdivision number
                sub_no=os.path.basename(file)
                SubDivisionNumberRegex="[0-9]{2}[ ]{1}[0-9]{2}[ ]{1}[0-9]{2}[.]{0,1}[0-9]{0,2}"
                SubdivisionN0 = re.findall(SubDivisionNumberRegex,sub_no)
                print('Subdivision #', SubdivisionN0)
                SubdivisionNumber=str(SubdivisionN0[0])
                print(SubdivisionNumber)
                #Get Subdivision description
                Sub_desc=re.sub(SubDivisionNumberRegex,"", sub_no)
                SubDivisionDescription=Sub_desc[1:]
                #print("desc",SubDivisionDescription)
            
                try:
                    src = word.Documents.Open(file)
                except:
                    print("Couldn't open file", file)
            
                src.Activate()
                doc = word.ActiveDocument
                
                end =(len(doc.ListParagraphs)+1)
                i=0
                startrng =0
                endrng = 0
                while i < end:
                    i+=1
                    pgr=doc.Paragraphs(i)
                    if str(pgr).isupper() == True:
                        d=str(pgr)
                        if "END OF SECTION" in str(pgr):
                            endrng="done"
                        SectionDescription=str(pgr) #this doesnt work!
                        #print('SectionDescription',SectionDescription)
                        if startrng == 0:
                            startrng = i
                            #print('start',startrng)
                        elif startrng != 0:
                            endrng = i
                            #print('startrng',startrng)
                            #print('endrng',endrng)
                            
                        if startrng !=0 and endrng !=0:
                            rng=doc.Range(Start:=doc.Paragraphs(startrng).Range.Start, End:=doc.Paragraphs(endrng-1).Range.End)
                            rng.Select
                            rng_txt=str(rng)
                            SectionDetails=rng_txt
                            
                            #print("Description",SectionDetails)
                            
                            self.IntoDB(DivisionNumber,SubdivisionNumber,SubDivisionDescription,SectionDetails)
                        
                            #resets startrng and endrng values to work on next iteration
                            startrng = endrng #must be second last item in this IF statement
                            endrng = 0 #must be last item in this IF statement

            except: 
                print("SOMETHING ISNT WORKING...PUT THIS FILE IN THE DB MANUALLY", file)
                manual.append(file)
                print("Currently",len(manual),"files to fix")
                
            doc.Close() 
        
        print('FIX', len(manual), 'FILES MANUALLY:',manual)
        return None 


class Rebuilder: 
    def __init__(self):    
        self.AddSection(send)
    
    def clearGenPy(self):
        import os
        import shutil
        import win32com

        #Run this script if you get an error along the lines of gen_py is not working.
        #This deletes the whole gen_py temp folder
        #Folder repopulates with the important stuff when running new win32com using script

        path =  win32com.__gen_path__
        folder = os.path.dirname(path)
        print(folder)

        if os.path.exists(folder):
            shutil.rmtree(folder)
            print("CLEARED GENPY FOLDER")
            
        else:
            print('CLEAR GENPY FOLDER NOT FOUND')
    
    def AddSection(self,subdiv_list):
        DivisionNumber_andSubDivisionNumber_andDetails = subdiv_list[0]
        Details = subdiv_list[1]
        doc=word.ActiveDocument
        #word.Visible = True ##good for troubleshooting formatting stuff
        
        #use this process to get range of what youre entering and apply corresponding style to it only.
        EndRng = doc.Paragraphs.Count +1
        word.ActiveDocument.Content.InsertAfter(Text:=f"{DivisionNumber_andSubDivisionNumber_andDetails}\r")
        NewEndRng=word.ActiveDocument.Paragraphs.Count +1
        rng=word.ActiveDocument.Range(Start:=word.ActiveDocument.Paragraphs(EndRng-1).Range.Start, End:=word.ActiveDocument.Paragraphs(NewEndRng-2).Range.End)
        rng.Style = 'Division Title'
        
        for det in Details: 
            #word.Selection.Text = "\r"
            #word.ActiveDocument.Content.InsertBefore(chr(11))
            n=det.find("\r")
            if n != -1:
                if det.find("GENERAL") != -1 or det.find("PRODUCTS") != -1 or det.find("EXECUTION") != -1: #if General products of execution exists!
                    #PASTE AS HEADING 1
                    EndRng = doc.Paragraphs.Count +1
                    doc.Content.InsertAfter(Text:=det)
                    NewEndRng=word.ActiveDocument.Paragraphs.Count +1
                    rng=word.ActiveDocument.Range(Start:=word.ActiveDocument.Paragraphs(EndRng-1).Range.Start, End:=word.ActiveDocument.Paragraphs(NewEndRng-2).Range.End)
                    rng.Style = 'Heading 1'
                    rng.Select
                    activelist=word.Selection.Range.ListFormat
                    
                    try:
                        activelist.ContinuePreviousList= False
                    except: 
                        # print("Section was probably the first one!")
                        None
                    
                else:
                    #paste the all caps stuff (probably a sub section title (heading2))
                    Allcaps=det[0:n]
                    remainder=det[n+1:len(det)]
                    EndRng = doc.Paragraphs.Count +1
                    word.ActiveDocument.Content.InsertAfter(Text:=f"{Allcaps}\r")
                    NewEndRng=word.ActiveDocument.Paragraphs.Count +1
                    rng=word.ActiveDocument.Range(Start:=word.ActiveDocument.Paragraphs(EndRng-1).Range.Start, End:=word.ActiveDocument.Paragraphs(NewEndRng-2).Range.End)
                    rng.Style = 'Heading 2'
                    
                    #now paste remainder (heading 3)
                    EndRng = doc.Paragraphs.Count +1
                    word.ActiveDocument.Content.InsertAfter(Text:=remainder)
                    NewEndRng=word.ActiveDocument.Paragraphs.Count +1
                    rng=word.ActiveDocument.Range(Start:=word.ActiveDocument.Paragraphs(EndRng-1).Range.Start, End:=word.ActiveDocument.Paragraphs(NewEndRng-2).Range.End)
                    rng.Style = 'Heading 3'
    
            else: 
                ##no return char
                word.ActiveDocument.Content.InsertAfter(Text:=remainder)
        
                
        word.ActiveDocument.Content.InsertAfter("\r")     
        word.ActiveDocument.Content.InsertAfter(Text:="END OF SECTION\r\r")
        word.ActiveDocument.Save

def Final():
        word.ActiveDocument.Save()
        word.ActiveDocument.Close()
        print("Successfully saved output file!")

def preamble():
        #opens word once here
        global word
        word=win32com.client.gencache.EnsureDispatch("Word.Application")
        #word=win32com.client.Dispatch("Word.Application")
        
        try:
            #Opens template and runs template specific macro: TemplateFirstOpen
            #then saves template for this job as JobSpecificTemplate
            word.Visible = True #a macro pop up occurs that we need to see!
            SpecBuildingTemplate = os.path.join(cwd,"SpecBuildingTemplateR2.docm")
            word.Documents.Open(SpecBuildingTemplate)
            word.Application.Run("ThisDocument.TemplateFirstOpen")
            word.Visible = False
            word.ActiveDocument.Save    
                
        except: 
            print("Could not complete first save!")
      
def saveWnewname(pth,outputname="OutputSpecDraftDocument"):
    word=win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    
    try:
        SpecBuildingTemplate = os.path.join(cwd,"SpecBuildingTemplateR2.docm")
        word.Documents.Open(SpecBuildingTemplate)
        word.ActiveDocument.Save    
            
    except: 
        print("Could not complete first save!")
    
    
    try: 
        output=word.ActiveDocument
        # output.SaveAs2(pth,"tmp.docm")
        # output.VBProject.VBComponents.Remove(output.VBProject.VBComponents.Item(1))
        opath= os.path.join(pth,f"{outputname}.docx") ##saveas docm keeps macro in it
        output.SaveAs2(opath, FileFormat=16)
        
    except:
        print("Could not open and save brand new file!")
        word.ActiveDocument.Save
        #word.ActiveDocument.Close()

def initiate_db():
    conn = sqlite3.connect(DB_Path)
    cur=conn.cursor()
    cur.execute(''' CREATE TABLE Components (
        ID INTEGER PRIMARY KEY AUTOINCREMENT,
        Division_Number TEXT,
        Subdivision_ID_Number TEXT,
        Subdivision_Description TEXT,
        Details TEXT)
        ''')
    conn.commit()
    
    cur.execute('''CREATE TABLE FixManually (
        MissingFromDB TEXT,
        MissingFromRawSpecs TEXT,
        errorOnRegex TEXT,
        NumMissedByRegex TEXT)
        ''')
    
    conn.commit()
    
    cur.execute('''CREATE TABLE Stats (
        UseCounter INTEGER,
        NumFilesInDB TEXT,
        NumFilesInRawSpecs TEXT)
        ''')
    
    conn.commit()
    conn.close()
       

def paywall():
    er=QApplication(sys.argv)
    er.setWindowIcon(QIcon(os.path.join(os.path.dirname(os.path.realpath(__file__)),"CoreEngrSmallLogo.ico")))
    msg = QMessageBox()
    msg.setWindowTitle("Subscription Expired") 
    msg.setText("This is a pay wall. \n logan@palmers.ca")
    msg.setIcon(QMessageBox.Critical)
    msg.exec_() 
    exit()
    

if __name__ == '__main__':
    ################################################################################################################## checks if paywall for date expiry needs to be executed
    today = date.today()

    if today <= expiry_date:
        None
    else:
        paywall() ##calls paywall app to open
    ##################################################################################################################
    
    global start_time
    start_time = time.time()
    
    cwd=os.path.dirname(os.path.realpath(__file__)) ##actual folder the python file is running in!
    DB_Path = os.path.join(cwd,"Components.db")
    #this makes sure db exists, and makes if needed!
    conn = sqlite3.connect(DB_Path)
    cur=conn.cursor()
    exi=cur.execute(''' SELECT 1 FROM sqlite_master WHERE type='table' AND name='Components' ''')
    exists=False
    
    for i in exi:
        print(str(i[0]))
        exists=True
        
    if exists!=True: 
        print("DATABASE DOES NOT EXIST -- INITIALIZING NOW")
        initiate_db()
    
    ################################################################################################################## checks if paywall for use counter needs to be executed
    elif exists==True:
        cur.execute('SELECT UseCounter FROM Stats')
        n=cur.fetchone()[0]
        if n== None:
            n=0
        if n>MaxNumUses:
            paywall()
        new=n+1
        # cur.execute(f'''INSERT INTO Stats UseCounter = {new} ''')
        cur.execute("UPDATE Stats SET UseCounter = ? WHERE RowID = ?", (new, 1))
        conn.commit()
    ##################################################################################################################
      
    app = QApplication(sys.argv)
    window = App()
    sys.exit(app.exec_())