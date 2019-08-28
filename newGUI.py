"""
ACG INDUSTRY PLAN GENERATOR MAIN FILE


"""
from PyQt5 import QtWidgets as QtW
from PyQt5 import QtGui as QtG
from PyQt5 import QtCore
import BenchmarkingProj as BNFunc
import os
import sys
import json

#Varaibles needed for global use
#location of the PDF used in the input
pdfLoc = ''
#path to the desired output directory
outputPath = ''
#the name you want to give the generated file
fileName = ''

#lists and dictionaries used when a user selects questions
selectedQuestions = []
orgSelectedQuestions = {}
selectedTemplateQuestions = []

def resource_path(relative_path):
    """returns the correct path of a file varaible needed.

    :param relative_path: the name of the path you want to concatenate onto the current directory.
    :return: returns correct path to the file
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def setup():
    """pulls necessary info the GUI needs to populate fields.

    grabs data from settings file and dumps them into the settings list. then assigns data to vars.
    :return:
    """
    global pdfLoc, fileName, outputPath
    settings = []
    settingsFile = open(resource_path('sources\settings'), 'r', encoding="utf8")
    for line in settingsFile.readlines():
        line = line.replace('\n', "")
        settings.append(line)

    pdfLoc = settings[0]
    outputPath = settings[1]
    fileName = settings[2]

    settingsFile.close()

class mainGUI(QtW.QWidget):
    """Main GUI Window for Industry Plan Generator

    ...
    Methods
    -------
    questiongrabber(): Pulls categories and questions from text files.
    addQ(): Adds questions to textbox and needed lists
    deleteQ(): Deletes selected questions from textboxes and needed lists
    writeTextBox(): Writes selected questions into selectedQsTextBox
    start(): Starts the process of generating the Excel file
    getOptions(): Gets option associated with each selected question
    initPDCWiz(): Initializes the PDC Wizard so it doesn't close right away
    __init__(): init function for the main GUI
    initGen(): Initialization function for the generator tab
    saveSettings(): Saves changed settings from the settings tab.
    validateInputFile(): Validates the input file chosen
    validateName(): Validates the name of the excel file given
    openInputFile(): Opens file dialog for input file selection.
    openOutputFile(): Opens file dialog for output directory selection.
    initSettings(): Initializes the settings tab
    initDirections(): Initializes the directions tab
    """

    def questiongrabber(self):
        """Pulls categories and questions from text files.

        it then matches categories with questions and puts them within a dictionary category.
        :return: returns categories dict with matching questions
        """
        categories = {}
        questions = []
        cateList = []

        categoryFile = open(resource_path('sources\catagories'), 'r', encoding="utf8")
        for line in categoryFile.readlines():
            line = line.replace('\n', "")
            categories[line] = []
        categoryFile.close()

        questionFile = open(resource_path('sources\questions'), 'r', encoding="utf8")
        for line in questionFile.readlines():
            line = line.replace('\n', "")
            #if a divider is detected then the next set of questions starts else the line gets added to questions
            if "|" == line:
                questions.append(cateList)
                cateList = []
            else:
                cateList.append(line)

        questionFile.close()

        #each set of questions gets assigned to its correct category through indexing down the list and assigning them
        QIndex = 0
        for k in categories:
            try:
                categories[k] = questions[QIndex]
                QIndex += 1
            except:
                pass

        return categories

    def addQ(self):
        """Adds questions to textbox and needed lists

        loops through each item selected and sorts them based on if they were already added or not.
        :return:
        """
        catagories = self.questiongrabber()
        for item in self.qList.selectedItems():
            if item.text() in selectedQuestions:
                pass
            else:
                for c in catagories:
                    if item.text() in catagories[c]:
                        for cata in orgSelectedQuestions:
                            if cata == c:
                                orgSelectedQuestions[cata].append(item.text())

                selectedQuestions.append(item.text())

        self.writeTextBox()

    def deleteQ(self):
        """Deletes selected questions from textboxes and needed lists

        sorts through selected questions and removes matching questions from selectedQuestions & orgSelectedQuestions
        :return:
        """
        for item in self.selectedQsList.selectedItems():
            try:
                item = item.text().split("•")[1]
                if str(item) in selectedQuestions:
                    selectedQuestions.remove(item)
                    for cate in orgSelectedQuestions:
                        if item in orgSelectedQuestions[cate]:
                            orgSelectedQuestions[cate].remove(item)
            except:
                pass
        self.writeTextBox()

    def writeTextBox(self):
        """Writes selected questions into selectedQsTextBox
        :return:
        """
        self.selectedQsList.clear()
        for k,v in orgSelectedQuestions.items():
            if v == []:
                pass
            else:
                cate = QtW.QListWidgetItem()
                cate.setText(k)
                cate.setFlags(QtCore.Qt.NoItemFlags)
                self.selectedQsList.addItem(cate)
                for i in v:
                    self.selectedQsList.addItem('    •' + i)

    def start(self):
        """Starts the process of generating the Excel file

        :return:
        """
        global selectedTemplateQuestions
        if self.PDCradioButton.isChecked() == True:
            BNFunc.clearResources()
            self.pdcWiz = self.initPDCWiz(self.finishedSuccessLabel)
            self.pdcWiz.show()
        else:
            BNFunc.clearResources()
            workbook = BNFunc.workBook(fileName, outputPath, pdfLoc, selectedTemplateQuestions, False, 'All')
            workbook.createTitleSheet()
            workbook.createDataSheet()
            workbook.createDisclosureSheet()
            done = workbook.close()
            if done == True:
                self.finishedSuccessLabel.setText("Generated " + fileName)
            elif done == False:
                self.doneSaveMessageBox = QtW.QMessageBox()
                self.doneSaveMessageBox.setStyleSheet(self.stylesheet)
                self.doneSaveMessageBox.setWindowTitle("Save Error")
                errorMessage = "Error when saving. You most likely have a file open with \n the same name in the same location."
                self.doneSaveMessageBox.setText(errorMessage)
                self.doneSaveMessageBox.show()

    def initPDCWiz(self, ssLabel):
        """ Initializes the PDC Wizard so it doesn't close right away

        :param values: values returned from BNF.MatchQuestions
        :param options: dictionary that contains questions and options associated with each question
        :param ssLabel: success label so when the process is finished the text can be changed
        :return: PDCWiz: class instance for further method calls
        """

        pdcWiz = planDesignComparisonWiz(ssLabel)
        return pdcWiz


    def __init__(self):
        """init function for the main GUI

        """
        super(mainGUI, self).__init__()

        stylesheet_path = resource_path("sources\stylesheet.qss")
        self.stylesheet = open(stylesheet_path, "r").read()
        self.setStyleSheet(self.stylesheet)

        self.questionLists = []

        self.layout = QtW.QVBoxLayout(self)
        self.tabs = QtW.QTabWidget()

        #tabs
        self.questionsTab = QtW.QWidget()
        self.createTab = QtW.QWidget()
        self.settingsTab = QtW.QWidget()
        self.directionsTab = QtW.QWidget()

        #init the tabs and there layouts
        self.initGen()
        self.initSettings()
        self.initCreateTab()
        self.initDirections()

        #adds tabs to the main layout
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)

        #adds tabs to the tab bar
        self.tabs.addTab(self.settingsTab, "Settings")
        self.tabs.addTab(self.createTab, "Create")
        self.tabs.addTab(self.questionsTab, "Templates")
        self.tabs.addTab(self.directionsTab, "Directions")


        self.setFixedHeight(400)
        self.setFixedWidth(1100)
        self.setWindowTitle("ACG Benchmarking Generator")
        self.Icon = QtG.QIcon()
        self.Icon.addFile(resource_path('sources/ACGiGen.ico'))
        self.setWindowIcon(self.Icon)

    def initGen(self):
        """ Initialization function for the generator tab

        :return:
        """
        catagories = self.questiongrabber()
        actions = []
        self.mainGrid = QtW.QGridLayout()
        self.toolbarLayout = QtW.QHBoxLayout()
        self.textboxLayout = QtW.QHBoxLayout()
        self.templatesLayout = QtW.QHBoxLayout()
        self.buttonLayout = QtW.QHBoxLayout()
        self.successLLayout = QtW.QVBoxLayout()

        self.toolbar = QtW.QToolBar()

        self.catagoryLabel =QtW.QLabel('')
        self.catagoryLabel.setStyleSheet("font-weight: bold;")

        self.qList = QtW.QListWidget()
        self.qList.setSelectionMode(QtW.QListWidget.MultiSelection)

        self.selectedQsList = QtW.QListWidget()
        self.selectedQsList.setSelectionMode(QtW.QListWidget.MultiSelection)

        for c in catagories:
            orgSelectedQuestions[c] = []
            action = QtW.QAction(c, self)
            action.triggered.connect(lambda checked, c=c: changeTextBox(c))
            self.toolbar.addAction(action)
            actions.append(action)

        def changeTextBox(c):
            """changes the category label & questions within the select questions textbox
            :param c: selected category
            :return:
            """
            self.catagoryLabel.setText(c)
            self.qList.clear()
            for q in catagories[c]:
                self.qList.addItem(q)

        self.templateName = QtW.QLineEdit()
        self.templateName.setText('Template Name')

        self.saveTempButton = QtW.QPushButton('Create Template')
        self.saveTempButton.clicked.connect(self.saveTemplate)

        self.addButton = QtW.QPushButton('Add Field')
        self.addButton.clicked.connect(self.addQ)

        self.deleteButton = QtW.QPushButton('Delete Field')
        self.deleteButton.clicked.connect(self.deleteQ)

        self.templateSuccessLabel = QtW.QLabel('')

        self.templatesLayout.addWidget(self.templateName)
        self.templatesLayout.addWidget(self.saveTempButton)

        self.buttonLayout.addWidget(self.addButton)
        self.buttonLayout.addWidget(self.deleteButton)

        self.successLLayout.addWidget(self.templateSuccessLabel)

        self.toolbarLayout.addWidget(self.toolbar)
        self.textboxLayout.addWidget(self.qList)
        self.textboxLayout.addWidget(self.selectedQsList)

        self.mainGrid.addItem(self.toolbarLayout, 0, 0)
        self.mainGrid.addWidget(self.catagoryLabel)
        self.mainGrid.addItem(self.textboxLayout, 2, 0)
        self.mainGrid.addItem(self.templatesLayout, 3, 0)
        self.mainGrid.addItem(self.buttonLayout, 4, 0)
        self.mainGrid.addItem(self.successLLayout, 5, 0)


        self.questionsTab.setLayout(self.mainGrid)

    def saveSettings(self):
        """Saves changed settings from the settings tab.

        overwrites settings page with the inputs aquired from the GUI.

        :return:
        """
        valid = self.validateName(self.nameInput)
        if valid == True:
            settingsFile = open(resource_path('sources\settings'), 'w', encoding="utf8")
            settingsFile.writelines(self.linkInput.text() + "\n")
            settingsFile.writelines(self.outputPathInput.text() + "\n")
            settingsFile.writelines(self.nameInput.text() + "\n")
            settingsFile.close()
            setup()
            self.saveSuccessLabel.setText("Success")

    def validateInputFile(self, name):
        """ Validates the input file chosen

        only Windows 97-2003 excel files are allowed due to compatibility issues.

        :param name: the excel file path choosen
        :return: returns boolean valid if the name given is valid or not
        """
        errorMessage = ''
        valid = True
        name = name[0]
        self.inMessageBox = QtW.QMessageBox()
        self.inMessageBox.setStyleSheet(self.stylesheet)
        self.inMessageBox.setWindowTitle("Save Error")
        if ".xls" not in name and name != '':
            errorMessage += "needs to be a .xls file to work. please select a valid file and try again. \n"
            valid = False
        if valid == False:
            self.inMessageBox.setText(errorMessage)
            self.inMessageBox.show()
        return valid

    def validateName(self, name):
        """ Validates the name of the excel file given.

        checks for invalid symbols or if the name is blank. if those checks are false then the name is valid.
        :param name: name of the file given
        :return: returns boolean valid which is either true or false
        """

        invalidSymbols = ['#', '%', '*', '(', ')', '-', '_', '=', '[', ']', '{', '}', '/', '\\', '|', '@', '`', '~', '!', '?']
        errorMessage = ''
        valid = True
        self.inMessageBox = QtW.QMessageBox()
        self.inMessageBox.setStyleSheet(self.stylesheet)
        self.inMessageBox.setWindowTitle("Save Error")
        name = name.text()

        for i in invalidSymbols:
            if i in name:
                errorMessage += "invalid character in name input: " + i + "\n"
                valid = False
        if name == '':
            errorMessage += "invalid name, please enter a name. \n"
            valid = False
        if valid == False:
            self.inMessageBox.setText(errorMessage)
            self.inMessageBox.show()
        return valid


    def openInputFile(self):
        """ Opens file dialog for input file selection.

        saves the name of the file and runs a validation check on the selected file. if it passes then the input text is set and settings are saved.
        :return: if returned then it cancels the function.
        """
        name = QtW.QFileDialog.getOpenFileName(self, "Input File", "G:\\1) RETIREMENT\\Surveys & Benchmarking Data\\2019\\Converted")
        valid = self.validateInputFile(name)
        if name[0] == '':
            return
        if valid == True:
            self.linkInput.setText(name[0])

    def openOutputFile(self):
        """ Opens file dialog for output directory selection.

        if no directory is selected the function cancles with the return call. else the input text is updated and the save settings function is called.
        :return:
        """
        direct = QtW.QFileDialog.getExistingDirectory(self, "Output File", 'G:\\1) RETIREMENT\\Surveys & Benchmarking Data\\Compiled Reports')
        if direct == '':
            return
        else:
            self.outputPathInput.setText(direct)

    def clearSSLabel(self):
        self.saveSuccessLabel.setText("")

    def initCreateTab(self):
        self.createTabGrid = QtW.QGridLayout()

        self.currentTempGB = QtW.QGroupBox("Current Template")
        self.currentTempGB.setStyleSheet("max-height: 80px; background-color: #f4ce1f; border-style: solid; border-color: #deaf25; border-width: 2px;")

        self.createReportTabLayout = QtW.QHBoxLayout()
        self.finishedSuccessLayout = QtW.QVBoxLayout()
        self.selectedTemplateLayout = QtW.QHBoxLayout()
        self.createButtonLayout = QtW.QHBoxLayout()

        self.templateSelection = QtW.QComboBox()
        self.selectTempButton = QtW.QPushButton('Select Template')
        self.selectTempButton.clicked.connect(self.selectTemp)

        self.deleteTempButton = QtW.QPushButton('Delete Template')
        self.deleteTempButton.clicked.connect(self.deleteTemplate)

        self.startButton = QtW.QPushButton("Create Report")
        self.startButton.clicked.connect(self.start)

        self.selectedTemplateLabel = QtW.QLabel("Current Template")
        self.selectedTemplateLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.currentTemplateLabel = QtW.QLabel("")
        self.currentTemplateLabel.setAlignment(QtCore.Qt.AlignCenter)

        self.finishedSuccessLabel = QtW.QLabel("")
        self.finishedSuccessLabel.setFixedWidth(1100)
        self.finishedSuccessLabel.setFixedHeight(20)
        self.finishedSuccessLabel.setStyleSheet("padding: 0px;")
        self.finishedSuccessLabel.setAlignment(QtCore.Qt.AlignCenter)

        self.createReportTabLayout.addWidget(self.templateSelection)
        self.createReportTabLayout.addWidget(self.selectTempButton)
        self.createReportTabLayout.addWidget(self.deleteTempButton)

        self.loadTemplates()

        self.selectedTemplateLayout.addWidget(self.selectedTemplateLabel)
        self.selectedTemplateLayout.addWidget(self.currentTemplateLabel)

        self.currentTempGB.setLayout(self.selectedTemplateLayout)

        self.createButtonLayout.addWidget(self.startButton)

        self.finishedSuccessLayout.addWidget(self.finishedSuccessLabel)

        self.createTabGrid.addWidget(self.currentTempGB, 0, 0)
        self.createTabGrid.addItem(self.createReportTabLayout, 1, 0)
        self.createTabGrid.addItem(self.createButtonLayout, 2, 0)
        self.createTabGrid.addItem(self.finishedSuccessLayout, 3, 0)


        self.createTab.setLayout(self.createTabGrid)

    def selectTemp(self):
        global selectedTemplateQuestions
        selectedTemplateQuestions = []
        selectedTemp = self.templateSelection.currentText()
        self.currentTemplateLabel.setText(selectedTemp)
        with open(resource_path("sources/templates"), 'r') as f:
            data = json.load(f)
            for element in data:
                if element == selectedTemp:
                    for q in data[element]:
                        selectedTemplateQuestions.append(q)

    def saveTemplate(self):
        with open(resource_path("sources/templates"), 'r+') as f:
            data = json.load(f)
            tempNames = data.keys()
            if self.templateName.text() in tempNames:
                errorMessage = 'Already a template with the same name. \nPlease use a different name and try again.'
                self.tempMessageBox = QtW.QMessageBox()
                self.tempMessageBox.setStyleSheet(self.stylesheet)
                self.tempMessageBox.setWindowTitle("Save Error")
                self.tempMessageBox.setText(errorMessage)
                self.tempMessageBox.show()
            else:
                data[self.templateName.text()] = []
                for k in selectedQuestions:
                    data[self.templateName.text()].append(k)
                f.seek(0)
                json.dump(data, f, indent=4)
                f.truncate()

        self.templateSuccessLabel.setText('Created ' + str(self.templateName.text()))
        self.loadTemplates()

    def loadTemplates(self):

        self.templateSelection.clear()
        with open(resource_path("sources/templates"), 'r') as f:
            data = json.load(f)
            templateNames = data.keys()
            for k in templateNames:
                self.templateSelection.addItem(k)


    def deleteTemplate(self):
        delete_temp = self.templateSelection.currentText()
        temp_dict = {}
        with open(resource_path('sources/templates'), 'r') as data_file:
            data = json.load(data_file)

        for element in data:
            if element == delete_temp:
                pass
            else:
                temp_dict[element] = data[element]

        with open(resource_path('sources/templates'), 'w') as data_file:
            json.dump(temp_dict, data_file)

        self.loadTemplates()

    def initSettings(self):
        """ Initializes the settings tab

        :return:
        """
        self.settingsGrid = QtW.QGridLayout()
        self.settingsMainLayout = QtW.QVBoxLayout()


        self.ifLayout = QtW.QHBoxLayout()
        self.ofLayout = QtW.QHBoxLayout()
        self.nameLayout = QtW.QHBoxLayout()
        self.radioButtonLayout = QtW.QHBoxLayout()
        self.settingsLayout = QtW.QVBoxLayout()

        self.LUGB = QtW.QGroupBox("Settings")

        self.linkINLabel = QtW.QLabel("Excel File Input ")
        self.linkINLabel.setToolTip(" by clicking on 'Open File' a directory will pop up allowing you to select a desired file to pull data from")
        self.ifLayout.addWidget(self.linkINLabel)

        self.linkOPLabel = QtW.QLabel("Directory Output ")
        self.linkOPLabel.setToolTip("the desired directory that the program outputs too")
        self.ofLayout.addWidget(self.linkOPLabel)

        self.nameLabel = QtW.QLabel("File Name ")
        self.nameLabel.setToolTip("The name of the generated file")
        self.nameLayout.addWidget(self.nameLabel)

        self.linkInput = QtW.QLineEdit()
        self.linkInput.setReadOnly(True)
        self.linkInput.setText(pdfLoc)
        self.linkInput.textChanged.connect(self.clearSSLabel)
        self.ifLayout.addWidget(self.linkInput)

        self.outputPathInput = QtW.QLineEdit()
        self.outputPathInput.setReadOnly(True)
        self.outputPathInput.setText(outputPath)
        self.outputPathInput.textChanged.connect(self.clearSSLabel)
        self.ofLayout.addWidget(self.outputPathInput)

        self.nameInput = QtW.QLineEdit()
        self.nameInput.setText(fileName)
        self.nameInput.textChanged.connect(self.clearSSLabel)
        self.nameLayout.addWidget(self.nameInput)


        self.openFileButton = QtW.QPushButton("Open File")
        self.openFileButton.clicked.connect(self.openInputFile)
        self.ifLayout.addWidget(self.openFileButton)

        self.openOutputButton = QtW.QPushButton("Open Directory")
        self.openOutputButton.clicked.connect(self.openOutputFile)
        self.ofLayout.addWidget(self.openOutputButton)

        self.saveButton = QtW.QPushButton("Save")
        self.saveButton.setToolTip("Saves your changes")
        self.saveButton.clicked.connect(self.saveSettings)

        self.PDCLabel = QtW.QLabel("Plan Design Comparison")
        self.PDCLabel.setToolTip("if you would like to generate a Plan Design Comparison leave this checked.")
        self.PDCradioButton = QtW.QRadioButton()
        self.PDCradioButton.setChecked(True)

        self.saveSuccessLabel = QtW.QLabel("")

        self.radioButtonLayout.addWidget(self.PDCLabel)
        self.radioButtonLayout.addWidget(self.PDCradioButton)
        self.radioButtonLayout.addWidget(self.saveSuccessLabel)
        self.radioButtonLayout.addWidget(self.saveButton)

        self.settingsLayout.addItem(self.ifLayout)
        self.settingsLayout.addItem(self.ofLayout)
        self.settingsLayout.addItem(self.nameLayout)

        self.LUGB.setLayout(self.settingsLayout)
        self.settingsMainLayout.addWidget(self.LUGB)
        self.settingsMainLayout.addItem(self.radioButtonLayout)

        self.settingsGrid.addItem(self.settingsMainLayout, 0, 0)

        self.settingsTab.setLayout(self.settingsGrid)

    def initDirections(self):
        """ Initializes the directions tab

        :return:
        """
        self.directionsGrid = QtW.QGridLayout()
        self.mainGBDirections = QtW.QGroupBox("Directions")
        self.mainDirectionsLayout = QtW.QVBoxLayout()

        setBoxText = """
        <!DOCTYPE html>
        <html>
        <head>
        </head>
        <body>
        <h1> The Settings Tab </h1>
        <p> <b>Excel File Input: </b> by clicking on "Open File" a directory will pop up allowing you to select a desired file to pull data from.</p>
        <p> <b>Directory Output: </b> the desired directory that outputs the finished Excel file.</p>
        <p> <b>File Name: </b> the given name to the finished Excel file. *NOTE: if a file has the same name in the given directory it will overwrite that file.</p>
        <p> <b>Plan Design Comparison & Button: </b> if you would like a Plan Design Comparison generated check the radio button and a wizard will pop up.</p>
        <p> <b>Save Button: </b> saves your changes to the page.</p>
        <p><b>*NOTE</b> the program only allows microsoft Excel 97-2003 files. .xlsx won't work.</p>
        """

        genBoxText = """
        <!DOCTYPE html>
        <html>
        <head>
        </head>
        <body>
        <h1> The Generator Tab </h1>
        <p> <b>Dark gold buttons: </b> these are all the categories contained within the industry report. when you click on the slim button to the right more categories pop up.</p>
        <p> <b>Dark blue buttons: </b> Basic functions for the software.</p>
        <p> <b>Light blue boxes: </b> left: selectable questions based on categories show up here. right: selected questions show up here.</p>
        <div class='tab'>
        <ul>
        <li><br></li>
        <h3>Generating a Run</h3>
        <li>- By clicking on your desired category the field on the right gets populated with the questions present in the given category. </li>
        <li>- Select your desired questions within the left most box by clicking or clicking & sweeping over them. </li>
        <li>- When you have your desired questions selected from the desired category click "Add". </li>
        <li>*NOTE: if you change the category yo selected, your selected questions will reset for that category. Make sure you add them before you go to the next category. </li>
        <li>- After all desired questions are selected and displayed on the right side hit the "Go" button. </li>
        </ul>
        <ul>
        <li><br></li>
        <h3>Deleting Questions</h3>
        <li>- Select the category that contains the desired question(s).</li>
        <li>- Select the desired question(s) from the left most box.</li>
        <li>- Once your question(s) are selected hit the "Delete" button</li>
        </ul>
        </div>
        </body>
        </html>
        """

        self.gentabLabel = QtW.QLabel("<b>Genarator Tab Directions</b>")
        self.settabLabel = QtW.QLabel("<b>Settings Tab Directions</b>")

        self.genTabText = QtW.QPlainTextEdit()
        self.genTabText.appendHtml(genBoxText)
        self.genTabText.show()
        self.genTabText.verticalScrollBar().setValue(self.genTabText.verticalScrollBar().minimum())
        self.genTabText.setReadOnly(True)

        self.setTabText = QtW.QPlainTextEdit()
        self.setTabText.appendHtml(setBoxText)
        self.setTabText.show()
        self.setTabText.verticalScrollBar().setValue(self.setTabText.verticalScrollBar().minimum())
        self.setTabText.setReadOnly(True)

        self.mainDirectionsLayout.addWidget(self.settabLabel)
        self.mainDirectionsLayout.addWidget(self.setTabText)

        self.mainDirectionsLayout.addWidget(self.gentabLabel)
        self.mainDirectionsLayout.addWidget(self.genTabText)

        self.mainGBDirections.setLayout(self.mainDirectionsLayout)
        self.directionsGrid.addWidget(self.mainGBDirections)
        self.directionsTab.setLayout(self.directionsGrid)

class planDesignComparisonWiz(QtW.QWidget):
    """Plan Design Comparison Wizard.

    this opens up upon the check of PDC radio button.
    it furthers options of the generated spreadsheet.

    Methods
    -------
    createSheet(): starts the creation of the generated sheet.
    validateWiz(): Validates inputs entered into the Wizard.
    __init__(): Initialization of the plan design comparison wizard

    """

    def validateWiz(self):
        """Validates inputs entered into the Wizard.

        specifically the client name and the recordkeeper name. if the checks are passed true is returned. if False then an error message pops up and allows you to redo.

        :return:
        """
        errorMessage = ''
        valid = True
        self.inMessageBox = QtW.QMessageBox()
        self.inMessageBox.setStyleSheet(self.stylesheet)

        self.inMessageBox.setWindowTitle("Create Error")

        if self.clientNameIN.text() == '':
            errorMessage += "please enter a client name and try again. \n"
            valid = False
        if self.recordkeeperIN.text() == '':
            errorMessage += "please enter a recordkeeper name and try again. \n"
            valid = False
        if valid == False:
            self.inMessageBox.setText(errorMessage)
            self.inMessageBox.show()
        return valid

    def createSheet(self):
        """starts the creation of the generated sheet.

        when the validation is passed BNFunc.addTable is called with the values provided. after that function finishes
        BNFunc.PlanDesignCompareWrite is called with the needed arguments
        :return:
        """
        valid = self.validateWiz()
        if valid == True:
            selectedOptions = {}
            selectedMarket = self.MarketDD.currentText().replace(" ", '')
            self.workbook.updateMarket(selectedMarket)

            for q in self.selectedQO:
                dd = self.selectedQO[q]
                selectedOptions[q.text()] = [[dd.currentText()]]

            self.workbook.createTitleSheet()
            self.workbook.createPDCSheet(selectedOptions, self.clientNameIN.text(), self.recordkeeperIN.text())
            self.workbook.createDataSheet()
            self.workbook.createDisclosureSheet()
            done = self.workbook.close()

            if done == True:
                self.close()
                self.SSLABEL.setText("Generated " + fileName)
                return True

            elif done == False:
                self.doneSaveMessageBox = QtW.QMessageBox()
                self.doneSaveMessageBox.setStyleSheet(self.stylesheet)
                self.doneSaveMessageBox.setWindowTitle("Save Error")
                errorMessage = "Error when saving. You most likely have a file open with \n the same name in the same location."
                self.doneSaveMessageBox.setText(errorMessage)
                self.doneSaveMessageBox.show()
                return False

    def __init__(self, ssLabel):
        """Initialization of the plan design comparison wizard

        :param values: values attained from the BNF.MatchQuestions
        :param options: dictionary with questions and a list of options within each
        :param ssLabel: success label contained in the main window
        """
        global selectedTemplateQuestions
        super(planDesignComparisonWiz, self).__init__()
        self.workbook = BNFunc.workBook(fileName, outputPath, pdfLoc, selectedTemplateQuestions, True, "All")
        self.OPTIONS = self.workbook.returnOptions()
        self.SSLABEL = ssLabel

        self.PDCGrid = QtW.QGridLayout()
        self.addedLayouts = []
        self.addedDDs = []
        self.addedLabels = []

        self.selectedQO = {}

        self.QuestionsGB = QtW.QGroupBox("Questions and Dropdowns")
        self.questionsVLayout = QtW.QVBoxLayout()

        self.scrollArea = QtW.QScrollArea()
        self.scrollArea.setWidget(self.QuestionsGB)
        self.scrollArea.setWidgetResizable(True)

        sub_headingsText = ['Overall', '<$5MM', '$5MM- $50MM', '>$50MM- $200MM', '>$200MM -$1B', '$>1B']

        self.otherInputLayout = QtW.QVBoxLayout()
        self.MarketVal = QtW.QVBoxLayout()
        self.clientNameInputs = QtW.QVBoxLayout()
        self.recordKeeperInputs = QtW.QVBoxLayout()
        self.labelLayout = QtW.QVBoxLayout()

        stylesheet_path = resource_path("sources\stylesheet.qss")
        self.stylesheet = open(stylesheet_path, "r").read()
        self.setStyleSheet(self.stylesheet)

        self.MarketLabel = QtW.QLabel("Target Market")
        self.MarketLabel.setToolTip("Select your desired market size from the dropdown")
        self.MarketVal.addWidget(self.MarketLabel)

        self.MarketDD = QtW.QComboBox()
        self.MarketVal.addWidget(self.MarketDD)

        self.clientNameLabel = QtW.QLabel("Client Name")
        self.clientNameLabel.setToolTip("Enter the clients name within the input below ")
        self.clientNameInputs.addWidget(self.clientNameLabel)

        self.clientNameIN = QtW.QLineEdit()
        self.clientNameInputs.addWidget(self.clientNameIN)

        self.recordkeeperLabel = QtW.QLabel("Recordkeeper")
        self.recordkeeperLabel.setToolTip("Emter the recordkeeper for the plan in the input below")
        self.recordKeeperInputs.addWidget(self.recordkeeperLabel)

        self.recordkeeperIN = QtW.QLineEdit()
        self.recordKeeperInputs.addWidget(self.recordkeeperIN)

        for i in sub_headingsText:
            self.MarketDD.addItem(i)

        self.otherInputLayout.addItem(self.MarketVal)
        self.otherInputLayout.addItem(self.clientNameInputs)
        self.otherInputLayout.addItem(self.recordKeeperInputs)

        for k,v in self.selectedQO:
            k.deleteLater()
            v.deleteLater()

        self.selectedQO = {}

        for layout in self.addedLayouts:
            self.questionsVLayout.removeItem(layout)

        self.addedLayouts = []
        for q in self.OPTIONS:
            layout = QtW.QHBoxLayout()
            label = QtW.QLabel(q)
            label.setWordWrap(True)
            dropdown = QtW.QComboBox()
            for i in self.OPTIONS[q]:
                dropdown.addItem(i)
            layout.addWidget(label)
            layout.addWidget(dropdown)

            self.selectedQO[label] = dropdown
            self.addedLayouts.append(layout)

        for i in self.addedLayouts:
            self.questionsVLayout.addItem(i)


        self.createPDCButton = QtW.QPushButton("Create Report")
        self.createPDCButton.clicked.connect(self.createSheet)

        self.QuestionsGB.setLayout(self.questionsVLayout)
        self.labelLayout.addWidget(self.scrollArea)

        self.PDCGrid.addItem(self.otherInputLayout, 0, 0)
        self.PDCGrid.addItem(self.labelLayout, 1, 0)
        self.PDCGrid.addWidget(self.createPDCButton)


        self.setFixedHeight(400)
        self.setFixedWidth(700)
        self.setWindowTitle("Plan Design Comparison Wizard")

        self.setLayout(self.PDCGrid)

#starts the loop for the GUI
if __name__ == '__main__':
    setup()
    app = QtW.QApplication(sys.argv)
    gui = mainGUI()
    gui.show()
    sys.exit(app.exec_())