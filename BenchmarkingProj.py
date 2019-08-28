import xlrd
import xlsxwriter
import os
import sys

#questions = {}
indexes = []
blankVal = 0
tableCount = 0
pageBreaks = []

#GLOBAL VARS
title = ''
disclosureList = []
book = ''
questions = {}
sorted_table = []
options = {}

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

def clearResources():
    global title, disclosureList, book, questions, sorted_table, options

    title = ''
    disclosureList = []
    book = ''
    questions = {}
    sorted_table = []
    options = {}


def readPDF(pdf, selectedQs):
    """ Reads the given Excel file

    first creates questions dict with the question as the key.


    :param pdf:
    :param selectedQs:
    :return:
    """
    global title, questions

    wb = xlrd.open_workbook(pdf, formatting_info=True)
    sheet = wb.sheet_by_index(0)
    rows = sheet.nrows
    ATRrange = 0
    validUntilList = []

    #grabs the title of the worksheet
    title = sheet.cell_value(5, 0)

    for q in selectedQs:
        questions[q] = []

    for x in range(rows):
        if str(sheet.cell_value(x, 0)) == 'About this report':
            ATRrange = x
        if 'Valid until' in str(sheet.cell_value(x, 0)) and len(validUntilList) < 10:
            validUntilList.append(x)
        for q in questions:
            if 'How are the following plan expenses/fees paid?' in q:
                text = q.split('(')[1].split(")")[0]
            else:
                text = q
            if str(sheet.cell_value(x, 0)) in text:
                if str(sheet.cell_value(x, 0)) == ":":
                    pass
                elif str(sheet.cell_value(x, 0)) == "/":
                    pass
                elif str(sheet.cell_value(x, 0)) == "%":
                    pass
                elif str(sheet.cell_value(x, 0)) == "":
                    pass
                elif str(sheet.cell_value(x, 0)) == " ":
                    pass
                else:
                    questions[q].append(x)

    for q in questions:
        if len(questions[q]) > 2:
            for x in questions[q]:
                for y in questions[q]:
                    if x + 1 == y:
                        questions[q] = [x, y]

    disclosureTextGrab(sheet, ATRrange, validUntilList)

    #compare ones with the length of 2, see if they picked up a random line with one word match

def matchQuestions(pdfLoc):
    """

    :param questions:
    :param pdfLoc:
    :param name:
    :param dir:
    :return:
    """
    global sorted_table, questions

    wb = xlrd.open_workbook(pdfLoc, formatting_info=True)
    sheet = wb.sheet_by_index(0)
    table = []
    columns = sheet.ncols
    rows = sheet.nrows
    stop = False

    for q in questions:
        try:
            for x in range(int(questions[q][0]), int(rows)):
                for y in range(int(columns)):
                    if str(sheet.cell(x, y).value) == '':
                        if y == 0 and table[-1][1][1] == 12:
                            table.append((" ", (x, y)))
                            stop = True
                    else:
                        table.append((str(sheet.cell(x, y).value), (x, y)))
                if stop == True:
                    stop = False
                    break
        except:
            break

    for i in table:
        if "All Industries" in i[0]:
            pass
        elif title in i[0]:
            pass
        else:
            sorted_table.append(i)


    insert_list = []
    numO = 0

    for i in sorted_table:
        if i[0] == 'Overall' and i[1][1] == 1:
            insert_list.append([sorted_table.index(i) + numO, (" ", (i[1][0], int(i[1][1] - 1)))])
            insert_list.append([sorted_table.index(i) + numO, (title, (i[1][0] - 1, int(i[1][1] + 6)))])
            insert_list.append([sorted_table.index(i) + numO, ("All Industries", (i[1][0] - 1, int(i[1][1])))])
            numO += 3


    for i in insert_list:
        sorted_table.insert(i[0], i[1])



def getOptions(selectedQuestions):
        """Gets option associated with each selected question

        :param values:
        :return: options, finalized dictionary that will hold question(k) and options(v)
        """
        global sorted_table, options
        option_list = [] #list that holds all options undivided
        option_list_mini = [] #recyclable list that holds options associated to the same questions
        option_main = [] #option list that holds all options divided based on similiar questions
        options = {} #finalized dictionary that will hold question(k) and options(v)

        SIndex = 0
        for v in sorted_table:
            if v[1][1] == 0: #if the values column matches 0 it is added to the options list
                option_list.append(v[0])
        for i in option_list:
            if i == " ":
                if option_list_mini != [] and SIndex % 2 != 0: #gets the space between each option set and appends the set into options_main
                    option_main.append(option_list_mini)
                    option_list_mini = []
                    SIndex += 1
                else:
                    SIndex += 1
            else: #if no space is detected then the value is added to the mini container list
                option_list_mini.append(i)

        for l in option_main: #loops through each option list
            question = ''
            Avaloptions = []
            for i in l:
                if any(i in q for q in selectedQuestions) and i != "Average": #if any question in selected questions matches i then the question is picked up and added to the string
                    question += i + " "
                else:
                    Avaloptions.append(i) #otherwise the option is added into AvalOptions
            options[question] = Avaloptions #a key with the question and the options are added into options

def addTable(book, sheet):
    """

    :param book:
    :param table:
    :param questions:
    :param pdcSelected:
    :return:
    """
    global pageBreaks, disclosureList, questions, sorted_table
    x = 0
    y = 0
    pastVal = sorted_table[0][1][0]
    sheet.set_column(0, 0, 65)

    for c in sorted_table:
        if pastVal == c[1][0]:
            writeCell(sheet, x, y, c, book, questions)
            y += 1
        elif pastVal != c[1][0]:
            x += 1
            y = 0
            writeCell(sheet, x, y, c, book, questions)
            y += 1

        pastVal = c[1][0]
    sheet.set_h_pagebreaks(pageBreaks)

def disclosureTextGrab(sheet, ATRrange, validuntilList):
    global disclosureList

    endRange = 0
    for x in validuntilList:
        if int(x) < int(ATRrange):
            pass
        else:
            endRange = x
            break

    for x in range(ATRrange, endRange):
        disclosureList.append(sheet.cell_value(x, 0))



def createSpreedsheet(name, direct):
    """

    :param name:
    :param direct:
    :return:
    """
    book = workBook(name, direct)
    return book

def writeCell(sheet, x, y, c, bk, questions):
    """

    :param sheet:
    :param x:
    :param y:
    :param c:
    :param bk:
    :param questions:
    :return:
    """
    global blankVal, tableCount, pageBreaks
    heading = bk.add_format()
    heading.set_font_size(26)
    heading.set_bold()
    heading.set_bg_color('#153866')
    heading.set_font_color('#ffffff')

    option = bk.add_format()
    option.set_text_wrap(True)
    option.set_font_size(12)
    option.set_italic()
    option.set_bold()
    option.set_font_color('#153866')
    option.set_bottom()

    overall_format = bk.add_format()
    overall_format.set_font_size(12)
    overall_format.set_bg_color('#deaf25')
    overall_format.set_font_color('#ffffff')
    overall_format.set_text_wrap()
    overall_format.set_left()

    sub_heading = bk.add_format()
    sub_heading.set_font_size(12)
    sub_heading.set_bg_color('#deaf25')
    sub_heading.set_font_color('#ffffff')
    sub_heading.set_text_wrap()

    percentForm = bk.add_format()
    percentForm.set_num_format('0.0%')
    percentForm.set_bottom()

    border_format = bk.add_format()
    border_format.set_bottom()

    dollarForm = bk.add_format()
    dollarForm.set_num_format('$0')
    dollarForm.set_bottom()
    headingsText = []

    marketSeg = bk.add_format()
    marketSeg.set_right()
    marketSeg.set_left()
    marketSeg.set_bg_color('#0977b9')
    marketSeg.set_font_color('#ffffff')

    for q in questions:
        headingsText.append(q)

    sub_headingsText = ['Overall', '<$5MM', '$5MM- $50MM', '>$50MM- $200MM', '>$200MM- $1B', '>$200MM -$1B', '$>1B']

    for i in headingsText:
        if c[0] == "Average":
            break
        if c[0] == "3 years":
            break
        if c[0] == "":
            return
        if c[0] == " ":
            if blankVal % 2 == 1:
                if tableCount % 2 == 1:
                    pageBreaks.append(x)
                    blankVal += 1
                    tableCount += 1
                else:
                    blankVal += 1
                    tableCount += 1
                return
            elif blankVal % 2 == 0:
                blankVal += 1
                return
        elif c[0] in i:
            sheet.merge_range('A' + str(x+1) + ':M' + str(x+1), c[0], heading)
            return

    if c[0] in sub_headingsText:
        if c[0] == 'Overall':
            sheet.write(x, y, c[0], overall_format)
            return
        else:
            sheet.write(x, y, c[0], sub_heading)
            return
    elif c[1][1] == 0:
        sheet.write(x, y, c[0], option)
        return
    elif c[0] == 'n/a':
        sheet.write(x, y, c[0], border_format)
    else:
        if c[0] == " ":
            return
        try:
            valueString = c[0].split('.')
            #percent format if
            if len(valueString[0]) == 1:
                if int(valueString[0]) == 1:
                    sheet.write(x, y, float(c[0]), percentForm)
                    return
                elif int(valueString[0]) == 0:
                    sheet.write(x, y, float(c[0]), percentForm)
                    return
                else:
                    sheet.write(x, y, float(c[0]), border_format)
                    return
            #money formating if
            elif len(valueString[0]) > 2:
                sheet.write(x, y, float(c[0]), dollarForm)
                return
            elif len(valueString[0]) > 0 and len(valueString[0]) < 3:
                sheet.write(x, y, float(c[0]), border_format)
                return
            elif int(valueString[1]) != 0:
                sheet.write(x, y, float(c[0]), border_format)
                return
            else:
                sheet.write(x, y, int(c[0]), border_format)
                return
        except:
            if c[0] == "All Industries":
                sheet.merge_range('B' + str(x+1) + ':G' + str(x+1), c[0], marketSeg)
                return
            elif c[0] == title:
                sheet.merge_range('H' + str(x + 1) + ':M' + str(x + 1), c[0], marketSeg)
                return
            else:
                return

def planDesignCompareWrite(book, pdcSheet, selectedMarket, selectedOptions, clientName, recordkeeper):
    """

    :param book:
    :param table:
    :param selectedMarket:
    :param selectedOptions:
    :param questions:
    :param clientName:
    :param recordkeeper:
    :return:
    """
    global sorted_table, questions
    pdcSheet.set_landscape()
    pdcSheet.set_print_scale(90)
    pdcSheet.set_margins(.25, .25, .75, .75)

    ranges = []
    all_ranges = []
    for i in sorted_table:
        if i[0].replace(" ", '') == selectedMarket:
            if len(ranges) == 0:
                ranges.append(i)
            elif len(ranges) == 1:
                ranges.append(i)
                all_ranges.append(ranges)
                ranges = []

    for q in selectedOptions:
        for i in sorted_table:
            if selectedOptions[q][0][0] in i[0]:
                selectedOptions[q][0].append(i[1])

    Qindex = 0
    for k in selectedOptions:
        try:
            selectedOptions[k].append(all_ranges[Qindex])
            Qindex += 1
        except:
            pass

    spaces = []
    for q in selectedOptions:
        for i in selectedOptions[q][0][1:]:
            space_between = abs(int(i[0]) - int(selectedOptions[q][1][0][1][0]))
            spaces.append([space_between, i])
        spaces.sort()
        for v in selectedOptions[q][0][1:]:
            if v != spaces[0][1]:
                selectedOptions[q][0].remove(v)

        spaces = []

    valueCoords = []
    for q in selectedOptions:
        valueCoords.append((selectedOptions[q][0][1][0], selectedOptions[q][1][0][1][1]))
        valueCoords.append((selectedOptions[q][0][1][0], selectedOptions[q][1][1][1][1]))

    PDCvalues = []
    pdcTuple = []
    for i in sorted_table:
        if i[1] in valueCoords:
            if len(pdcTuple) == 0:
                pdcTuple.append(i[0])
            elif len(pdcTuple) == 1:
                pdcTuple.append(i[0])
                PDCvalues.append(pdcTuple)
                pdcTuple = []

    QandValues = {}

    Vindex = 0
    for q in selectedOptions:
        QandValues[q] = [selectedOptions[q][0][0], PDCvalues[Vindex]]
        Vindex += 1

    heading_format = book.add_format()
    heading_format.set_font_size(16)
    heading_format.set_bold()

    givenName_format = book.add_format()
    givenName_format.set_text_wrap()
    givenName_format.set_bottom()

    subheading_format = book.add_format()
    subheading_format.set_font_size(12)
    subheading_format.set_left()
    subheading_format.set_right()
    subheading_format.set_bg_color('#153866')
    subheading_format.set_font_color('#ffffff')

    options_format = book.add_format()
    options_format.set_text_wrap()
    options_format.set_left()
    options_format.set_right()
    options_format.set_bottom()
    options_format.set_bg_color('#f4ce1f')
    options_format.set_font_color('#000000')

    selectedOption_format = book.add_format()
    selectedOption_format.set_text_wrap()
    selectedOption_format.set_bottom()
    selectedOption_format.set_right()

    disclosure_format = book.add_format()
    disclosure_format.set_text_wrap()
    disclosure_format.set_left()
    disclosure_format.set_right()
    disclosure_format.set_top()
    disclosure_format.set_bottom()

    pdcSheet.merge_range('A0:D0', "2019 Plan Design Benchmarking Highlights", heading_format)
    pdcSheet.merge_range('A1:D1', "Plan Name", heading_format)
    pdcSheet.merge_range('E1:N1', clientName, givenName_format)
    pdcSheet.merge_range('A2:D2', "Recordkeeper", heading_format)
    pdcSheet.merge_range('E2:N2', recordkeeper, givenName_format)
    pdcSheet.merge_range('A4:F4', "Plan Design Element", subheading_format)
    pdcSheet.merge_range('G4:J4', clientName, subheading_format)
    pdcSheet.merge_range('K4:M4', str(selectedMarket) + " " + title, subheading_format)
    pdcSheet.merge_range('N4:P4', "All " + str(selectedMarket) + " Plans", subheading_format)

    x = 5
    for i in QandValues:
        pdcSheet.merge_range('A' + str(x) + ':F' + str(x+1), i, options_format)
        pdcSheet.merge_range('G' + str(x) + ':J' + str(x+1), QandValues[i][0], selectedOption_format)
        writePDCCell(pdcSheet,'K' + str(x) + ':M' + str(x+1), QandValues[i][1][0], book)
        writePDCCell(pdcSheet, 'N' + str(x) + ':P' + str(x+1), QandValues[i][1][1], book)
        x += 2

    path = resource_path('sources\ACG-Logo-Full-S.jpg')
    pdcSheet.insert_image("A" + str(x+2), path)

def disclosureSheet(book, dcSheet, selectedMarket):

    global disclosureList

    disclosure_format = book.add_format()
    disclosure_format.set_text_wrap()
    disclosure_format.set_left()
    disclosure_format.set_right()
    disclosure_format.set_top()
    disclosure_format.set_bottom()
    disclosureScrappedtext = ''

    for x in disclosureList:
        disclosureScrappedtext += x + "\n"
    disclosure_text = 'Industry statistics are taken from the PLANSPONSOR Defined Contribution Survey, 2018 for ' + (
        title) + ' plans (' + str(
        selectedMarket) + "). " + disclosureScrappedtext
    dcSheet.merge_range('A' + str(1) + ":P" + str(35), disclosure_text, disclosure_format)

def writePDCCell(sheet, range, c, bk):
    """

    :param sheet:
    :param range:
    :param c:
    :param bk:
    :return:
    """

    percentForm = bk.add_format()
    percentForm.set_num_format('0.0%')
    percentForm.set_bottom()
    percentForm.set_right()

    dollarForm = bk.add_format()
    dollarForm.set_num_format('$0')
    dollarForm.set_bottom()
    dollarForm.set_right()

    normal_format = bk.add_format()
    normal_format.set_bottom()
    normal_format.set_right()

    try:
        valueString = c.split('.')
        # percent format if
        if len(valueString[0]) == 1:
            if int(valueString[0]) == 1:
                sheet.merge_range(range, float(c), percentForm)
                return
            elif int(valueString[0]) == 0:
                sheet.merge_range(range, float(c), percentForm)
                return
            else:
                sheet.merge_range(range, float(c), normal_format)
                return
        # money formating if
        elif len(valueString[0]) > 2:
            sheet.merge_range(range, float(c), dollarForm)
            return
        elif len(valueString[0]) > 0 and len(valueString[0]) < 3:
            sheet.merge_range(range, float(c), normal_format)
            return
        elif int(valueString[1]) != 0:
            sheet.merge_range(range, float(c), normal_format)
            return
        else:
            sheet.merge_range(range, int(c), normal_format)
            return
    except:
        if c == "n/a":
            sheet.merge_range(range, c, normal_format)
            return
        else:
            return

class workBook():
    """Class to generate a workbook used in writing values too

    Methods
    -------

    __init__(): Initialization of the excel workbook
    close(): Closes the workbook aka saves
    createTitleSheet(): Creates the title sheet.
    """
    def __init__(self, name, direct, pdfLoc, selectedQuestions, PDCCheck, selectedMarket):
        """ Initialization of the excel workbook

        this creates the workbook in the desired directory. then creates the title sheet.

        :param name: the name of the excel file
        :param direct: directory being saved into
        """
        self.NAME = name
        self.DIRECTORY = direct
        self.PDFLOC = pdfLoc
        self.SELECTEDQS = selectedQuestions
        self.PDCCHECK = PDCCheck
        self.SELECTEDMARKET = selectedMarket
        self.pullInfo()
        self.book = xlsxwriter.Workbook(self.DIRECTORY + "/" + self.NAME + ".xlsx")

    def pullInfo(self):
        readPDF(self.PDFLOC, self.SELECTEDQS)
        matchQuestions(self.PDFLOC)
        getOptions(self.SELECTEDQS)

    def updateMarket(self, selectedMarket):
        self.SELECTEDMARKET = selectedMarket

    def returnOptions(self):
        global options
        return options


    def createTitleSheet(self):
        """Creates the title sheet.

        it adds in a picture and the title to the page

        :return:
        """
        global title
        txtBoxOption = {
            'border': {'width': 0,
                       'none': True},
            'font': {'bold': True,
                     'size': 26,
                     'color': '#153866'},
            'align': {'horizontal': 'center',
                      'vertical': 'middle'},
            'width': 400
        }

        self.titleSheet = self.book.add_worksheet('Cover')
        path = resource_path('sources\ACG-Horizontal-Background.jpg')
        self.titleSheet.insert_image("A1", path, {'x_scale': 0.262, 'y_scale': 0.282})
        self.titleSheet.insert_textbox(19, 3, str(title), txtBoxOption)
        self.titleSheet.set_landscape()

    def createPDCSheet(self, selectedOptions, clientName, recordkeeper):
        self.PDCSheet = self.book.add_worksheet("Plan Design Comparison")
        planDesignCompareWrite(self.book, self.PDCSheet, self.SELECTEDMARKET, selectedOptions, clientName, recordkeeper)

    def createDataSheet(self):
        self.dataSheet = self.book.add_worksheet("Data")
        addTable(self.book, self.dataSheet)
        self.dataSheet.set_landscape()
        self.dataSheet.set_print_scale(75)
        self.dataSheet.set_margins(.25, .25, .75, .75)

    def createDisclosureSheet(self):
        self.dcSheet = self.book.add_worksheet("Disclosure")
        disclosureSheet(self.book, self.dcSheet, self.SELECTEDMARKET)

    def close(self):
        """ Closes the workbook aka saves

        :return:
        """
        try:
            self.book.close()
            return True
        except:
            return False



if __name__ == "__main__":
    pass
