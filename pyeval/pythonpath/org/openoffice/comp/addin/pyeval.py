import unohelper
from org.openoffice.addin import XPyEval
from com.sun.star.sheet import XAddIn
from com.sun.star.lang import XLocalizable, XServiceName, Locale

#from pyextension_helper import ExtensionBase

class ExtensionBase(unohelper.Base, XPyEval, XAddIn, XServiceName):
    service_name = "ServiceName"
    function_description = "text describing the function"
    argument_names = ["argument"]
    argument_descriptions = ["argument 0 description"]
    def __init__( self, ctx ):
        self.ctx = ctx
        self.locale = Locale("en","EN", "" )

    def getServiceName( self ):
        return self.service_name

    def setLocale(self, locale):
        self.locale = locale

    def getLocale(self):
        return self.locale

    def getProgrammaticFuntionName(self, aProgramaticName):
        return aProgramaticName

    def getDisplayFunctionName(self, aDisplayName):
        return aDisplayName

    def getFunctionDescription(self , aProgrammaticName):
        return self.function_description

    def getDisplayArgumentName(self, aProgrammaticFunctionName, argument_index):
        return self.argument_names[argument_index]

    def getArgumentDescription(self, aProgrammaticFunctionName, argument_index):
        return self.argument_descriptions[argument_index]

    def getProgrammaticCategoryName(self, aProgrammaticFunctionName):
        return "Add-In"

    def getDisplayArgumentName(self, aProgrammaticFunctionName):
        return "Add-In"

    #def pyeval( self, str ):
        #return eval(str)

result = ""


class PyEval(ExtensionBase):
    service_name = "pythonEvaluator"
    function_description = "evaluates a python expression and returns a float"
    argument_names = ["expression"]
    argument_descriptions = ["python expression to evaluate"]

    def pyeval(self, formula):
        return eval(formula)

    def pyexec(self, str):
        global result

        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            'com.sun.star.frame.Desktop', self.ctx)

        # get current document model
        model = desktop.getCurrentComponent()

        # access the document:
        sheets = model.Sheets
        s1  = sh.Sheet1

        exec(formula, globals(), locals())
        return result

class Cell(object):
    def __init__(self, uno_cell, address=None):
        self._cell = uno_cell
        self.address = address

    def __repr__(self):
        return "Cell {} - {}".format(self.address, self._cell.getFormula())

    def setFormula(self, value):
        self._cell.setFormula(value)

class CellArray(object):
    pass

class Sheet(object):
    def __init__(self, uno_sheet, parent=None):
        self._sheet = uno_sheet
        self.parent = parent

    def __getitem__(self, name_or_address):
        if isinstance(name_or_address, str):
            cell = self._sheet.getCellRangeByName(name_or_address)
            if not hasattr(cell, "getFormula"):
                return CellArray(cell)
        elif isinstance(name_or_address, tuple):
            if len(name_or_address) == 3:
                cell = self.parent.getCellByPosition(*name_or_address)
            else:
                cell = self._sheet.getCellByPosition(*name_or_address)
        return Cell(cell, name_or_address)

    def __setitem__(self, name_or_address, value):
        cell = self.__getitem__(name_or_address)
        cell.setFormula(value)

    name = property(lambda s:s._sheet.getName())



class SpreadSheet(object):
    def __init__(self, sheets):
        self._sheets = sheets
        self.sheets = []
        self.sheets_by_name = {}
        for sheet in self.enumerate_sheets():
            sh = Sheet(sheet, sheets)
            self.sheets.append(sh)
            self.sheets_by_name[sh.name] = sh

    def enumerate_sheets(self):
        enum = self._sheets.createEnumeration()
        sheet_list = []
        while enum.hasMoreElements():
            yield enum.nextElement()

    def __getitem__(self, index):
        if isinstance(index, str):
            return self.sheets_by_name[index]
        return self.sheets[index]

    def __repr__(self):
        return "Spreadsheet({})".format(", ".join("'{}'".format(sh.name)
                                        for sh in self.sheets))
