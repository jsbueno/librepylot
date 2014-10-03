# coding: utf-8

from datetime import date, datetime, timedelta
import unohelper
from org.openoffice.addin import XPyEval
from com.sun.star.sheet import XAddIn
from com.sun.star.lang import XLocalizable, XServiceName, Locale


date_formats = {30, 32, 33, 34, 35, 36,37, 38, 39, 82, 83, 84}

###############

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



class PyEval(ExtensionBase):
    service_name = "pythonEvaluator"
    function_description = "evaluates a python expression and returns a float"
    argument_names = ["expression"]
    argument_descriptions = ["python expression to evaluate"]

    def pyeval(self, formula):
        return eval(formula)

    def pyexec(self, formula):
        result = ""
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            'com.sun.star.frame.Desktop', self.ctx)

        # get current document model
        model = desktop.getCurrentComponent()

        # access the document:
        spreadsheet = S = SpreadSheet(model.Sheets)

        exec(formula, globals(), locals())
        return result

class Cell(object):
    def __init__(self, uno_cell, address=None):
        self._cell = uno_cell
        self.address = address

    def __repr__(self):
        return "Cell {} - {}".format(self.address, self.formula)

    def setFormula(self, value):
        if isinstance(value, (datetime, date)):
            value = date_to_number(value)
            if self._cell.NumberFormat not in date_formats:
                # ISO yyyy-mm-dd format:
                self._cell.NumberFormat = 84
        self._cell.setFormula(value)

    def getFormula(self, value=False):
        if value:
            value = self._cell.getValue()
        else:
            value = self._cell.getFormula()
        if self._cell.NumberFormat in date_formats:
            return number_to_date(int(value))
        return value

    formula = property(getFormula, setFormula, lambda s: s.setFormula(""))
    value = property(lambda s: s.getFormula(True), setFormula,
                     lambda s: s.setFormula(""))

    def _set_color(self, color=None):
        if color is None:
            color = -1
        else:
            color = (color[0] << 16) + (color[1] << 8) + color[2]
        self._cell.CellBackColor = color

    def _get_color(self):
        color = self._cell.CellBackColor
        if color == -1:
            return None
        return (color >> 16, color >> 8 & 0xff, color &0xff)

    backcolor = back_color = property(_get_color, _set_color)


class CellArray(object):
    pass


class Line(object):
    def __init__(self, line, direction, address=None):
        self.line = line
        self.address = address
        self.direction = direction

    def __getitem__(self, index):
        args = (index, 0) if self.direction == "Rows" else (0, index)
        # TODO: calculate cell addresss
        return Cell(self.line.getCellByPosition(*args))

    def __setitem__(self, index, value):
        self.__getitem__(index).formula = value

    visible = property(lambda s:getattr(s.line, "IsVisible"),
                       lambda s, v: setattr(s.line, "IsVisible", v))


class LineContainer(object):
    def __init__(self, sheet,direction):
        self.sheet = sheet
        self.direction = direction
        self.lines = getattr(self.sheet._sheet, direction)

    def __getitem__(self, index):
        if isinstance(index, str):
            return Line(self.lines.getByName(index), self.direction, index)
        return Line(self.lines.getByIndex(index), self.direction, index)


class LineProperty(object):
    def __init__(self, direction):
        self.direction = direction

    def __get__(self, instance, owner):
        if not instance:
            return self
        return LineContainer(instance, self.direction)


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
        cell.formula = value

    rows = Rows = LineProperty("Rows")
    cols = columns = Columns = LineProperty("Columns")

    name = property(lambda s:s._sheet.getName())



class SpreadSheet(object):
    def __init__(self, sheets):
        self._sheets = sheets
        self.update()

    def update(self):
        self.sheets = []
        self.sheets_by_name = {}
        for sheet in self.enumerate_sheets():
            sh = Sheet(sheet, self._sheets)
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

    def __len__(self):
        return len(self.sheets)

    def __repr__(self):
        return "Spreadsheet({})".format(", ".join("'{}'".format(sh.name)
                                        for sh in self.sheets))

#########
# For external use of L.O.:

def connect():
    """
    Start LibreOffice with
    libreoffice -calc -accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
    and call this from a Python interactive console
    """
    import uno

    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
                                    "com.sun.star.bridge.UnoUrlResolver", localContext )
    ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)


    # get current document model
    model = desktop.getCurrentComponent()
    return SpreadSheet(model.Sheets)

# Convert between sequential numbers since 1900-1-1 and Python datetime
# taking into account excel's 1900 Saint Tiby's bug:

def number_to_date(value):
    return datetime(1900, 1, 1) + timedelta(value - 2)

def date_to_number(date_value):
    if isinstance(date_value, date):
        date_value = datetime(date_value.year, date_value.month, date_value.day)
    # TODO: add proper accounting to actually take time into account
    return  (date_value - datetime(1900,1,1)).days + 2
