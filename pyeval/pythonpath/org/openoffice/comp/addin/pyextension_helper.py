import unohelper
from org.openoffice.addin import XPyEval
from com.sun.star.sheet import XAddIn
from com.sun.star.lang import XLocalizable, XServiceName, Locale

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



