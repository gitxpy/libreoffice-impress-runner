import uno
import unohelper

from com.sun.star.uno import RuntimeException as _rtex

# equivalent to Basic, for file URL only
from uno import systemPathToFileUrl as convertToURL
from uno import fileUrlToSystemPath as convertFromURL

from com.sun.star.beans import PropertyValue
from com.sun.star.task import XJob

implementation_name = "org.openoffice.extensions.indesko.ImpressRunner"
implementation_services = ("com.sun.star.task.Job",)

#### usefull helpers ####

import sys
from com.sun.star.connection import NoConnectException
from com.sun.star.beans import PropertyValue

class OOoTools:
    """helper tools for using pyUNO"""
        
    #----------------------------------------
    #   Danny's stuff to make programming less convenient.
    #   http://www.oooforum.org/forum/viewtopic.phtml?t=9115 
    #----------------------------------------

    def __init__(self, ctx):
        self.oCoreReflection = None
        self.desktop = None
        self.ctx = ctx
        self.desktop = self.getDesktop()
        return

    def getServiceManager( self ):
        """Get the ServiceManager from the running OpenOffice.org.
        """
        return self.ctx.ServiceManager
   
    def createUnoService( self, cClass ):
        """A handy way to create a global objects within the running OOo.
        """
        oServiceManager = self.getServiceManager()
        oObj = oServiceManager.createInstance( cClass )
        return oObj
   
    def getDesktop( self ):
        """An easy way to obtain the Desktop object from a running OOo.
        """
        if self.desktop == None:
            self.desktop = self.createUnoService("com.sun.star.frame.Desktop")
        return self.desktop

    def getCoreReflection( self ):
        if self.oCoreReflection == None:
            self.oCoreReflection = self.createUnoService(
                                    "com.sun.star.reflection.CoreReflection" )
        return self.oCoreReflection 
        
    def createUnoStruct( self, cTypeName ):
        """Create a UNO struct and return it.
        """
        oCoreReflection = self.getCoreReflection()

        # Get the IDL class for the type name
        oXIdlClass = oCoreReflection.forName( cTypeName )

        # Create the struct.
        oReturnValue, oStruct = oXIdlClass.createObject( None )

        return oStruct

    def makePropertyValue( self, cName=None, uValue=None,
                                 nHandle=None, nState=None ):
        """Create a com.sun.star.beans.PropertyValue struct and return it.
        """
        oPropertyValue = self.createUnoStruct(
                                    "com.sun.star.beans.PropertyValue" )

        if cName != None:
            oPropertyValue.Name = cName
        if uValue != None:
            oPropertyValue.Value = uValue
        if nHandle != None:
            oPropertyValue.Handle = nHandle
        if nState != None:
            oPropertyValue.State = nState

        return oPropertyValue 

class ImpressRunner(unohelper.Base, XJob):

    def __init__ (self, ctx):

        self.ctx = ctx
        self.tools = OOoTools(ctx)
        self.desktop = self.tools.getDesktop()

    def execute(self, args):
        self.document = self.desktop.getCurrentComponent()
        if self.document.supportsService("com.sun.star.presentation.PresentationDocument"):
            #can we detect autostart to any of the 4 UserFieldValue
            is_autostart = False
            for i in range(4):
                if self.document.DocumentInfo.getUserFieldValue(i) == "autostart":
                    is_autostart = True
            # the document is an impress document
            if (is_autostart or self.document.URL.endswith('.pps')):
                # the metadata is there !! or it is a .pps file
                # launch the presentation
                self.document.Presentation.start()

        return

g_TypeTable = {}
# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper ()

# add the FormatFactory class to the implementation container,
# which the loader uses to register/instantiate the component.
g_ImplementationHelper.addImplementation (ImpressRunner,
					implementation_name,
					implementation_services)