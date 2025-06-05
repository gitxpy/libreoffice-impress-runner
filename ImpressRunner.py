#   The Contents of this file are made available subject to the terms of
#   the following license
#
#          - GNU Lesser General Public License Version 2.1
#
#   GNU Lesser General Public License Version 2.1
#   =============================================
#   Copyright 2005 by Sun Microsystems, Inc.
#   901 San Antonio Road, Palo Alto, CA 94303, USA
#
#   This library is free software; you can redistribute it and/or
#   modify it under the terms of the GNU Lesser General Public
#   License version 2.1, as published by the Free Software Foundation.
#
#   This library is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
#   Lesser General Public License for more details.
#
#   You should have received a copy of the GNU Lesser General Public
#   License along with this library; if not, write to the Free Software
#   Foundation, Inc., 59 Temple Place, Suite 330, Boston,
#   MA  02111-1307  USA
#
# Author Laurent Godard - LaurentGodard@openoffice.org

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
        arg1 = args[0]
        for struct in arg1.Value:
            if struct.Name=='Model':
                self.document = struct.Value

        if self.document.supportsService("com.sun.star.presentation.PresentationDocument"):
            is_autostart = False
            docInfo = self.document.DocumentProperties.UserDefinedProperties
            if docInfo.PropertySetInfo.hasPropertyByName("autostart"):
                is_autostart = docInfo.getPropertyValue("autostart")

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