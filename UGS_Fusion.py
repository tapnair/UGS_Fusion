#Author-Patrick Rainsberry
#Description-Universal G-Code Sender plugin for Fusion 360

import adsk.core, traceback
import adsk.fusion
import tempfile


from xml.etree import ElementTree
from xml.etree.ElementTree import SubElement

from os.path import expanduser
import os

handlers = []

def getFileName():
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        
        home = expanduser("~")
        home += '/UGS_Fusion/'
        
        if not os.path.exists(home):
            os.makedirs(home)

        xmlFileName = home  + 'settings.xml'
        
        return xmlFileName
    
    except:
        if ui:
            ui.messageBox('Panel command created failed:\n{}'
            .format(traceback.format_exc()))
def writeSettings(xmlFileName, UGS_path, UGS_post, UGS_platform, showOperations):
    
    if not os.path.isfile(xmlFileName):
        new_file = open( xmlFileName, 'w' )                        
        new_file.write( '<?xml version="1.0"?>' )
        new_file.write( "<UGS_Fusion /> ")
        new_file.close()
        tree = ElementTree.parse(xmlFileName) 
        root = tree.getroot()
    else:
        # TODO delete node
        tree = ElementTree.parse(xmlFileName) 
        root = tree.getroot()
        root.remove(root.find('settings'))

    settings = SubElement(root, 'settings')
    SubElement(settings, 'UGS_path', value = UGS_path)
    SubElement(settings, 'UGS_post', value = UGS_post)
    if UGS_platform == True:
        SubElement(settings, 'UGS_platform', value = 'True')
    else:
        SubElement(settings, 'UGS_platform', value = 'False')
    
    if showOperations == True:
        SubElement(settings, 'showOperations', value = 'True')
    else:
        SubElement(settings, 'showOperations', value = 'False')
    
    # Write settings to XML File
    tree.write(xmlFileName)
    
def readSettings(xmlFileName):
    
    tree = ElementTree.parse(xmlFileName) 
    root = tree.getroot()

    UGS_path = root.find('settings/UGS_path').attrib[ 'value' ]
    UGS_post = root.find('settings/UGS_post').attrib[ 'value' ]
    UGS_platform_text = root.find('settings/UGS_platform').attrib[ 'value' ]
    showOperations_text = root.find('settings/showOperations').attrib[ 'value' ]
    
    if UGS_platform_text == 'True':
        UGS_platform = True
    else :
        UGS_platform = False
    
    if showOperations_text == 'True':
        showOperations = True
    else:
        showOperations = False    
        
        
    return(UGS_path, UGS_post, UGS_platform, showOperations)
    
def exportFile(opName, UGS_path, UGS_post, UGS_platform):

    app = adsk.core.Application.get()
    ui = app.userInterface
    doc = app.activeDocument
    products = doc.products
    product = products.itemByProductType('CAMProductType')

    # check if the document has a CAMProductType.  I will not if there are no CAM operations in it.
    if product == None:
        ui.messageBox('There are no CAM operations in the active document.  This script requires the active document to contain at least one CAM operation.',
                        'No CAM Operations Exist',
                        adsk.core.MessageBoxButtonTypes.OKButtonType,
                        adsk.core.MessageBoxIconTypes.CriticalIconType)
        return

    cam = adsk.cam.CAM.cast(product)

    # Create a temporary directory.
    outputFolder = tempfile.mkdtemp()
    
    # Get Setup
    for setup in cam.setups:
        if setup.name == opName:
            toPost = setup
        else:
            for operation in setup.operations:
                if operation.name == opName:
                    toPost = operation
    
    # Set the post options        
    postConfig = os.path.join(cam.genericPostFolder, UGS_post) 
    viewResult = False
    units = adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput
    programName = opName
    
    # create the postInput object
    postInput = adsk.cam.PostProcessInput.create(programName, postConfig, outputFolder, units)
    postInput.isOpenInEditor = viewResult

    # Post Process Setup
    cam.postProcess(toPost, postInput)
    
    # Get the resulting filename
    resultFilename = outputFolder + '//' + programName
    resultFilename = resultFilename + '.nc'
    import subprocess
    
#    
    if UGS_platform:
        process = subprocess.Popen([UGS_path, '--open', '%s'% resultFilename])
    else:
        process = subprocess.Popen(['java', '-jar', UGS_path, '--open', '%s'% resultFilename])
   
   # open the output folder in Finder on Mac or in Explorer on Windows
#    if (os.name == 'posix'):
#        os.system('open "%s"' % outputFolder)
#    elif (os.name == 'nt'):
#        os.startfile(outputFolder)
            
    return resultFilename

# Get the current values of the command inputs.
def getInputs(inputs):
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        opInput = inputs.itemById('operations')
        opName = opInput.selectedItem.name
        UGS_path = inputs.itemById('UGS_path').text
        UGS_post = inputs.itemById('UGS_post').text
        UGS_platform = inputs.itemById('UGS_platform').value
        saveSettings = inputs.itemById('saveSettings').value
        showOperations = inputs.itemById('showOperations').value
        
        return (opName, UGS_path, UGS_post, UGS_platform, saveSettings, showOperations)
    except:
        app = adsk.core.Application.get()
        ui = app.userInterface
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

# Define the event handler for Octoprint command is executed (the "Create RFQ" button is clicked on the dialog).
class UGSExecutedEventHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):

        ui = []
        try:
            app = adsk.core.Application.get()
            ui  = app.userInterface

            # Get the inputs.
            inputs = args.command.commandInputs
            (opName, UGS_path, UGS_post, UGS_platform, saveSettings, showOperations) = getInputs(inputs)
            
            # Save Settings:
            if saveSettings:
                xmlFileName = getFileName()
                writeSettings(xmlFileName, UGS_path, UGS_post, UGS_platform, showOperations)
            
            # Export the file and launch UGS
            exportFile(opName, UGS_path, UGS_post, UGS_platform)

            
        except:

            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Define the event handler for when any input changes.
class UGSInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        ui = []
        try:
            # Get the application.
            app = adsk.core.Application.get()
            doc = app.activeDocument
            products = doc.products
            product = products.itemByProductType('CAMProductType')
            
            # Check if the document has a CAMProductType. It will not if there are no CAM operations in it.
            if product == None:
                 ui.messageBox('There are no CAM operations in the active document')
                 return
            
            # Cast the CAM product to a CAM object (a subtype of product).
            cam = adsk.cam.CAM.cast(product)
            ui  = app.userInterface

            input_changed = args.input
            inputs = args.inputs

            # Check to see if the file type has changed and clear the selection
            # re-populate the selection filter.
            if input_changed.id == 'showOperations':
                opInput = inputs.itemById('operations')
                
                for item in opInput.listItems:
                    item.deleteMe()
                
                for setup in cam.setups:
                    opInput.listItems.add(setup.name, False)
                    if input_changed.value:
                        for operation in setup.operations:
                            opInput.listItems.add(operation.name, False)
                            
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

                
# Define the event handler for when the command is activated.
class UGSCommandActivatedHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        ui = []
        try:
            app = adsk.core.Application.get()
            ui  = app.userInterface

        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

                
# Define the event handler for when the Octoprint command is run by the user.
class UGSCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        ui = []
        try:
            
            # Get the application.
            app = adsk.core.Application.get()
            
            # Get the active document.
            doc = app.activeDocument
            
            # Get the products collection on the active document.
            products = doc.products
            
            # Get the CAM product.
            product = products.itemByProductType('CAMProductType')
            
            # Check if the document has a CAMProductType. It will not if there are no CAM operations in it.
            if product == None:
                 ui.messageBox('There are no CAM operations in the active document')
                 return
            
            # Cast the CAM product to a CAM object (a subtype of product).
            cam = adsk.cam.CAM.cast(product)

            ui  = app.userInterface

            # Connect to the command executed event.
            cmd = args.command
            cmd.isExecutedWhenPreEmpted = False
            onExecute = UGSExecutedEventHandler()
            cmd.execute.add(onExecute)
            handlers.append(onExecute)

            onInputChanged = UGSInputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            handlers.append(onInputChanged)
            
            # Connect to the command activated event.
            onActivate = UGSCommandActivatedHandler()
            cmd.activate.add(onActivate)
            handlers.append(onActivate)

            # Define the inputs.
            inputs = cmd.commandInputs
            
            # Labels
            inputs.addTextBoxCommandInput('labelText2', '', '<a href="http://winder.github.io/ugs_website/">Universal Gcode Sender</a></span> A full featured gcode platform used for interfacing with advanced CNC controllers like GRBL and TinyG.', 4, True)
            inputs.addTextBoxCommandInput('labelText3', '', 'Choose the Setup or Operation to send to UGS', 2, True)
            
            # UGS User Information 
            UGS_path_input = inputs.addTextBoxCommandInput('UGS_path', 'UGS Path: ', 'Location of UGS', 1, False)
            UGS_post_input = inputs.addTextBoxCommandInput('UGS_post', 'Post to use: ', 'Name of post', 1, False)
            
            # Whether using classic or platform
            # Could automate this based on path
            UGS_platform_input = inputs.addBoolValueInput("UGS_platform", 'Using UGS Platform?', True)
            
            # Show only setups or also operations?
            showOperations_input = inputs.addBoolValueInput("showOperations", 'Select Operations?', True)
            
            # Save user settings?
            inputs.addBoolValueInput("saveSettings", 'Save settings?', True)
            
            # Drop down for Operations and Setups
            opDropDown = inputs.addDropDownCommandInput('operations', 'Select Setup or Operation', adsk.core.DropDownStyles.LabeledIconDropDownStyle)

            cmd.commandCategoryName = 'UGS'
            cmd.setDialogInitialSize(500, 300)
            cmd.setDialogMinimumSize(500, 300)

            cmd.okButtonText = 'POST'
            
            xmlFileName = getFileName()
            if os.path.isfile(xmlFileName):
                (UGS_path, UGS_post, UGS_platform, showOperations) = readSettings(xmlFileName)
                
                UGS_path_input.text = UGS_path
                UGS_post_input.text = UGS_post
                UGS_platform_input.value = UGS_platform
                showOperations_input.value = showOperations
            
            for setup in cam.setups:
                opDropDown.listItems.add(setup.name, False)
                if showOperations_input.value:
                    for operation in setup.operations:
                        opDropDown.listItems.add(operation.name, False)

        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def run(context):
    ui = None

    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        if ui.commandDefinitions.itemById('UGSButtonID'):
            ui.commandDefinitions.itemById('UGSButtonID').deleteMe()

        # Get the CommandDefinitions collection.
        cmdDefs = ui.commandDefinitions

        # Create a button command definition for the comamnd button.  This
        # is also used to display the disclaimer dialog.
        tooltip = '<div style=\'font-family:"Calibri";color:#B33D19; padding-top:-20px;\'><span style=\'font-size:20px;\'><b>winder.github.io/ugs_website</b></span></div>Universal Gcode Sender'
        UGSButtonDef = cmdDefs.addButtonDefinition('UGSButtonID', 'Post to UGS', tooltip, './/Resources')
        onUGSCreated = UGSCreatedEventHandler()
        UGSButtonDef.commandCreated.add(onUGSCreated)
        handlers.append(onUGSCreated)

        # Find the "ADD-INS" panel for the solid and the surface workspaces.
        solidPanel = ui.allToolbarPanels.itemById('CAMActionPanel')
        
        # Add a button for the "Request Quotes" command into both panels.
        buttonControl = solidPanel.controls.addCommand(UGSButtonDef, '', False)
    except:
        pass
        #if ui:
        #    ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        if ui.commandDefinitions.itemById('UGSButtonID'):
            ui.commandDefinitions.itemById('UGSButtonID').deleteMe()

        # Find the controls in the solid and surface panels and delete them.
        camPanel = ui.allToolbarPanels.itemById('CAMActionPanel')
        cntrl = camPanel.controls.itemById('UGSButtonID')
        if cntrl:
            cntrl.deleteMe()


    except:
        pass
        #if ui:
        #    ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


