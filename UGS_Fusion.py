#Author-Patrick Rainsberry
#Description-Universal G-Code Sender plugin for Fusion 360

import adsk.core, traceback
import adsk.fusion

from xml.etree import ElementTree
from xml.etree.ElementTree import SubElement

from os.path import expanduser
import os
import tempfile
import subprocess

# Global Variable to handle command events
handlers = []

# Global Program ID for settings
programID = 'UGS_Fusion'


def get_folder():
    # Get user's home directory
    home = expanduser("~")

    # Create a subdirectory for this application settings
    home += '/' + programID + '/'

    # Create the folder if it does not exist
    if not os.path.exists(home):
        os.makedirs(home)

    return home


# Will get the proposed location of the settings file for this application
# Note currently creates the directory with out being prompted to save settings
def getFileName():
    
    home = get_folder()
    
    # Get full path for the settings file
    xmlFileName = home  + 'settings.xml'
    
    return xmlFileName

# Write users settings to a local file 
def writeSettings(xmlFileName, UGS_path, UGS_post, UGS_platform, showOperations):
    
    # If the file does not exist create it
    if not os.path.isfile(xmlFileName):
        # Create the file
        new_file = open( xmlFileName, 'w' )                        
        new_file.write( '<?xml version="1.0"?>' )
        new_file.write( '<' + programID + ' /> ')
        new_file.close()
        
        # Open the file and parse as XML        
        tree = ElementTree.parse(xmlFileName) 
        root = tree.getroot()
    
    # Read in the file and get the tree
    else:
        tree = ElementTree.parse(xmlFileName) 
        root = tree.getroot()
        
        # Remove old settings info
        root.remove(root.find('settings'))
    
    # Write settings info into XML file
    settings = SubElement(root, 'settings')
    SubElement(settings, 'UGS_path', value = UGS_path)
    SubElement(settings, 'UGS_post', value = UGS_post)
    SubElement(settings, 'showOperations', value = showOperations)
    
    # Create local boolean value for platform setting    
    if UGS_platform == True:
        SubElement(settings, 'UGS_platform', value = 'True')
    else:
        SubElement(settings, 'UGS_platform', value = 'False')

    # Write settings to XML File
    tree.write(xmlFileName)

# Read in user's local settings 
# No real error checking done here
def readSettings(xmlFileName):
    
    # Parse settings file as XML
    tree = ElementTree.parse(xmlFileName) 
    root = tree.getroot()

    # Find relevant settings
    UGS_path = root.find('settings/UGS_path').attrib[ 'value' ]
    UGS_post = root.find('settings/UGS_post').attrib[ 'value' ]
    UGS_platform_text = root.find('settings/UGS_platform').attrib[ 'value' ]
    showOperations = root.find('settings/showOperations').attrib[ 'value' ]
    
    # Handle booleans in settings
    if UGS_platform_text == 'True':
        UGS_platform = True
    else :
        UGS_platform = False

    return(UGS_path, UGS_post, UGS_platform, showOperations)

# Post process selected operation and launch UGS
def exportFile(opName, UGS_path, UGS_post, UGS_platform):

    app = adsk.core.Application.get()
    doc = app.activeDocument
    products = doc.products
    product = products.itemByProductType('CAMProductType')
    cam = adsk.cam.CAM.cast(product)

    # Iterate through CAM objects for operation, folder or setup
    # Currently doesn't handle duplicate in names
    for setup in cam.setups:
        if setup.name == opName:
            toPost = setup
        else:
            for folder in setup.folders:
                if folder.name == opName:
                    toPost = folder       
   
    for operation in cam.allOperations:
        if operation.name == opName:
            toPost = operation
            
    # Create a temporary directory for post file
    # outputFolder = tempfile.mkdtemp()

    outputFolder = get_folder() + "//output/"
    
    # Set the post options        
    postConfig = os.path.join(cam.genericPostFolder, UGS_post) 
    units = adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput

    # create the postInput object
    postInput = adsk.cam.PostProcessInput.create(opName, postConfig, outputFolder, units)
    postInput.isOpenInEditor = False
    cam.postProcess(toPost, postInput)
    
    # Get the resulting filename
    resultFilename = outputFolder + '//' + opName
    resultFilename = resultFilename + '.nc'

    # Use subprocess to launch UGS in a new process, check if platform or java
    if UGS_platform:
        subprocess.Popen([UGS_path, '--open', '%s' % resultFilename])
    else:
        subprocess.Popen(['java', '-jar', UGS_path, '--open', '%s'% resultFilename])
   
    return resultFilename

# Get the current values of the command inputs.
def getInputs(inputs):
        
        # Look up name of input and get value
        UGS_path = inputs.itemById('UGS_path').text
        UGS_post = inputs.itemById('UGS_post').text
        UGS_platform = inputs.itemById('UGS_platform').value
        saveSettings = inputs.itemById('saveSettings').value
        showOperationsInput = inputs.itemById('showOperations')
        showOperations = showOperationsInput.selectedItem.name
        
        # Only attempt to get a value if the user has made a selection
        setupInput = inputs.itemById('setups')
        setupItem = setupInput.selectedItem
        if setupItem:
            setupName = setupItem.name
        
        folderInput = inputs.itemById('folders')
        folderItem = folderInput.selectedItem
        if folderItem:
            folderName = folderItem.name
        
        operationInput = inputs.itemById('operations')
        operationItem = operationInput.selectedItem
        if operationItem:
            operationName = operationItem.name
        
        # Get the name of setup, folder, or operation depending on radio selection
        # This is the operation that will post processed
        if (showOperations == 'Setups'):
            opName = setupName
        elif (showOperations == 'Folders'):
            opName = folderName
        elif (showOperations == 'Operations'):
            opName = operationName

        return (opName, UGS_path, UGS_post, UGS_platform, saveSettings, showOperations)

# Will update visibility of 3 selection dropdowns based on radio selection
# Also updates radio selection which is only really useful when command is first launched.
def setDropdown(inputs, showOperations):
    
    # Get input objects
    setupInput = inputs.itemById('setups')
    folderInput = inputs.itemById('folders')
    operationInput = inputs.itemById('operations')
    showOperationsInput = inputs.itemById('showOperations')

    # Set visibility based on appropriate selection from radio list
    if (showOperations == 'Setups'):
        setupInput.isVisible = True
        folderInput.isVisible = False
        operationInput.isVisible = False
        showOperationsInput.listItems[0].isSelected = True
    elif (showOperations == 'Folders'):
        setupInput.isVisible = False
        folderInput.isVisible = True
        operationInput.isVisible = False
        showOperationsInput.listItems[1].isSelected = True
    elif (showOperations == 'Operations'):
        setupInput.isVisible = False
        folderInput.isVisible = False 
        operationInput.isVisible = True
        showOperationsInput.listItems[2].isSelected = True
    else:
        # TODO add error check
        return
    return

# Define the event handler for when the command is executed 
class UGSExecutedEventHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
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
            app = adsk.core.Application.get()
            ui  = app.userInterface
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

# Define the event handler for when any input changes.
class UGSInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        try:
            # Get inputs and changed inputs
            input_changed = args.input
            inputs = args.inputs

            # Check to see if the post type has changed and show appropriate drop down
            if input_changed.id == 'showOperations':
                showOperations = input_changed.selectedItem.name
                setDropdown(inputs, showOperations)
                
        except:
            app = adsk.core.Application.get()
            ui  = app.userInterface
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
               
# Define the event handler for when the Octoprint command is run by the user.
class UGSCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        ui = []
        try:
            app = adsk.core.Application.get()
            ui  = app.userInterface
            doc = app.activeDocument
            products = doc.products
            product = products.itemByProductType('CAMProductType')
            
            # Check if the document has a CAMProductType. It will not if there are no CAM operations in it.
            if product == None:
                 ui.messageBox('There are no CAM operations in the active document')
                 return            
            # Cast the CAM product to a CAM object (a subtype of product).
            cam = adsk.cam.CAM.cast(product)

            # Setup Handlers and options for command
            cmd = args.command
            cmd.isExecutedWhenPreEmpted = False           
            
            onExecute = UGSExecutedEventHandler()
            cmd.execute.add(onExecute)
            handlers.append(onExecute)
            
            onInputChanged = UGSInputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            handlers.append(onInputChanged)

            # Define the inputs.
            inputs = cmd.commandInputs
            
            # Labels
            inputs.addTextBoxCommandInput('labelText2', '', '<a href="http://winder.github.io/ugs_website/">Universal Gcode Sender</a></span> A full featured gcode platform used for interfacing with advanced CNC controllers like GRBL and TinyG.', 4, True)
            inputs.addTextBoxCommandInput('labelText3', '', 'Choose the Setup or Operation to send to UGS', 2, True)
            
            # UGS local path and post information 
            UGS_path_input = inputs.addTextBoxCommandInput('UGS_path', 'UGS Path: ', 'Location of UGS', 1, False)
            UGS_post_input = inputs.addTextBoxCommandInput('UGS_post', 'Post to use: ', 'Name of post', 1, False)
            
            # Whether using classic or platform
            # TODO Could automate this based on path
            UGS_platform_input = inputs.addBoolValueInput("UGS_platform", 'Using UGS Platform?', True)
            
            # What to select from?  Setups, Folders, Operations?
            showOperations_input = inputs.addRadioButtonGroupCommandInput("showOperations", 'What to Post?')  
            radioButtonItems = showOperations_input.listItems
            radioButtonItems.add("Setups", False)
            radioButtonItems.add("Folders", False)
            radioButtonItems.add("Operations", False)

            # Drop down for Setups
            setupDropDown = inputs.addDropDownCommandInput('setups', 'Select Setup:', adsk.core.DropDownStyles.LabeledIconDropDownStyle)
            # Drop down for Folders
            folderDropDown = inputs.addDropDownCommandInput('folders', 'Select Folder:', adsk.core.DropDownStyles.LabeledIconDropDownStyle)
            # Drop down for Operations
            opDropDown = inputs.addDropDownCommandInput('operations', 'Select Operation:', adsk.core.DropDownStyles.LabeledIconDropDownStyle)
        
            # Populate values in dropdowns based on current document:
            for setup in cam.setups:
                setupDropDown.listItems.add(setup.name, False)
                for folder in setup.folders:
                    folderDropDown.listItems.add(folder.name, False)         
            for operation in cam.allOperations:
                opDropDown.listItems.add(operation.name, False)
                
            # Save user settings, values written to local computer XML file
            inputs.addBoolValueInput("saveSettings", 'Save settings?', True)
            
            # Defaults for command dialog
            cmd.commandCategoryName = 'UGS'
            cmd.setDialogInitialSize(500, 300)
            cmd.setDialogMinimumSize(500, 300)
            cmd.okButtonText = 'POST'  
            
            # Check if user has saved settings and update UI to reflect preferences
            xmlFileName = getFileName()
            if os.path.isfile(xmlFileName):
                
                # Read Settings                
                (UGS_path, UGS_post, UGS_platform, showOperations) = readSettings(xmlFileName)
                
                # Update dialog values
                UGS_path_input.text = UGS_path
                UGS_post_input.text = UGS_post
                UGS_platform_input.value = UGS_platform
                setDropdown(inputs, showOperations)
            else:
                setDropdown(inputs, 'Folders')

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
        solidPanel.controls.addCommand(UGSButtonDef, '', False)
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

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
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


