# Author-schneik.
# Description-This is add-in to take the active design document tab and insert into a new document as a child. the new document is named based on a user selection to help capture intent.

# from pydoc import doc
import adsk.core
import adsk.fusion
import traceback
import os
import os.path
import json

# Global Command inputs
app = adsk.core.Application.get()
ui = app.userInterface
dropDownCommandInput = adsk.core.DropDownCommandInput.cast(None)
boolvalueInput = adsk.core.BoolValueCommandInput.cast(None)
stringDocname = adsk.core.StringValueCommandInput.cast(None)
globalCommand = " Create Related Document"
panelId = "SolidCreatePanel"
commandIdOnPanel = globalCommand
my_hub = app.data.activeHub

# create doc name values
docSeed = ""
docTitle = ""
docSeed = ""
myDocsDict = ()

# local_handlers
local_handlers = []


# Load project and folder from json
def loadProject(__file__):
    my_addin_path = os.path.dirname(os.path.realpath(__file__))
    my_json_path = os.path.join(my_addin_path, "data.json")
    global data, docSeed, docTitle, myDocsDict

    with open(my_json_path) as json_file:
        data = json.load(json_file)
        print(data)

    app = adsk.core.Application.get()
    # ui = app.userInterface
    my_hub = app.data.activeHub
    my_project = my_hub.dataProjects.itemById(data["PROJECT_ID"])
    if my_project is None:
        ui.messageBox(
            f"Project with id:{data['PROJECT_ID']} not found, review the readme file for instructions on how to set up the add-in."
        )
        return data
    my_folder = my_project.rootFolder.dataFolders.itemById(data["FOLDER_ID"])
    if my_folder is None:
        ui.messageBox(
            f"Folder with id:{data['FOLDER_ID']} not found, review the readme file for instructions on how to set up the add-in."
        )
        return data

    docActive = app.activeDocument
    doc_with_ver = docActive.name
    docSeed = doc_with_ver.rsplit(" ", 1)[0]  # trim version

    myDocsDictUnsorted = {}
    for data_file in my_folder.dataFiles:
        if data_file.fileExtension == "f3d":
            dname = data_file.name + "dict"
            myDocsDictUnsorted.update(
                {
                    dname: {
                        "name": data_file.name,
                        "urn": data_file.id,
                    }
                }
            )

    myDocsDict = dict(sorted(myDocsDictUnsorted.items()))
    print(myDocsDict)
    ...

    return data


data = loadProject(__file__)


def commandDefinitionById(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox("commandDefinition id is not specified")
        return None
    commandDefinitions_ = ui.commandDefinitions
    commandDefinition_ = commandDefinitions_.itemById(id)
    return commandDefinition_


def commandControlByIdForPanel(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox("commandControl id is not specified")
        return None
    workspaces_ = ui.workspaces
    modelingWorkspace_ = workspaces_.itemById("FusionSolidEnvironment")
    toolbarPanels_ = modelingWorkspace_.toolbarPanels
    toolbarPanel_ = toolbarPanels_.itemById(panelId)
    toolbarControls_ = toolbarPanel_.controls
    toolbarControl_ = toolbarControls_.itemById(id)
    return toolbarControl_


def destroyObject(uiObj, tobeDeleteObj):
    if uiObj and tobeDeleteObj:
        if tobeDeleteObj.isValid:
            tobeDeleteObj.deleteMe()
        else:
            uiObj.messageBox(tobeDeleteObj + " is not a valid object")


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandEventArgs):
        global doc_urn, docSeed, docTitle
        try:

            docActiveUrn = app.data.findFileById(doc_urn)
            docActive = app.activeDocument
            docTitleinput: adsk.core.StringValueCommandInput = (
                args.command.commandInputs.itemById("stringValueInput_")
            )
            docTitle = docTitleinput.value
            docNew = app.documents.open(docActiveUrn)
            docNew.saveAs(
                docTitle,
                docActive.dataFile.parentFolder,
                "Auto created by related data add-in",
                "",
            )
            transform = adsk.core.Matrix3D.create()
            seedDoc = adsk.fusion.Design.cast(
                docNew.products.itemByProductType("DesignProductType")
            )
            seedDoc.rootComponent.occurrences.addByInsert(
                docActive.dataFile, transform, True
            )

            docNew.save(
                "Auto saved by related data add-in"
            )  # Save new doc and add boiler plate comment

        #             doc_urn=""
        #             docSeed = ""
        #             docTitle=""

        except:
            if ui:
                ui.messageBox(
                    "command executed failed: {}".format(traceback.format_exc())
                )


class CommandCreatedEventHandlerPanel(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        global docSeed, doc_urn
        try:
            cmd = args.command
            cmd.helpFile = "help.html"

            onExecute = CommandExecuteHandler()
            cmd.execute.add(onExecute)

            onInputChanged = InputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            # keep the handler referenced beyond this function
            local_handlers.append(onExecute)
            local_handlers.append(onInputChanged)

            commandInputs_ = cmd.commandInputs

            dropDownCommandInput = commandInputs_.addDropDownCommandInput(
                "dropDownCommandInput",
                "Type",
                adsk.core.DropDownStyles.LabeledIconDropDownStyle,
            )

            dropDownItems_ = dropDownCommandInput.listItems
            for key, val in myDocsDict.items():
                if isinstance(val, dict):
                    dropDownItems_.add(val.get("name"), True),
                    docActive = app.activeDocument
                    doc_with_ver = docActive.name
                    docSeed = doc_with_ver.rsplit(" ", 1)[0]  # trim version
                    docTitle = docSeed + " -  - " + (val.get("name"))
                    doc_urn = val.get("urn")

            boolCommandInput = commandInputs_.addBoolValueInput(
                "boolvalueInput_", "Auto-Name", True
            )
            boolCommandInput.value = True

            stringDocName = commandInputs_.addStringValueInput(
                "stringValueInput_", "Name", docTitle
            )
            stringDocName.isEnabled = False

        except:
            if ui:
                ui.messageBox(
                    _("Panel command created failed: {}").format(traceback.format_exc())
                )


class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            cmdInput = args.input
            global doc_urn, docTitle, docSeed
            stringDocname = args.inputs.itemById("stringValueInput_")

            # handle the combobox change event
            if cmdInput.id == "dropDownCommandInput":
                searchDict = cmdInput.selectedItem.name

                # find the right dictionary based on the combo box value
                listOfKeys = ""
                for i in myDocsDict.keys():
                    for j in myDocsDict[i].values():
                        if searchDict in j:
                            if i not in listOfKeys:
                                listOfKeys = i

                doc_urn = (myDocsDict).get(listOfKeys).get("urn")  # set the urn
                doctempname = (myDocsDict).get(listOfKeys).get("name")
                docTitle = docSeed + " -  - " + doctempname
                stringDocname.value = docTitle

            # Auto name or user name input
            if cmdInput.id == "boolvalueInput_":
                if cmdInput.value == True:
                    stringDocname.isEnabled = False
                else:
                    stringDocname.isEnabled = True

        except:
            if ui:
                ui.messageBox(
                    "Input changed event failed: {}".format(traceback.format_exc())
                )


def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        commandName = globalCommand
        commandDescription = globalCommand
        commandResources = "./resources"
        # iconResources = "./resources"

        commandDefinitions_ = ui.commandDefinitions

        # add a command on create panel in modeling workspace
        workspaces_ = ui.workspaces
        modelingWorkspace_ = workspaces_.itemById("FusionSolidEnvironment")
        toolbarPanels_ = modelingWorkspace_.toolbarPanels

        # add the new command under the CREATE panel
        toolbarPanel_ = toolbarPanels_.itemById(panelId)
        toolbarControlsPanel_ = toolbarPanel_.controls
        toolbarControlPanel_ = toolbarControlsPanel_.itemById(commandIdOnPanel)
        if not toolbarControlPanel_:
            commandDefinitionPanel_ = commandDefinitions_.itemById(commandIdOnPanel)
            if not commandDefinitionPanel_:
                commandDefinitionPanel_ = commandDefinitions_.addButtonDefinition(
                    commandIdOnPanel, commandName, commandDescription, commandResources
                )
            onCommandCreated = CommandCreatedEventHandlerPanel()
            commandDefinitionPanel_.commandCreated.add(onCommandCreated)
            # keep the handler referenced beyond this function
            local_handlers.append(onCommandCreated)
            toolbarControlPanel_ = toolbarControlsPanel_.addCommand(
                commandDefinitionPanel_
            )
            toolbarControlPanel_.isVisible = True

    except:
        if ui:
            ui.messageBox("Failed to start Add-In: {}".format(traceback.format_exc()))


def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        objArrayPanel = []

        commandControlPanel_ = commandControlByIdForPanel(commandIdOnPanel)
        if commandControlPanel_:
            objArrayPanel.append(commandControlPanel_)

        commandDefinitionPanel_ = commandDefinitionById(commandIdOnPanel)
        if commandDefinitionPanel_:
            objArrayPanel.append(commandDefinitionPanel_)

        for obj in objArrayPanel:
            destroyObject(ui, obj)

    except:
        if ui:
            ui.messageBox("Failed to stop Add-In: {}").format(traceback.format_exc())
