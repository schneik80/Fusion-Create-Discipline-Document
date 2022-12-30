# Author-schneik.
# Description-This is add-in to take the active design document tab and insert into a new document as a child. the new document is named based on a user selection to help capture intent.

from pydoc import doc
import adsk.core
import adsk.fusion
import traceback
import os
import os.path
import gettext
import json

# Load project and folder from json
my_addin_path = os.path.dirname(os.path.realpath(__file__))
my_json_path = os.path.join(my_addin_path, "data.json")
with open(my_json_path) as json_file:
    data = json.load(json_file)
    print(data)

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

# handlers
handlers = []

# Support localization
_ = None


def getUserLanguage():
    app = adsk.core.Application.get()

    return {
        adsk.core.UserLanguages.ChinesePRCLanguage: "zh-CN",
        adsk.core.UserLanguages.ChineseTaiwanLanguage: "zh-TW",
        adsk.core.UserLanguages.CzechLanguage: "cs-CZ",
        adsk.core.UserLanguages.EnglishLanguage: "en-US",
        adsk.core.UserLanguages.FrenchLanguage: "fr-FR",
        adsk.core.UserLanguages.GermanLanguage: "de-DE",
        adsk.core.UserLanguages.HungarianLanguage: "hu-HU",
        adsk.core.UserLanguages.ItalianLanguage: "it-IT",
        adsk.core.UserLanguages.JapaneseLanguage: "ja-JP",
        adsk.core.UserLanguages.KoreanLanguage: "ko-KR",
        adsk.core.UserLanguages.PolishLanguage: "pl-PL",
        adsk.core.UserLanguages.PortugueseBrazilianLanguage: "pt-BR",
        adsk.core.UserLanguages.RussianLanguage: "ru-RU",
        adsk.core.UserLanguages.SpanishLanguage: "es-ES",
    }[app.preferences.generalPreferences.userLanguage]


# Get loc string by language
def getLocStrings():
    currentDir = os.path.dirname(os.path.realpath(__file__))
    return gettext.translation(
        "resource", currentDir, [getUserLanguage(), "en-US"]
    ).gettext


def commandDefinitionById(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox(_("commandDefinition id is not specified"))
        return None
    commandDefinitions_ = ui.commandDefinitions
    commandDefinition_ = commandDefinitions_.itemById(id)
    return commandDefinition_


def commandControlByIdForPanel(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox(_("commandControl id is not specified"))
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
            uiObj.messageBox(_(tobeDeleteObj + " is not a valid object"))


class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            cmdInput = args.input

            # we need to get the active document without the version at the end
            doc_a = app.activeDocument
            doc_with_ver = doc_a.name
            docSeed = doc_with_ver.rsplit(" ", 1)[0]  # trim version

            global doc_urn

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
                docTitle = (myDocsDict).get(listOfKeys).get(
                    "newDocTitle"
                ) + docSeed  # set the document title
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
                    _("Input changed event failed: {}").format(traceback.format_exc())
                )


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandEventArgs):
        global doc_urn
        try:
            command = args.firingEvent.sender

            sF = app.data.findFileById(doc_urn)
            doc_a = app.activeDocument
            docTitleinput: adsk.core.StringValueCommandInput = (
                args.command.commandInputs.itemById("stringValueInput_")
            )
            docTitle = docTitleinput.value
            doc_b = app.documents.open(sF)
            doc_b.saveAs(
                docTitle,
                doc_a.dataFile.parentFolder,
                "Auto created by related data add-in",
                "",
            )

            transform = adsk.core.Matrix3D.create()
            design_b = adsk.fusion.Design.cast(
                doc_b.products.itemByProductType("DesignProductType")
            )
            design_b.rootComponent.occurrences.addByInsert(
                doc_a.dataFile, transform, True
            )

            doc_b.save("Auto saved by related data add-in")

        except:
            if ui:
                ui.messageBox(
                    _("command executed failed: {}").format(traceback.format_exc())
                )


class CommandCreatedEventHandlerPanel(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            cmd = args.command
            cmd.helpFile = "help.html"
            global myDocsDict
            app = adsk.core.Application.get()
            ui = app.userInterface
            my_hub = app.data.activeHub
            my_project = my_hub.dataProjects.itemById(data["PROJECT_ID"])
            my_folder = my_project.rootFolder.dataFolders.itemById(data["FOLDER_ID"])
            myDocsDict = {}
            for data_file in my_folder.dataFiles:
                if data_file.fileExtension == "f3d":
                    dname = data_file.name + "dict"
                    myDocsDict.update(
                        {
                            dname: {
                                "name": data_file.name,
                                "urn": data_file.id,
                                "newDocTitle": data_file.name + " doc from ",
                            }
                        }
                    )
            ...

            onExecute = CommandExecuteHandler()
            cmd.execute.add(onExecute)

            onInputChanged = InputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            # keep the handler referenced beyond this function
            handlers.append(onExecute)
            handlers.append(onInputChanged)

            commandInputs_ = cmd.commandInputs

            # we need to get the active document without the version at the end
            doc_a = app.activeDocument
            doc_with_ver = doc_a.name
            docSeed = doc_with_ver.rsplit(" ", 1)[0]  # trim version

            dropDownCommandInput = commandInputs_.addDropDownCommandInput(
                "dropDownCommandInput",
                _("Type"),
                adsk.core.DropDownStyles.LabeledIconDropDownStyle,
            )

            dropDownItems_ = dropDownCommandInput.listItems
            for key, val in myDocsDict.items():
                if isinstance(val, dict):
                    dropDownItems_.add(_(val.get("name")), True),
                    docTitle = (val.get("newDocTitle")) + docSeed

            boolCommandInput = commandInputs_.addBoolValueInput(
                "boolvalueInput_", _("Auto-Name"), True
            )
            boolCommandInput.value = True

            stringDocName = commandInputs_.addStringValueInput(
                "stringValueInput_", _("Name"), _(docTitle)
            )
            stringDocName.isEnabled = False

        except:
            if ui:
                ui.messageBox(
                    _("Panel command created failed: {}").format(traceback.format_exc())
                )


def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        global _
        _ = getLocStrings()

        commandName = _(globalCommand)
        commandDescription = globalCommand
        commandResources = "./resources"
        iconResources = "./resources"

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
            handlers.append(onCommandCreated)
            toolbarControlPanel_ = toolbarControlsPanel_.addCommand(
                commandDefinitionPanel_
            )
            toolbarControlPanel_.isVisible = True

    except:
        if ui:
            ui.messageBox(
                _("Failed to start Add-In: {}").format(traceback.format_exc())
            )


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
            ui.messageBox(_("Failed to stop Add-In: {}").format(traceback.format_exc()))
