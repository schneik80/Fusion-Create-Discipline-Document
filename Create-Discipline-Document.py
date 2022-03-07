# Author-schneik.
# Description-This is add-in to take the active design document tab and insert into a new document as a child. the new document is named based on a user selection to help capture intent.

import adsk.core
import adsk.fusion
import traceback
import os
import gettext

# Global Command inputs
_app = adsk.core.Application.get()
ui = _app.userInterface
dropDownCommandInput = adsk.core.DropDownCommandInput.cast(None)
boolvalueInput = adsk.core.BoolValueCommandInput.cast(None)
stringDocname = adsk.core.StringValueCommandInput.cast(None)
commandIdOnPanel = "Create Discipline Document"
panelId = "SolidCreatePanel"
doc_seed = "Seed Document"
doc_title_ = "Document Title"
handlers = []

myDocsDict = {

    ### TODO Edit the below section adding a xxxDist section to the Documents Dictonary. This is the one place to setup your start documents. 

    "asmDict" : {
    "name": "Assembly",
    "urn": "urn:adsk.wipprod:dm.lineage:zAVmyja7TCyq_vmZ53Bg4g",
    "docTitle": "ASSY Doc from "
    },

    "mfgDict" : {
    "name": "Manufacturing",
    "urn" : "urn:adsk.wipprod:dm.lineage:g1eVUEzaQVqCCr9nGQxhFg",
    "docTitle" : "MFG Doc from "
    },

    "simDict" : {
    "name": "Simulation",
    "urn" : "urn:adsk.wipprod:dm.lineage:g1eVUEzaQVqCCr9nGQxhFg",
    "docTitle" : "SIM doc from "
    },

    "genDict" : {
    "name": "Generative",
    "urn" : "urn:adsk.wipprod:dm.lineage:fR9IDnS6S9uygX3Cukoc-A",
    "docTitle" : "GEN doc from "
    },

    "vizDict" : {
    "name": "Render",
    "urn" : "urn:adsk.wipprod:dm.lineage:GYrlC5yUQuWUlnp_1wp5wQ",
    "docTitle" : "VIZ doc from "
    },

    "anmDict" : {
    "name": "Animation",
    "urn" : "urn:adsk.wipprod:dm.lineage:OZOBQ0S0SMelun1F7rL8-A",
    "docTitle" : "ANM doc from "
    }

}

# doc_urn = (myDocsDict).get("asmDict").get("urn") # Assembly URN  # Start out with the assembly template urn

# Support localization
_ = None


def getUserLanguage():
    _app = adsk.core.Application.get()

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
    }[_app.preferences.generalPreferences.userLanguage]


# Get loc string by language


def getLocStrings():
    currentDir = os.path.dirname(os.path.realpath(__file__))
    return gettext.translation(
        "resource", currentDir, [getUserLanguage(), "en-US"]
    ).gettext


def commandDefinitionById(id):
    _app = adsk.core.Application.get()
    ui = _app.userInterface
    if not id:
        ui.messageBox(_("commandDefinition id is not specified"))
        return None
    commandDefinitions_ = ui.commandDefinitions
    commandDefinition_ = commandDefinitions_.itemById(id)
    return commandDefinition_


def commandControlByIdForPanel(id):
    _app = adsk.core.Application.get()
    ui = _app.userInterface
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
            uiObj.messageBox(_("tobeDeleteObj is not a valid object"))


class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            cmdInput = args.input
            # Get Document name
            doc_a = _app.activeDocument
            doc_seedv = doc_a.name
            doc_seed = doc_seedv.rsplit(" ", 1)[0]
            global doc_urn
            global myDocsDict
            stringDocname = args.inputs.itemById("stringValueInput_")
            
            if cmdInput.id == "dropDownCommandInput":
                if cmdInput.selectedItem.name == "Assembly":
                    doc_urn = (myDocsDict).get("asmDict").get("urn")  # Assembly URN
                    doc_title_ = "ASSY Doc from " + doc_seed
                    stringDocname.value = doc_title_
                    # ui.messageBox(_(doc_title_))
                if cmdInput.selectedItem.name == "Manufacturing":
                    doc_urn = "urn:adsk.wipprod:dm.lineage:yY2rAlHkRXuhiXTN1Q9P6Q"  # Manufacturing URN
                    doc_title_ = "MFG Doc from " + doc_seed
                    stringDocname.value = doc_title_
                    # ui.messageBox(_(doc_title_))

                if cmdInput.selectedItem.name == "Simulation":
                    doc_urn = "urn:adsk.wipprod:dm.lineage:g1eVUEzaQVqCCr9nGQxhFg"  # Simulation URN
                    doc_title_ = "SIM Doc from " + doc_seed
                    stringDocname.value = doc_title_
                    # ui.messageBox(_(doc_title_))

                if cmdInput.selectedItem.name == "Generative":
                    doc_urn = "urn:adsk.wipprod:dm.lineage:fR9IDnS6S9uygX3Cukoc-A"  # Generative URN
                    doc_title_ = "GEN Doc from " + doc_seed
                    stringDocname.value = doc_title_
                    # ui.messageBox(_(doc_title_))

                if cmdInput.selectedItem.name == "Render":
                    doc_urn = "urn:adsk.wipprod:dm.lineage:GYrlC5yUQuWUlnp_1wp5wQ"  # Rendering URN
                    doc_title_ = "VIZ Doc from " + doc_seed
                    stringDocname.value = doc_title_
                    # ui.messageBox(_(doc_title_))

                if cmdInput.selectedItem.name == "Animation":
                    doc_urn = "urn:adsk.wipprod:dm.lineage:OZOBQ0S0SMelun1F7rL8-A"  # Amimation URN
                    doc_title_ = "ANM Doc from " + doc_seed
                    stringDocname.value = doc_title_
                    # ui.messageBox(_(doc_title_))

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
        try:
            command = args.firingEvent.sender
            global doc_urn
            sF = _app.data.findFileById(doc_urn)
            doc_a = _app.activeDocument
            doc_title_input: adsk.core.StringValueCommandInput = (
                args.command.commandInputs.itemById("stringValueInput_")
            )
            doc_title_ = doc_title_input.value
            doc_b = _app.documents.open(sF)
            doc_b.saveAs(
                doc_title_,
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
            onExecute = CommandExecuteHandler()
            cmd.execute.add(onExecute)

            onInputChanged = InputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            # keep the handler referenced beyond this function
            handlers.append(onExecute)
            handlers.append(onInputChanged)

            commandInputs_ = cmd.commandInputs

             # we need to get the active document without the version at the end
            doc_a = _app.activeDocument
            doc_with_ver = doc_a.name
            doc_seed = doc_with_ver.rsplit(" ", 1)[0] # trim version

            dropDownCommandInput = commandInputs_.addDropDownCommandInput(
                "dropDownCommandInput",
                _("Type"),
                adsk.core.DropDownStyles.LabeledIconDropDownStyle,
            )


            dropDownItems_ = dropDownCommandInput.listItems
            for key, val in myDocsDict.items():
                    if isinstance(val, dict):
                        dropDownItems_.add(_(val.get("name")), True),
                        doc_title_ = (val.get("docTitle")) + doc_seed

            boolCommandInput = commandInputs_.addBoolValueInput(
                "boolvalueInput_", _("Auto-Name"), True
            )
            boolCommandInput.value = True

            stringDocName = commandInputs_.addStringValueInput(
                "stringValueInput_", _("Name"), _(doc_title_)
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
        _app = adsk.core.Application.get()
        ui = _app.userInterface

        global _
        _ = getLocStrings()

        commandName = _("Create Related Document")
        commandDescription = _("Create Related Document")
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
            ui.messageBox(_("AddIn Start Failed: {}").format(traceback.format_exc()))


def stop(context):
    ui = None
    try:
        _app = adsk.core.Application.get()
        ui = _app.userInterface
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
            ui.messageBox(_("AddIn Stop Failed: {}").format(traceback.format_exc()))
