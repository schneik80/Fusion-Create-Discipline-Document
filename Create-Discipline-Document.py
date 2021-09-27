# Author-schneik.
# Description-This is addin to take the active design document tab and insert into a new document as a child. the new document is named based on a user selection to help capture intent.

import adsk.core
import adsk.fusion
import traceback
import os
import gettext

# Globals
commandIdOnPanel = 'Create Discipline Document'
panelId = 'SolidCreatePanel'
doc_seed = 'seed'
doc_title_ = 'testing'

# Global Command inputs
_app = adsk.core.Application.get()
ui = _app.userInterface
dropDownCommandInput = adsk.core.DropDownCommandInput.cast(None)
boolvalueInput = adsk.core.BoolValueCommandInput.cast(None)
stringDocname = adsk.core.StringValueCommandInput.cast(None)

# global set of event handlers to keep them referenced for the duration of the command
handlers = []

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
            stringDocname = args.inputs.itemById('stringValueInput_')
            if cmdInput.id == 'dropDownCommandInput':
                if cmdInput.selectedItem.name == 'Assembly':
                    doc_title_ = "ASSY Doc from " + doc_seed
                    stringDocname.value = doc_title_
                    # ui.messageBox(_(doc_title_))
                if cmdInput.selectedItem.name == 'Manufacturing':
                    doc_title_ = "MFG Doc from " + doc_seed
                    stringDocname.value = doc_title_
                    # ui.messageBox(_(doc_title_))

            if cmdInput.id == 'boolvalueInput_':
                if cmdInput.value == True:
                    stringDocname.isEnabled = False
                else:
                    stringDocname.isEnabled = True

        except:
            if ui:
                ui.messageBox(
                    _("Input changed event failed: {}").format(
                        traceback.format_exc()
                    )
                )


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args: adsk.core.CommandEventArgs):
        try:
            command = args.firingEvent.sender
            doc_a = _app.activeDocument
            doc_title_input: adsk.core.StringValueCommandInput = args.command.commandInputs.itemById(
                "stringValueInput_")
            doc_title_ = doc_title_input.value
            doc_b = _app.documents.add(
                adsk.core.DocumentTypes.FusionDesignDocumentType
            )

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

            ui.messageBox(
                _("command: {} executed successfully").format(
                    command.parentCommandDefinition.id
                )
            )
        except:
            if ui:
                ui.messageBox(
                    _("command executed failed: {}").format(
                        traceback.format_exc()
                    )
                )


class CommandCreatedEventHandlerPanel(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            cmd = args.command
            cmd.helpFile = "help.html"

            onExecute = CommandExecuteHandler()
            cmd.execute.add(onExecute)

            onInputChanged = InputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            # keep the handler referenced beyond this function
            handlers.append(onExecute)
            handlers.append(onInputChanged)

            commandInputs_ = cmd.commandInputs

            dropDownCommandInput = commandInputs_.addDropDownCommandInput(
                'dropDownCommandInput', _('Type'), adsk.core.DropDownStyles.LabeledIconDropDownStyle)
            dropDownItems_ = dropDownCommandInput.listItems
            dropDownItems_.add(_("Assembly"), True)
            dropDownItems_.add(_("Manufacturing"), False)
            dropDownItems_.add(_("Simulation"), False)
            dropDownItems_.add(_("Gennerative"), False)
            dropDownItems_.add(_("Render"), False)
            dropDownItems_.add(_("Animation"), False)

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
                    _("Panel command created failed: {}").format(
                        traceback.format_exc()
                    )
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
            commandDefinitionPanel_ = commandDefinitions_.itemById(
                commandIdOnPanel)
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
            ui.messageBox(
                _(
                    "Command is successfully added to the create panel in modeling workspace"
                )
            )

    except:
        if ui:
            ui.messageBox(_("AddIn Start Failed: {}").format(
                traceback.format_exc()))


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
            ui.messageBox(_("AddIn Stop Failed: {}").format(
                traceback.format_exc()))
