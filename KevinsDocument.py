from time import sleep

import adsk.cam
import adsk.core
import adsk.fusion
import traceback

MAX_TRIES = 5   # Maximum number of loops to attempt to get saved data file
SLEEP_TIME = 3  # Time to sleep between attempts


def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        doc_a = app.activeDocument
        doc_title = doc_a.name

        doc_a.activate()
        doc_b = app.documents.add(
            adsk.core.DocumentTypes.FusionDesignDocumentType)

        doc_b.saveAs("MFG Model - " + doc_title,
                     doc_a.dataFile.parentFolder, "Auto created by script", '')

        transform = adsk.core.Matrix3D.create()
        design_b = adsk.fusion.Design.cast(
            doc_b.products.itemByProductType("DesignProductType"))
        design_b.rootComponent.occurrences.addByInsert(
            doc_a.dataFile, transform, True)

        doc_b.save("Auto saved by script")

        # To go the other way you need to check for the completion of upload
        # doc_a.activate()
        # doc_a = app.documents.add(adsk.core.DocumentTypes.FusionDesignDocumentType)
        # design_a = adsk.fusion.Design.cast(doc_a.products.itemByProductType("DesignProductType"))
        # for i in range(MAX_TRIES):
        #     adsk.doEvents()
        #     try:
        #         data_file_b = doc_b.dataFile
        #         design_a.rootComponent.occurrences.addByInsert(data_file_b, transform, True)
        #         return
        #     except:
        #         sleep(SLEEP_TIME)
        #
        # ui.messageBox("Failed to create insert, try increasing sleep time or max tries")

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
