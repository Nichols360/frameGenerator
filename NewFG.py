import adsk.core
import adsk.fusion
from .lib import xlrd
import os

from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase
from .Fusion360Utilities.Fusion360Utilities import get_app_objects

_rowNumber = 0
count = 0
fHeight = 0
fWidth = 0
wallThick = 0
table_input = adsk.core.TableCommandInput.cast(None)
path = os.getenv('APPDATA')
loc = path + "\\Autodesk\\ApplicationPlugins\\Nichols360-FG.bundle\\Contents\\" + "Struct Table.xlsx"
wb = xlrd.open_workbook(loc)
sqsheet = wb.sheet_by_name('Ansi Square')
rectSheet = wb.sheet_by_name('Ansi Rectangular')
pipeSheet = wb.sheet_by_name('Ansi Pipe')


def tube_frame(tab1, inputs, input_values):
    global fHeight, fWidth, wallThick, sqsheet, rectSheet, pipeSheet
    ao = get_app_objects()
    # mem_select is now fam_select with the new changes.
    row_select = table_input.selectedRow
    fam_select = inputs.itemById('Family')

    size_select = inputs.itemById('Size')
    std_select = inputs.itemById('Standard')
    origin_select = inputs.itemById('OriginLoc' + str(row_select))

    if fam_select.selectedItem.name == 'Square':
        sizeSelect = size_select.selectedItem.name
        fHeight, fWidth, wallThick = sizeSelect.split("x")
        fHeight = float(fHeight) * 2.54
        fWidth = float(fWidth) * 2.54
        wallThick = float(wallThick) * 2.54
    elif fam_select.selectedItem.name == 'Rectangular':
        sizeSelect = size_select.selectedItem.name
        if rotation.selectedItem.name == 'HeightxWidth':
            fHeight, fWidth, wallThick = sizeSelect.split("x")
            fHeight = float(fHeight) * 2.54
            fWidth = float(fWidth) * 2.54
            wallThick = float(wallThick) * 2.54
        elif rotation.selectedItem.name == 'WidthxHeight':
            fWidth, fHeight, wallThick = sizeSelect.split("x")
            fHeight = float(fHeight) * 2.54
            fWidth = float(fWidth) * 2.54
            wallThick = float(wallThick) * 2.54

    elif fam_select.selectedItem.name == 'Pipe':
        sizeSelect = size_select.selectedItem.name
        fHeight, wallThick = sizeSelect.split("x")
        fHeight = float(fHeight) * 2.54
        wallThick = float(wallThick) * 2.54

    # product = ao.product
    app_objects = get_app_objects()
    design = app_objects['design']
    comp = app_objects['root_comp']
    geom_select = input_values['FSketchSel']
    frame_line = geom_select[row_select]
    ExtrusionLengthSelect = frame_line.length
    root_comp = app_objects['root_comp']
    planes = comp.constructionPlanes
    planeInput = planes.createInput()
    planeInput.setByDistanceOnPath(frame_line, adsk.core.ValueInput.createByReal(0))
    plane1 = planes.add(planeInput)
    sketches = root_comp.sketches
    sketch = sketches.add(plane1)
    lines = sketch.sketchCurves.sketchLines

    if origin_select.selectedItem.name == 'Lower Left':
        line = lines.addByTwoPoints(adsk.core.Point3D.create(0, 0, 0), adsk.core.Point3D.create(0, fWidth, 0))
        line1 = lines.addByTwoPoints(line.endSketchPoint, adsk.core.Point3D.create(fHeight, fWidth, 0))
        line2 = lines.addByTwoPoints(line1.endSketchPoint, adsk.core.Point3D.create(fHeight, 0, 0))
        line3 = lines.addByTwoPoints(line2.endSketchPoint, line.startSketchPoint)
        arc = sketch.sketchCurves.sketchArcs.addFillet(line, line.endSketchPoint.geometry, line1,
                                                       line1.startSketchPoint.geometry, (wallThick * 2))
        arc1 = sketch.sketchCurves.sketchArcs.addFillet(line1, line1.endSketchPoint.geometry, line2,
                                                        line2.startSketchPoint.geometry, (wallThick * 2))
        arc2 = sketch.sketchCurves.sketchArcs.addFillet(line2, line2.endSketchPoint.geometry, line3,
                                                        line3.startSketchPoint.geometry, (wallThick * 2))
        arc3 = sketch.sketchCurves.sketchArcs.addFillet(line3, line3.endSketchPoint.geometry, line,
                                                        line.startSketchPoint.geometry, (wallThick * 2))

        curves = sketch.findConnectedCurves(line)

        dirpoint = adsk.core.Point3D.create(0, 0, 0)
        offsetCurves = sketch.offset(curves, dirpoint, (-(wallThick)))

        WebOffset = fHeight - wallThick
        WebOffset2 = fWidth - wallThick
        profile = sketch.profiles.item(1)
        extent_distance = adsk.core.ValueInput.createByReal(ExtrusionLengthSelect)
        extrudes = root_comp.features.extrudeFeatures
        extrude1 = extrudes.addSimple(profile, extent_distance,
                                      adsk.fusion.FeatureOperations.NewComponentFeatureOperation)
    elif origin_select.selectedItem.name == 'Lower Right':
        line = lines.addByTwoPoints(adsk.core.Point3D.create(0, 0, 0), adsk.core.Point3D.create(0, (-(fWidth)), 0))
        line1 = lines.addByTwoPoints(line.endSketchPoint, adsk.core.Point3D.create(fHeight, (-(fWidth)), 0))
        line2 = lines.addByTwoPoints(line1.endSketchPoint, adsk.core.Point3D.create(fHeight, 0, 0))
        line3 = lines.addByTwoPoints(line2.endSketchPoint, line.startSketchPoint)

        arc = sketch.sketchCurves.sketchArcs.addFillet(line, line.endSketchPoint.geometry, line1,
                                                       line1.startSketchPoint.geometry, (wallThick * 2))
        arc1 = sketch.sketchCurves.sketchArcs.addFillet(line1, line1.endSketchPoint.geometry, line2,
                                                        line2.startSketchPoint.geometry, (wallThick * 2))
        arc2 = sketch.sketchCurves.sketchArcs.addFillet(line2, line2.endSketchPoint.geometry, line3,
                                                        line3.startSketchPoint.geometry, (wallThick * 2))
        arc3 = sketch.sketchCurves.sketchArcs.addFillet(line3, line3.endSketchPoint.geometry, line,
                                                        line.startSketchPoint.geometry, (wallThick * 2))

        curves = sketch.findConnectedCurves(line)

        dirpoint = adsk.core.Point3D.create(0, 0, 0)
        offsetCurves = sketch.offset(curves, dirpoint, (-(wallThick)))

        profile = sketch.profiles.item(0)
        extent_distance = adsk.core.ValueInput.createByReal(ExtrusionLengthSelect)
        extrudes = root_comp.features.extrudeFeatures
        extrude1 = extrudes.addSimple(profile, extent_distance,
                                      adsk.fusion.FeatureOperations.NewComponentFeatureOperation)
    elif origin_select.selectedItem.name == 'Upper Right':
        line = lines.addByTwoPoints(adsk.core.Point3D.create(0, 0, 0), adsk.core.Point3D.create(0, (-(fWidth)), 0))
        line1 = lines.addByTwoPoints(line.endSketchPoint, adsk.core.Point3D.create((-(fHeight)), (-(fWidth)), 0))
        line2 = lines.addByTwoPoints(line1.endSketchPoint, adsk.core.Point3D.create((-(fHeight)), 0, 0))
        line3 = lines.addByTwoPoints(line2.endSketchPoint, line.startSketchPoint)

        arc = sketch.sketchCurves.sketchArcs.addFillet(line, line.endSketchPoint.geometry, line1,
                                                       line1.startSketchPoint.geometry, (wallThick * 2))
        arc1 = sketch.sketchCurves.sketchArcs.addFillet(line1, line1.endSketchPoint.geometry, line2,
                                                        line2.startSketchPoint.geometry, (wallThick * 2))
        arc2 = sketch.sketchCurves.sketchArcs.addFillet(line2, line2.endSketchPoint.geometry, line3,
                                                        line3.startSketchPoint.geometry, (wallThick * 2))
        arc3 = sketch.sketchCurves.sketchArcs.addFillet(line3, line3.endSketchPoint.geometry, line,
                                                        line.startSketchPoint.geometry, (wallThick * 2))

        curves = sketch.findConnectedCurves(line)

        dirpoint = adsk.core.Point3D.create(0, 0, 0)
        offsetCurves = sketch.offset(curves, dirpoint, (-(wallThick)))

        profile = sketch.profiles.item(1)
        extent_distance = adsk.core.ValueInput.createByReal(ExtrusionLengthSelect)
        extrudes = root_comp.features.extrudeFeatures
        extrude1 = extrudes.addSimple(profile, extent_distance,
                                      adsk.fusion.FeatureOperations.NewComponentFeatureOperation)
    elif origin_select.selectedItem.name == 'Upper Left':
        line = lines.addByTwoPoints(adsk.core.Point3D.create(0, 0, 0), adsk.core.Point3D.create(0, fWidth, 0))
        line1 = lines.addByTwoPoints(line.endSketchPoint, adsk.core.Point3D.create((-(fHeight)), fWidth, 0))
        line2 = lines.addByTwoPoints(line1.endSketchPoint, adsk.core.Point3D.create((-(fHeight)), 0, 0))
        line3 = lines.addByTwoPoints(line2.endSketchPoint, line.startSketchPoint)

        arc = sketch.sketchCurves.sketchArcs.addFillet(line, line.endSketchPoint.geometry, line1,
                                                       line1.startSketchPoint.geometry, (wallThick * 2))
        arc1 = sketch.sketchCurves.sketchArcs.addFillet(line1, line1.endSketchPoint.geometry, line2,
                                                        line2.startSketchPoint.geometry, (wallThick * 2))
        arc2 = sketch.sketchCurves.sketchArcs.addFillet(line2, line2.endSketchPoint.geometry, line3,
                                                        line3.startSketchPoint.geometry, (wallThick * 2))
        arc3 = sketch.sketchCurves.sketchArcs.addFillet(line3, line3.endSketchPoint.geometry, line,
                                                        line.startSketchPoint.geometry, (wallThick * 2))

        curves = sketch.findConnectedCurves(line)

        dirpoint = adsk.core.Point3D.create(0, 0, 0)
        offsetCurves = sketch.offset(curves, dirpoint, (-(wallThick)))

        profile = sketch.profiles.item(0)
        extent_distance = adsk.core.ValueInput.createByReal(ExtrusionLengthSelect)
        extrudes = root_comp.features.extrudeFeatures
        extrude1 = extrudes.addSimple(profile, extent_distance,
                                      adsk.fusion.FeatureOperations.NewComponentFeatureOperation)
    elif origin_select.selectedItem.name == 'Center':
        line = lines.addByTwoPoints(adsk.core.Point3D.create((-fHeight/2), (-fWidth/2), 0),
                                    adsk.core.Point3D.create((-fHeight/2), (fWidth/2), 0))
        line1 = lines.addByTwoPoints(line.endSketchPoint, adsk.core.Point3D.create((fHeight/2), (fWidth/2), 0))
        line2 = lines.addByTwoPoints(line1.endSketchPoint, adsk.core.Point3D.create(((fHeight/2)), -fWidth/2, 0))
        line3 = lines.addByTwoPoints(line2.endSketchPoint, line.startSketchPoint)

        arc = sketch.sketchCurves.sketchArcs.addFillet(line, line.endSketchPoint.geometry, line1,
                                                       line1.startSketchPoint.geometry, (wallThick * 2))
        arc1 = sketch.sketchCurves.sketchArcs.addFillet(line1, line1.endSketchPoint.geometry, line2,
                                                        line2.startSketchPoint.geometry, (wallThick * 2))
        arc2 = sketch.sketchCurves.sketchArcs.addFillet(line2, line2.endSketchPoint.geometry, line3,
                                                        line3.startSketchPoint.geometry, (wallThick * 2))
        arc3 = sketch.sketchCurves.sketchArcs.addFillet(line3, line3.endSketchPoint.geometry, line,
                                                        line.startSketchPoint.geometry, (wallThick * 2))

        curves = sketch.findConnectedCurves(line)

        dirpoint = adsk.core.Point3D.create(0, 0, 0)
        offsetCurves = sketch.offset(curves, dirpoint, (-(wallThick)))

        profile = sketch.profiles.item(0)
        extent_distance = adsk.core.ValueInput.createByReal(ExtrusionLengthSelect)
        extrudes = root_comp.features.extrudeFeatures
        extrude1 = extrudes.addSimple(profile, extent_distance,
                                      adsk.fusion.FeatureOperations.NewComponentFeatureOperation)


def addRowToTable(tab1):
    global _rowNumber
    # Get the CommandInputs object associated with the parent command.
    cmdInputs = adsk.core.CommandInputs.cast(tab1.commandInputs)

    # Create three new command inputs.
    labelText = cmdInputs.addStringValueInput('label' + str(_rowNumber), 'Label ' + str(_rowNumber),
                                              'Member ' + str(_rowNumber))
    labelText.isReadOnly = True

    origin = cmdInputs.addDropDownCommandInput('OriginLoc' + str(_rowNumber), 'Origin Position',
                                               adsk.core.DropDownStyles.TextListDropDownStyle)
    origin.listItems.add('Quadrant', False)
    origin.listItems.add('Center', False)
    origin.listItems.add('Lower Left', True)
    origin.listItems.add('Lower Right', False)
    origin.listItems.add('Upper Left', False)
    origin.listItems.add('Upper Right', False)

    # Add the inputs to the table.
    row = tab1.rowCount
    tab1.addCommandInput(labelText, row, 0)
    tab1.addCommandInput(origin, row, 1)

    _rowNumber = _rowNumber + 1


class FrameGenerator(Fusion360CommandBase):
    # Run whenever a user makes any change to a value or selection in the addin UI
    # Commands in here will be run through the Fusion processor and changes will be reflected in  Fusion graphics area
    def on_preview(self, command: adsk.core.Command, inputs: adsk.core.CommandInputs, changed_input,
                   input_values):
        rows = tab1.rowCount
        table_input = inputs.itemById('MemberTable')
        row_select = table_input.selectedRow
        mem_select = inputs.itemById('MST' + str(row_select))
        size_select = inputs.itemById('SizeSelect' + str(row_select))
        geom_select = input_values['FSketchSel']
        i = 0
        while i < rows:
            table_input.selectedRow = i
            tube_frame(tab1, inputs, input_values)
            i += 1

    # Run after the command is finished.
    # Can be used to launch another command automatically or do other clean up.
    def on_destroy(self, command, inputs, reason, input_values):
        pass

    # Run when any input is changed.
    # Can be used to check a value and then update the add-in UI accordingly
    def on_input_changed(self, command: adsk.core.Command, inputs: adsk.core.CommandInputs, changed_input,
                         input_values):
        # eventArgs = adsk.core.InputChangedEventArgs.cast(args)
        # select = eventArgs.inputs
        # cmdInput = eventArgs.input
        global count, _rowNumber, table_input, loc
        geom_select = inputs.itemById('FSketchSel')
        if changed_input == inputs.itemById('FSketchSel'):
            IC = geom_select.selectionCount
            strIC = str(IC)

            for count in strIC:
                addRowToTable(tab1)
            if IC == 0:
                tab1.clear()
                _rowNumber = 0

        table_input = inputs.itemById('MemberTable')
        row_select = table_input.selectedRow
        # mat_select = inputs.itemById('Material')
        std_select = inputs.itemById('Standard')
        fam_select = inputs.itemById('Family')
        size_select = inputs.itemById('Size')

        if changed_input == fam_select:

            # sheet.cell_value(0, 0)
            if changed_input.selectedItem.name == 'Square':
                size_select.listItems.clear()
                wb.sheet_loaded('Ansi Square')
                for i in range(sqsheet.nrows):
                    sizeString = str(sqsheet.cell_value(i, 0)) + "x" + str(sqsheet.cell_value(i, 1)) + "x" + \
                                 str(sqsheet.cell_value(i, 2))
                    size_select.listItems.add(sizeString, False)
            elif changed_input.selectedItem.name == 'Rectangular':
                size_select.listItems.clear()
                wb.sheet_loaded('Ansi Rectangular')
                for i in range(rectSheet.nrows):
                    rect_string = str(rectSheet.cell_value(i, 0)) + "x" + str(rectSheet.cell_value(i, 1)) + "x" + \
                                  str(rectSheet.cell_value(i, 2))
                    size_select.listItems.add(rect_string, False)
            elif changed_input.selectedItem.name == 'Pipe':
                size_select.listItems.clear()
                wb.sheet_loaded('Ansi Pipe')
                for i in range(pipeSheet.nrows):
                    pipe_string = str(pipeSheet.cell_value(i, 0)) + "x" + str(pipeSheet.cell_value(i, 1))
                    size_select.listItems.add(pipe_string, False)

    # Run when the user presses OK
    # This is typically where your main program logic would go
    def on_execute(self, command: adsk.core.Command, inputs: adsk.core.CommandInputs, args, input_values):
        global tab1, _rowNumber, count
        rows = tab1.rowCount
        table_input = inputs.itemById('MemberTable')
        row_select = table_input.selectedRow
        mem_select = inputs.itemById('MST' + str(row_select))
        size_select = inputs.itemById('SizeSelect' + str(row_select))
        geom_select = input_values['FSketchSel']
        i = 0
        while i < rows:
            table_input.selectedRow = i
            tube_frame(tab1, inputs, input_values)
            i += 1

        _rowNumber = 0
        count = 0
        
    # Run when the user selects your command icon from the Fusion 360 UI
    # Typically used to create and display a command dialog box
    # The following is a basic sample of a dialog UI
    def on_create(self, command: adsk.core.Command, inputs: adsk.core.CommandInputs):

        global geo_select, tab1, trimmed_dist, rotation

        # Create a default value using a string
        default_value = adsk.core.ValueInput.createByString('1.0 in')
        ao = get_app_objects()

        # selection input for Frame skeleton
        geo_select = inputs.addSelectionInput('FSketchSel', 'Select Geometry', 'Basic select command input')
        geo_select.setSelectionLimits(1)
        geo_select.selectionFilters = ['SketchLines']

        standard_select = inputs.addDropDownCommandInput('Standard', 'Select Standard',
                                                         adsk.core.DropDownStyles.TextListDropDownStyle)
        standard_select.listItems.add('Ansi', True)
        standard_select.listItems.add('Iso', False)

        family_select = inputs.addDropDownCommandInput('Family', 'Select Family',
                                                       adsk.core.DropDownStyles.TextListDropDownStyle)
        family_select.listItems.add('Rectangular', False)
        family_select.listItems.add('Square', True)
        # family_select.listItems.add('Pipe', False)

        size_select = inputs.addDropDownCommandInput('Size', 'Select Size',
                                                     adsk.core.DropDownStyles.TextListDropDownStyle)
        wb.sheet_loaded('Ansi Square')
        for i in range(sqsheet.nrows):
            sizeString = str(sqsheet.cell_value(i, 0)) + "x" + str(sqsheet.cell_value(i, 1)) + "x" + \
                         str(sqsheet.cell_value(i, 2))
            size_select.listItems.add(sizeString, False)

        # material_select = inputs.addDropDownCommandInput('Material', 'Select Material',
        #                                                  adsk.core.DropDownStyles.TextListDropDownStyle)
        # material_select.listItems.add('Aluminum', True)
        # material_select.listItems.add('Cold Rolled Steel', False)
        # material_select.listItems.add('Hot Rolled Steel', False)
        # material_select.listItems.add('Stainless Steel', False)
        # material_select.listItems.add('Galvanized', False)
        # material_select.isFullWidth = False

        # table input for Selections
        tab1 = inputs.addTableCommandInput('MemberTable', 'Selected Lines', 4, '.75:2:1:1:1')
        # tab1.maximumVisibleRows = 8
        # addRowToTable(tab1)

        # Add inputs into the table.
        # addButtonInput = tab1.addBoolValueInput('tableAdd', 'Add', False, '', True)
        # tab1.addToolbarCommandInput(addButtonInput)
        # deleteButtonInput = tab1.addBoolValueInput('tableDelete', 'Delete', False, '', True)
        # tab1.addToolbarCommandInput(deleteButtonInput)
        # trimmed_dist = inputs.addValueInput('Trimmed', 'Trimmed Distance', 'in',
        #                                     adsk.core.ValueInput.createByString('2 in'))
        inputs.addTextBoxCommandInput('spacer_1', '', '<hr>', 1, True)
        inputs.addTextBoxCommandInput('lT10', '', 'This Section is for Orientation of the Member Row Selected', 1, True)

        rotation = inputs.addDropDownCommandInput('MemRotate', 'Rotation of Frame',
                                                  adsk.core.DropDownStyles.TextListDropDownStyle)
        rotation.listItems.add('HeightxWidth', True)
        rotation.listItems.add('WidthxHeight', False)
