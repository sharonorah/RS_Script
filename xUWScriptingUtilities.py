# -------------------------------------------------------------------------------
# Name:        ScriptingUtilities (v2.0)
#
# Originally Written for RS Version: 6.1.1.2
# Rewritten for RS Version: 8.1.1.2 - Not backwards compatible.
#
# Purpose:     This script contains utility functions that may be useful in a number of other scripts.
#
# Note:        Any modifications to this script should be tested against scripts importing it for compatability.
#              As of 04/15/2019 the following scripts import this script:
#                  -PlanCheck
#                  -NameBeams
#                  -SetOptParams
#                  -RenameCT
#                  -DoseSliceReport
#                  -PlanSetup
#                  -CouchAuto
#                  -ROISetup
#                  -AutoPlanWBRT
#
# Author:      LSW (UWMC) / WL (UWMC)
#
# Created:     23 May 2018 (v1.00)
#
# Updated:     25 June 2018 (v1.01)     - Wedge nameing bug fix.
#              10 July 2018 (v1.02)     - Updated dependencies to reflect use of this script by NameBeams script.
#              01 August 2018 (v1.03)   - Updated to incorporate function to set optimization parameters. Added some documentation.
#                                         Added functions to determine if the External ROI is contoured, or if an ROI with a given name is contoured.
#              13 August 2018 (v1.04)   - Updated set_opt_parameters(). Will now exit cleanly rather than crashing if plan is already approved. Sets the
#                                         number of segments to 10x the number of beams if there are beams. Fixed a bug causing an incorrect error message
#                                         if a beamset is specified that has no beams.
#              25 November 2018 (v1.05) - Incorporated CT renaming utility to facilitate plan check script check.
#                                         Added function to reorder beams with static gantry angles and an associated utility function to generate a relative area for a beam segment.
#              10 December 2018 (v1.06) - Pulled in slice report generator functions as part of an effort to generate composite dose slice reports.
#              22 January 2019 (v1.07)  - Pulled in functions for plan setup script.
#              31 January 2019 (v1.08)  - Minor edits to plan setup script functions.
#              01 February 2019 (v1.09) - Removed ui actions from ROI setup script.
#              15 April 2019 (v1.10)    - Updated Set Optimization Parameters function to work correctly for co-optimized beams.
#              23 September 2019 (v1.11)- Updated to function in RS 8.1.1.8. DLL imports updated with new names, 'Arc' delivery technique replaced by 'Static Arc' and 'Dynamic Arc'.
#                                         Code changes to future proof - e.g. all if, elif, ... else statements should evaluated before else, otherwise error message.
#              20 April 2020 (v1.12)    - Bug fix for function roi_contoured().
#              30 April 2020 (v1.13)    - Added code to calculate the wedged MU of a wedged beam.
#              11 May 2020 (v1.14)      - Bug fix (2) for function roi_countoured().
#              8/18/2023                - Updated statestree to 12A - SC
#              11/22/2023               - Updated create_doc(), define_styles(), generate_slice_report() to work with DoseSliceReport. -AJE
#              1/23/2024                - Updated - SC
#              4/12/2024                - Updated generate_slice_report(), add_section_with_image and added find_closest_z() to work with DoseSliceReport -SC. 
# -------------------------------------------------------------------------------

import string
from math import sin, cos, pi, tan, e
import sys
import clr
import subprocess
# import wpf
import os
from System import IO, Windows, DateTime
from System.Windows import MessageBox

clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import DialogResult, OpenFileDialog

script_path = IO.Path.GetDirectoryName(sys.argv[0])
path = script_path.rsplit('\\', 1)[0]
sys.path.append(path)

clr.AddReference("MigraDoc.DocumentObjectModel-WPF")
clr.AddReference("MigraDoc.Rendering-WPF")
clr.AddReference("PdfSharp-WPF")

from MigraDoc.DocumentObjectModel import Document, Colors, Section, Unit, ParagraphAlignment, \
    Paragraph
from MigraDoc.DocumentObjectModel.Tables import Table
from MigraDoc.DocumentObjectModel.Shapes import ShapePosition
from MigraDoc.Rendering import PdfDocumentRenderer
from PdfSharp import Pdf
from connect import get_current, CompositeAction


def max_leaf_travel_li(segments):
    """Determine the maximum distance traveled by any one MLC between each of the supplied segments, and return a list of these distances for each pair of consecutive segments."""
    #   print("pass1 in max_leaf")

    result = []
    for i in range(segments.Count - 1):
        max_difference = 0.0
        seg1 = segments[i]
        seg2 = segments[i + 1]
        for j in range(len(seg1.LeafPositions[0])):
            left_difference = abs(seg1.LeafPositions[0][j] - seg2.LeafPositions[0][j])
            right_difference = abs(seg1.LeafPositions[1][j] - seg2.LeafPositions[1][j])
            max_difference = max(max_difference, left_difference, right_difference)
        result.append(max_difference)
    #  print("pass in max_leaf 2")
    return result


def calc_time(beam_set):
    """Estimate the time required to deliver the beams in a beam set, from beam on for the first beam to beam off for the last, including mechanical setup between beams."""
    beam_on = 0.3  # Time after setup is complete for beam to turn on.
    gantry_rpm = 1
    leaf_speed = 2.2  # cm/s

    beams = beam_set.Beams
    delivery_time = 0.0
    if beams.Count > 0:
        if beams[0].DeliveryTechnique in ['Arc', 'StaticArc', 'DynamicArc']:
            for beam in beams:
                delivery_time += beam_on
                segments = beam.Segments
                beamMU = float(beam.BeamMU)
                for segment in segments:
                    if segment.RelativeWeight == 0:
                        continue
                    if float(segment.DoseRate) >= 1.1:  # Static field dose rates are equal to '1'.
                        doserate = float(segment.DoseRate / 60)  # MU/second
                    else:
                        doserate = 10  # Nominal 600 MU per minute.
                    MU = beamMU * float(segment.RelativeWeight)
                    delivery_time += MU / doserate

        elif beams[0].DeliveryTechnique == 'DMLC':
            return None  # No DMLC capable machines currently.

        elif beams[0].DeliveryTechnique == 'SMLC':
            for beam_index in range(beams.Count):
                if not beam_index == 0:
                    gantry_angle_difference = beams[beam_index].GantryAngle - beams[
                        beam_index - 1].GantryAngle
                    delivery_time += (gantry_angle_difference / 360.0) / (gantry_rpm * 60.0)

                segments = beams[beam_index].Segments
                for segment in segments:
                    delivery_time += beam_on
                    if float(segment.DoseRate) >= 1.1:  # Static field dose rates are equal to '1'.
                        doserate = float(segment.DoseRate / 60)  # MU/second
                    else:
                        doserate = 10  # Nominal 600 MU per minute.
                    delivery_time += beams[beam_index].BeamMU * float(
                        segment.RelativeWeight) / doserate
                if segments.Count > 1:
                    delivery_time += sum(max_leaf_travel_li(segments)) / leaf_speed
        else:
            raise ValueError('Unknown delivery technique: %s' % (beams[0].DeliveryTechnique))
    return delivery_time


def segment_area(segment):
    """Returns an approximate segment area."""
    numleaves = segment.LeafPositions[0].Count
    leafwidth = 1 * (numleaves == 40) + 0.5 * (numleaves == 80)
    bottomjaw = segment.JawPositions[3]
    topjaw = segment.JawPositions[2]
    start = int((bottomjaw + 20) / leafwidth)
    end = int((topjaw + 20) / leafwidth)
    if start > end:  # Some confusion regarding Raystation jaw labeling.
        start, end = end, start
    area = 0
    for i in range(start, end):
        print(i)
        area += segment.LeafPositions[1][i] - segment.LeafPositions[0][i]
    print(area)
    return area


def get_wedged_MU(beam):
    """
    Calculate the wedged MU for a Raystation beam object and return it.
    """
    machine = beam.MachineReference.MachineName
    energy = beam.MachineReference.Energy
    MU = beam.BeamMU
    wedge_angle = beam.Wedge.Angle
    to_rad = pi / 180

    mdb = get_current('MachineDB')
    machine = mdb.GetTreatmentMachine(machineName=machine)  # , lockMode = 'Read')
    beam_quality = [each for each in machine.PhotonBeamQualities if each.NominalEnergy == energy][0]
    q0 = beam_quality.BeamModels[0].BeamModel.MotorizedWedgeParameters.WedgeModulationParametersX[
        -1]
    p0 = beam_quality.BeamModels[0].BeamModel.MotorizedWedgeParameters.WedgeModulationParametersY[
        -1]

    tcax = e ** (tan(q0) * p0)
    v = wedge_angle * to_rad
    phi = 60 * to_rad
    ratio = tan(v) / (tan(phi) * tcax + tan(v) * (1 - tcax))

    return MU * ratio


def reorder_beamset(beamset):
    """Reorder the beams in the beamset, first by gantry angle (180.1 clockwise to 180.0) for couch = 0 beams, and then by gantry angle for increasing couch angle.
    For beams at the same angle, beams will be ordered from largest to smallest segment size. If there are multiple segments the first segment will be used for the determination."""
    if beamset.DeliveryTechnique in ['Arc', 'StaticArc', 'DynamicArc']:
        return False
    assert beamset.DeliveryTechnique in ['SMLC', 'DMLC'], 'Unkown delivery technique: %s' % (
        beamset.DeliveryTechnique)
    beamli = sorted(beamset.Beams, key=lambda x: x.Number)
    hassegments = True
    for each in beamli:
        try:
            each.Segments[0]
        except:
            hassegments = False
    if hassegments:
        beamli = sorted(beamli, key=lambda x: (
            x.CouchAngle + (x.CouchAngle < 180) * (x.CouchAngle != 0) * 360,
            # Sort couch angles, 0 first, and then everything else afterwards in a consistent direction.
            x.GantryAngle - 180.1 + 360 * (x.GantryAngle < 180.1),
            # Sort by Gantry next, starting with 180.1 in a clockwise direction facing the gantry
            -1 * segment_area(x.Segments[
                              0])))  # Sort by segment size if couch and gantry angle are the same. May be useful prior to merging SnS beams.
    else:
        beamli = sorted(beamli, key=lambda x: (
            x.CouchAngle + (x.CouchAngle < 180) * (x.CouchAngle != 0) * 360,
            # Sort couch angles, 0 first, and then everything else afterwards in a consistent direction.
            x.GantryAngle - 180.1 + 360 * (x.GantryAngle < 180.1),
            # Sort by Gantry next, starting with 180.1 in a clockwise direction facing the gantry
            1))  # Sort by segment size if couch and gantry angle are the same. May be useful prior to merging SnS beams.

    print([each.Name for each in beamli])
    oldbeamnums = [each.Number for each in beamli]
    tempadd = beamset.Beams.Count
    while True:
        tempbeamnums = [each + tempadd for each in oldbeamnums]
        if len(set(oldbeamnums).intersection(set(tempbeamnums))) > 0:
            tempadd += 1
        else:
            break
    for each in beamli:
        each.Number = each.Number + tempadd
    for i, each in enumerate(beamli):
        each.Number = i + 1
    return True


###########################
#                         #
#  Beam Naming Utilities  #
#                         #
###########################


def name_beam(gantry, couch, orientation):
    """Returns the base name for a beam given the gantry and couch angle, and the patient orientation. The function name_beam_standard is used to generate the name for a HFS patient,
       and this function accounts for patient orientation either using a lookup table for special cases (i.e. along patient axes) or by simply swapping R-L, S-I, A-P as appropriate for
       the patient orientation."""
    name = name_beam_standard(gantry, couch)
    if name in ['AP', 'L_Lat', 'PA', 'R_Lat', 'Inf', 'Vertex']:
        if orientation == 'FeetFirstSupine':
            name = {'AP': 'AP', 'L_Lat': 'R_Lat', 'PA': 'PA', 'R_Lat': 'L_Lat', 'Inf': 'Vertex',
                    'Vertex': 'Inf'}[name]
        if orientation == 'HeadFirstProne':
            name = {'AP': 'PA', 'L_Lat': 'R_Lat', 'PA': 'AP', 'R_Lat': 'L_Lat', 'Inf': 'Inf',
                    'Vertex': 'Vertex'}[name]
        if orientation == 'FeetFirstProne':
            name = {'AP': 'PA', 'L_Lat': 'L_Lat', 'PA': 'AP', 'R_Lat': 'R_Lat', 'Inf': 'Vertex',
                    'Vertex': 'Inf'}[name]
        if orientation == 'HeadFirstDecubitusRight':
            name = {'AP': 'L_Lat', 'L_Lat': 'PA', 'PA': 'R_Lat', 'R_Lat': 'AP', 'Inf': 'Inf',
                    'Vertex': 'Vertex'}[name]
        if orientation == 'FeetFirstDecubitusRight':
            name = {'AP': 'L_Lat', 'L_Lat': 'AP', 'PA': 'R_Lat', 'R_Lat': 'PA', 'Inf': 'Vertex',
                    'Vertex': 'Inf'}[name]
        if orientation == 'HeadFirstDecubitusLeft':
            name = {'AP': 'R_Lat', 'L_Lat': 'AP', 'PA': 'L_Lat', 'R_Lat': 'PA', 'Inf': 'Inf',
                    'Vertex': 'Vertex'}[name]
        if orientation == 'FeetFirstDecubitusRight':
            name = {'AP': 'R_Lat', 'L_Lat': 'PA', 'PA': 'L_Lat', 'R_Lat': 'AP', 'Inf': 'Vertex',
                    'Vertex': 'Inf'}[name]
    else:
        name = translate_position(name, orientation)
    return name


def name_beam_standard(gantry, couch):
    """Returns the base name for a beam for a head first supine patient given the gantry and couch angle according to the UW naming schema."""
    # Special Cases
    if couch % 90 == 0 and gantry % 90 == 0:
        if couch == 0:
            result = {0: 'AP', 90: 'L_Lat', 180: 'PA', 270: 'R_Lat', 360: 'AP'}[gantry]
        if couch == 270:
            result = {0: 'AP', 90: 'Vertex', 180: 'PA', 270: 'Inf', 360: 'AP'}[gantry]
        if couch == 90:
            result = {0: 'AP', 90: 'Inf', 180: 'PA', 270: 'Vertex', 360: 'AP'}[gantry]
        print('name_beam_standard result: ', result)
        return result
    # Determine gantry and couch grouping for naming pattern lookup.
    gantrygroup = [0 <= gantry <= 45, 45 < gantry <= 90, 90 < gantry < 135, 135 <= gantry < 180,
                   180 <= gantry <= 225, 225 < gantry <= 270, 270 < gantry < 315,
                   315 <= gantry <= 360]
    assert gantrygroup.count(True) == 1
    gantrygroup = gantrygroup.index(True)

    couchgroup = [315 < couch <= 360, 270 <= couch <= 315, 250 < couch < 270, 90 < couch < 110,
                  45 < couch <= 90, 0 <= couch <= 45]
    assert couchgroup.count(True) == 1
    couchgroup = couchgroup.index(True)

    patternlookup = [['ALS', 'ASL', 'ASR', 'AIR', 'AIL', 'ALI'],
                     ['LAS', 'SAL', 'SAR', 'IAR', 'IAL', 'LAI'],
                     ['LPS', 'SPL', 'SPR', 'IPR', 'IPL', 'LPI'],
                     ['PLS', 'PSL', 'PSR', 'PIR', 'PIL', 'PLI'],
                     ['PRI', 'PIR', 'PIL', 'PLS', 'PSR', 'PRS'],
                     ['RPI', 'IPR', 'IPL', 'LPS', 'SPR', 'RPS'],
                     ['RAI', 'IAR', 'IAL', 'LAS', 'SAR', 'RAS'],
                     ['ARI', 'AIR', 'AIL', 'ALS', 'ASR', 'ARS']]
    # Usage: patternlookup[gantrygoup][couchgroup]

    pattern = patternlookup[gantrygroup][couchgroup]
    gantrynum = abs(
        min(gantry % 90, 90 - gantry % 90))  # Deviation from nearest cardinal angle for gantry...
    couchnum = abs(min(couch % 90, 90 - couch % 90))  # and couch.

    result = pattern[0]
    if gantrynum != 0:
        result += str(gantrynum) + pattern[1]
    if couchnum != 0:
        result += str(couchnum) + pattern[2]
    return result


def translate_position(name, orientation):
    """Helper function to translate string from one patient orientation to another."""
    if orientation == 'HeadFirstSupine':
        pass
    elif orientation == 'FeetFirstSupine':
        intab = 'APRLSI'
        outtab = 'APLRIS'
        name = name.replace(intab, outtab)
    elif orientation == 'HeadFirstProne':
        intab = 'APRLSI'
        outtab = 'PALRSI'
        name = name.replace(intab, outtab)
    elif orientation == 'FeetFirstProne':
        intab = 'APRLSI'
        outtab = 'PARLIS'
        name = name.replace(intab, outtab)
    elif orientation == 'HeadFirstDecubitusRight':
        intab = 'APRLSI'
        outtab = 'LRAPSI'
        name = name.replace(intab, outtab)
    elif orientation == 'FeetFirstDecubitusRight':
        intab = 'APRLSI'
        outtab = 'LRPAIS'
        name = name.replace(intab, outtab)
    elif orientation == 'HeadFirstDecubitusLeft':
        intab = 'APRLSI'
        outtab = 'RLPASI'
        name = name.replace(intab, outtab)
    elif orientation == 'FeetFirstDecubitusLeft':
        intab = 'APRLSI'
        outtab = 'RLAPIS'
        name = name.replace(intab, outtab)
    else:
        print("Unknown patient orientation, result will likely be incorrect.")
    return name


def get_wedge_orientation(coll_rot, gantry_rot, couch_rot, pat_orientation):
    """Determine the anatomical orientation of the wedge heel for the given machine settings and patient orientation."""
    # Patient vectors, outward from patient (i.e. ant points upwards for supine patient in cartesian coordinates)
    left = [1, 0, 0]
    sup = [0, 1, 0]
    ant = [0, 0, 1]

    patient = [left, sup, ant]

    # Gantry vectors, first is beam direction (i.e. from source outwards along beam path), and the second is the
    # vector that, when crossed with the first vector, gives the wedge orientation (from toe to heel).
    g1 = [0, 0, -1]
    g2 = [1, 0, 0]
    wedge = [g1, g2]

    wedge = [rot_vect(each, 'z', coll_rot) for each in
             wedge]  # Apply collimator rotation to the wedge.
    wedge = [rot_vect(each, 'y', gantry_rot) for each in
             wedge]  # Apply gantry rotation to the wedge.
    wedge = cp(wedge[0], wedge[1])

    patient = [rot_vect(each, 'z', couch_rot) for each in
               patient]  # Apply couch rotation to the patient.

    result = [dp(wedge, each) for each in
              patient]  # Take the dot product of the wedge orientation with the patient axes to determine which the wedge orientation most overlaps with.

    result = [(abs(result[0]), result[0], ['HL', 'HR']), (abs(result[1]), result[1], ['HS', 'HI']),
              (abs(result[2]), result[2], ['HA',
                                           'HP'])]  # Build up a list with results, absolute magnitudes, and appropriate names for each. [Positive,Negative]

    result = sorted(result,
                    key=lambda x: x[0])  # Sort, pick out correct axis, determine correct name.
    result = result[-1]
    result = result[2][result[1] < 0]

    # Translate for actual patient orientation.
    result = translate_position(result, pat_orientation)
    print('Get Wedge Orientation Result: ', result)

    return result


##########################
#                        #
# Mathematical Utilities #
#                        #
##########################

def cp(v1, v2):
    """Returns the cross product of two vectors supplied as lists."""
    return [v1[1] * v2[2] - v1[2] * v2[1],
            v1[2] * v2[0] - v1[0] * v2[2],
            v1[0] * v2[1] - v1[1] * v2[0]]


def dp(v1, v2):
    """Returns the dot product of two vectors supplied as lists."""
    return v1[0] * v2[0] + v1[1] * v2[1] + v1[2] * v2[2]


def rot_vect(v1, axis, theta):
    """Rotate a vector along the specified axis ('x','y','z') by an amount theta (in degrees). The vector is supplied as a list. Rotation direction is
       based on standard 3D Cartesian space (https://en.wikipedia.org/wiki/Cartesian_coordinate_system) and the right hand rule. For clarity, for a head first
       supine patient Cartesian x, y, and z correspond to patient right to left, inf to sup, and post to ant respectively.

       v1: A list of specifying the vector to be rotated. [x,y,z]
       axis: A string specifing the axis to rotate about. 'x'/'y'/'z'
       theta: The magnitude of rotation in degrees."""
    theta = pi / 180 * theta
    if axis == None:
        return v1
    if axis == 'x':
        r = [[1, 0, 0], [0, cos(theta), -1 * sin(theta)], [0, sin(theta), cos(theta)]]
    if axis == 'y':
        r = [[cos(theta), 0, sin(theta)], [0, 1, 0], [-1 * sin(theta), 0, cos(theta)]]
    if axis == 'z':
        r = [[cos(theta), -1 * sin(theta), 0], [sin(theta), cos(theta), 0], [0, 0, 1]]
    result = []
    for a, b in zip([v1, v1, v1], r):
        result.append(dp(a, b))
    return [round(each, 4) for each in result]


def cartesian_to_dicom(pt_li, pat_orientation):
    """Take a list of points in standard Cartesian space (equivalent to Raystation internal coordinates, where x is positive towards patient left for a HFS patient,
       y is towards patient superior, and z is towards patient anterior), and translate them to DICOM coordinates for the specified patient orientation.
       The points should be supplied as a list of lists, with each sublist having 3 entries corresponding to x,y,z.

       pt-li: A list of 3-entry x,y,z lists. E.g. [[x0,y0,z0],[x1,y1,z1],...[xn,yn,zn]]
       pat_orientation: A string specifying the patient orientation. {HeadFirst/FeetFirst}{Supine/Prone/DecubitusRight/DecubitusLeft}"""
    result = []
    if type(pt_li[0]) == type(float()) or type(pt_li[0]) == type(int()) and len(pt_li) == 3:
        pt_li = list(pt_li)
    for each in pt_li:
        assert len(each) == 3
    for each in pt_li:
        x, y, z = each[0], each[1], each[2]
        if pat_orientation == 'HeadFirstSupine':
            x, y, z = x, -1 * z, y
        if pat_orientation == 'FeetFirstSupine':
            x, y, z = -1 * x, -1 * z, -1 * y
        if pat_orientation == 'HeadFirstProne':
            x, y, z = -1 * x, z, y
        if pat_orientation == 'FeetFirstProne':
            x, y, z = x, z, -1 * y
        if pat_orientation == 'HeadFirstDecubitusRight':
            x, y, z = -1 * z, -1 * x, y
        if pat_orientation == 'FeetFirstDecubitusRight':
            x, y, z = z, -1 * x, -1 * y
        if pat_orientation == 'HeadFirstDecubitusLeft':
            x, y, z = z, x, y
        if pat_orientation == 'FeetFirstDecubitusRight':
            x, y, z = -1 * z, x, -1 * y
        result.append([x, y, z])
    return result


#################################
#                               #
#  Report Generation Functions  #
#                               #
#################################

def create_doc():
    """ This function creates a MigraDoc document object, and r eturns the created document object"""
    doc = Document()
    define_styles(doc)
    sec = doc.AddSection()
    table = Table()
    table.Borders.Width.Unit = 0.0
    col = table.AddColumn(Unit.FromCentimeter(6))
    col.Format.Alignment = ParagraphAlignment.Left
    col = table.AddColumn(Unit.FromCentimeter(5))
    col.Format.Alignment = ParagraphAlignment.Left
    row = table.AddRow()
    sec.Add(table)
    return doc


def define_styles(document):
    """ Function that defines the text styles for the MigraDoc document object 'document'
        Returns nothing"""

    for name, size, bold, fontname, color, spacebefore, spaceafter in zip(
            ["Normal", "Heading1", "Heading2", "Heading3"], [None, 14., 12., 20.],
            [None, True, True, True],
            ["Verdana", None, None, "Verdana"], [None, Colors.DarkBlue, None, Colors.DeepSkyBlue],
            [None, 15., 4., None], [4, None, None, 10.]):
        style = document.Styles[name]
        if size is not None:
            style.Font.Size.Unit = size
        if bold is not None:
            style.Font.Bold = bold
        if fontname is not None:
            style.Font.Name = fontname
        if color is not None:
            style.Font.Color = color
        if spacebefore is not None:
            style.ParagraphFormat.SpaceBefore.Unit = spacebefore
        if spaceafter is not None:
            style.ParagraphFormat.SpaceAfter.Unit = spaceafter

    style = document.Styles.AddStyle("NO style", "Normal")
    style.Font.Bold = True
    style.Font.Color = Colors.Red

    style = document.Styles.AddStyle("YES style", "Normal")
    style.Font.Bold = True
    style.Font.Color = Colors.Green

    style = document.Styles.AddStyle("Image style", "Normal")
    style.ParagraphFormat.SpaceAfter.Unit = 1.


def create_doc_file(doc, filename):
    """ Function that takes a MigraDoc document object 'doc' and a string 'filename' and creates a pdf-file named 'filename'
        from the document object 'doc' and displays it
        Returns nothing"""
    renderer = PdfDocumentRenderer(True, Pdf.PdfFontEmbedding.Always)
    renderer.Document = doc
    renderer.RenderDocument()
    renderer.PdfDocument.Save(filename)


def display_doc_file(filename):
    subprocess.call(filename, shell=True)


def add_section_with_image(document, image_files, square, first, description=[], data=[], title=None):
    """ Function that adds a section with images to the document. """
    
    # Widths in centimeters
    width_cm = ['', 17, 9, 6, 4]
    scale_width = ['', 0.8, 0.4, 0.25, 0.16]

    sec = document.LastSection
    if not first:
        sec.AddPageBreak()

    if title:
        sec.AddParagraph(title, "Heading2")

    # Text table setup
    table = Table()
    table.Borders.Width = Unit.FromPoint(0.0)
    col = table.AddColumn(Unit.FromCentimeter(6))
    col.Format.Alignment = ParagraphAlignment.Left
    col = table.AddColumn(Unit.FromCentimeter(5))
    col.Format.Alignment = ParagraphAlignment.Left

    for desc, dat in zip(description, data):
        row = table.AddRow()
        row.Cells[0].AddParagraph(desc)
        row.Cells[0].Style = "Image style"
        row.Cells[1].AddParagraph(dat)
        row.Cells[1].Style = "Image style"

    sec.Add(table)

    # Image table setup
    imgtable = Table()
    for i in range(square):
        col = imgtable.AddColumn()
        col.Format.Alignment = ParagraphAlignment.Center
        col.Width = Unit.FromCentimeter(width_cm[square])  # Convert string to Unit

    imgindex = 0
    for thisimage in image_files:
        if imgindex % square == 0:
            row = imgtable.AddRow()
            row.Height = Unit.FromCentimeter(width_cm[square])  # Convert string to Unit
            para = row.Cells[0].AddParagraph()
            img = para.AddImage(thisimage)
            imgindex = 0
        else:
            para = row.Cells[imgindex].AddParagraph()
            img = para.AddImage(thisimage)

        imgindex += 1
        img.LockAspectRatio = True
        img.ScaleWidth = scale_width[square]

    line_break = Paragraph()
    line_break.AddLineBreak()
    sec.Add(line_break)
    sec.Add(imgtable)


def find_closest_z(z_value, points):
    """Find the closest z value in points to the given z_value."""
    return min(points, key=lambda point: abs(point['z'] - z_value))

def generate_slice_report(numcol = 1, printevery = 2, printreverse = True, startstopfocus = None, maxdose = None):
    for each in dir(IO):
        print(each)
    print(help(IO))
    # valid number of columns = 1,2,3,4
    assert numcol in [1,2,3,4]
    number_of_columns = numcol
    
    # Print every _____ image, 2 = sencond, 3 = third, etc... 1 = all
    print_every = printevery
    
    # reverse print order
    printReversed = printreverse
    
    patient = get_current("Patient")
    plan = get_current("Plan")
    ui = get_current('ui')
    examination = get_current("Examination")
    bs = get_current("BeamSet")
    
    version = int(ui.GetApplicationVersion()[0])
    
    relative_slice_positions = examination.Series[0].ImageStack.SlicePositions
    absolute_start_slice_position = examination.Series[0].ImageStack.Corner.z
    
    absolute_slice_positions = []
    for relative_slice_position in relative_slice_positions:
        absolute_slice_positions.append(absolute_start_slice_position + relative_slice_position)
    
    # establish start and stop z coordinates from POIs
    # alternatively start_z and stop_z could be taken from the isocenter.z plus minus some distance
        
    if startstopfocus is None:
        try:
            pois = [each for each in patient.PatientModel.StructureSets[examination.Name].PoiGeometries if each.OfPoi.Name.strip().lower() in ['start','stop']]
            if len(pois) > 2:
                pois = [each for each in patient.PatientModel.StructureSets[examination.Name].PoiGeometries if each.OfPoi.Name in ['START','STOP']]
            for poi in pois:
                if poi.OfPoi.Name.strip().lower() == "start":
                    start_z = poi.Point.z
                focus_x = 0
                focus_y = 0
                if poi.OfPoi.Name.strip().lower() == "stop":
                    stop_z = poi.Point.z
                isocenter = bs.Beams[0].PatientToBeamMapping.IsocenterPoint
        except:
            case = get_current('Case')
            pois = [each for each in case.PatientModel.StructureSets[examination.Name].PoiGeometries if each.OfPoi.Name.strip().lower() in ['start','stop']]
            if len(pois) > 2:
                pois = [each for each in case.PatientModel.StructureSets[examination.Name].PoiGeometries if each.OfPoi.Name in ['START','STOP']]
            for poi in case.PatientModel.StructureSets[examination.Name].PoiGeometries:
                if poi.OfPoi.Name.strip().lower() == "start":
                    start_z = poi.Point.z
                focus_x = 0
                focus_y = 0
                if poi.OfPoi.Name.strip().lower() == "stop":
                    stop_z = poi.Point.z
    
        if start_z == None or stop_z == None:
            # couldn't find start and stop POIs
            Windows.MessageBox.Show("START and STOP POIs not found")
            sys.exit()
        
        # make sure that start < stop
        if start_z > stop_z:
            start_z,stop_z = stop_z,start_z
        
        #isocenter = bs.Beams[0].PatientToBeamMapping.IsocenterPoint
        
        orientations = []
        points = []
        focus = []
        
        # include images at least some distance apart (and the last one)
        # minimum_distance = 0.145 #cm
        print('CT Slices: ',len(absolute_slice_positions))
        for absolute_slice_position in absolute_slice_positions:
            if absolute_slice_position >= start_z and absolute_slice_position <= stop_z:
                ## if absolute_slice_position == start_z or abs(points[points.Count-1]['z'] - absolute_slice_position) > minimum_distance or absolute_slice_position == stop_z:
                if absolute_slice_position >= start_z or absolute_slice_position <= stop_z:
                    points.append({'x': focus_x, 'y': focus_y, 'z': absolute_slice_position})
                    orientations.append("Transversal")
                    focus.append(True)
        print('CT Slices used for Report: ',len(points))
   
    else:
        orientations = []
        points = []
        focus = []
        print('CT Slices: ',len(absolute_slice_positions))
        for absolute_slice_position in absolute_slice_positions:
            if True in [(absolute_slice_position >= each[0]) and (absolute_slice_position <= each[1]) for each in startstopfocus]:
                index = [absolute_slice_position >= each[0] and absolute_slice_position <= each[1] for each in startstopfocus].index(True)
                points.append({'x': startstopfocus[index][2], 'y': startstopfocus[index][3], 'z':absolute_slice_position})
                orientations.append("Transversal")
                focus.append(True)
        print('CT Slices used for Report: ',len(points))
            
    
    print("Creating images")
    GDIParams = {
        "Orientations":orientations,
        "Points":points,
        "FocusOnIsocenter":focus,
        "ImageSize":{'x': 800, 'y': 800},
        "FocusOnRoi":None
    }
    #if version > 5:
     #   GDIParams["FocusOnRoi"] = None
    # images = dict( bs.GetDoseImages(**GDIParams) )
    images = bs.GetDoseImages(**GDIParams) 
    
    doc = create_doc()
    # sorted_positions = []
    # for image in images:
    #     sorted_positions.append(image['z'])
    # sorted_positions = sorted(sorted_positions, key=lambda i: i, reverse=printReversed)
    
    # Map z positions to image paths
    z_to_image_path = {point['z']: img_path for point, img_path in zip(points, images)}

    # Sort positions
    
    sorted_positions = [point['z'] for point in points]
    sorted_positions = sorted(sorted_positions, reverse=printReversed)
    
    if maxdose is not None:
        GDIParams = {
            "Orientations":['Transversal'],
            "Points":[{'x':maxdose[1],'y':maxdose[2],'z':maxdose[3]}],
            "FocusOnIsocenter":[True],
            "ImageSize":{'x':800,'y':800},
            "FocusOnRoi":None}
        maxdoseimage = bs.GetDoseImages(**GDIParams)
        
   
        
    print("Building report")
    imagegroup = []
    igindex = 0
    totindex = 0
    first = True
    # Add images to the report
        
    if maxdose is not None:
        closest_z = find_closest_z(maxdose[3], points)
        closest_image_path = z_to_image_path[closest_z['z']]
        add_section_with_image(doc, [closest_image_path], 1, True, title='Max Dose: %i cGy' % maxdose[0])
    
    for position in sorted_positions:
        image_path = z_to_image_path[position]
        if totindex % print_every == 0:
            imagegroup.append(image_path)
            igindex += 1
            if igindex == numcol**2:
                add_section_with_image(doc, imagegroup, numcol, first)
                first = False
                igindex = 0
                imagegroup[:] = []
        totindex += 1
    
    if imagegroup:
        add_section_with_image(doc, imagegroup, numcol, first)
    
    print("Showing report")
    # note \\viptier1\radonc is mapped on most PCs as P:
    try:
        output_directory = r"R:/Slice_reports"
        output_filename ="Slice report, "+patient.Name+".pdf"
        for each in  ['<','>',':','"','/','|','?','*',' ',',']:
            output_filename = output_filename.replace(each,'')
        
        if not IO.Directory.Exists(output_directory):
            IO.Directory.CreateDirectory(output_directory)
        with open(output_directory + "\\test.txt",'w') as f: #Test to prevent creating a Migradoc document.
            f.write('test')
            f.close()
        
    except Exception as e:
        print("Could not generate report directly, trying remote connection routine...")
        print("Error message: ",e)
        try:
            success = True
            dialog = OpenFileDialog()
            dialog.Title = r"R:/Slice_reports"
            if dialog.ShowDialog() != DialogResult.OK:
                success = False
            if success:
                output_directory = dialog.FileName[0:dialog.FileName.rindex('\\')]    
                output_filename ="Slice report, "+patient.PatientName+".pdf"
                if not IO.Directory.Exists(output_directory):
                    IO.Directory.CreateDirectory(output_directory)
            
                for each in ['<','>',':','"','/','|','?','*',' ',',']:
                    output_filename = output_filename.replace(each,'')
        except Exception as e:
            print('Failed to generate report using remote routine.')
            print('Error message: ',e)
            
    try:
        create_doc_file(doc, output_directory + "\\" + output_filename)
    except Exception as e:
        print('Could not generate report pdf.')
        print('Error message: ',e)
        
    try:
        display_doc_file(output_directory + '\\' + output_filename)
    except Exception as e:
        print('Could not display pdf.')
        print('Error message: ',e)
        print('Filename:', output_directory + '\\' + output_filename)
            
    print("Removing images")
    for filename in images:
        try:
            IO.File.Delete(filename)
        except Exception as e:
            print(f'Could not delete image file {filename}. Error: {e}')




#####################
#                   #
#  Other Functions  #
#                   #
#####################

def set_opt_parameters(plan=None, beam_set=None, MaxNumberOfIterations=80, ComputeFinalDose=True,
                       ComputeIntermediateDose=None, IterationsInPreparationsPhase=20,
                       MaxNumberOfSegments=None, MinSegmentArea=4, MinSegmentMUPerFraction=6,
                       MinNumberOfOpenLeafPairs=2, MinLeafEndSeparation=2, MaxLeafSpeed=1.2,
                       popup=False):
    """Sets the optimization parameters for the specified plan/beamset to the specified values.
       If the plan AND beamset are not both provided, the currently used plan and beamset is used (to ensure the beamset is part of the plan)
       If a plan and beamset are both provided, a check will be performed to ensure the beamset is part of the plan.
       If one of the beamsets is co-optimized, the parameters will be updated for all beamsets.

        plan - The plan containing the beamset to set parameters for.
        beam_set - The beam set to set parameters for.
        MaxNumberOfIterations - An integer with the maximum number of iterations to be used. Defaults to 80.
        ComputeFinalDose - Boolean. Defaults to True.
        ComputeIntermediateDose - Boolean. Defaults to True for IMRT, and False for VMAT.
        IterationsInPreparationsPhase - Integer. The number of iterations before conversion. Defaults to 20.
        MaxNumberOfSegments. - Integer. Defaults to 50.
        MinSegmentArea - Float. Defaults to 4.
        MinSegmentMUPerFraction - Float. Defaults to 6.
        MinNumberOfOpenLeafPairs - Integer. Defaults to 2.
        MinLeafEndSeparation - Integer. Defaults to 2.
        MaxLeafSpeed - Float. The maximum distance traveled by any leaf over 1 degree of arc. Defaults to 1.2. If none or 0, no constraint will be implemented.
        popup - Boolean. If true, a window will pop up upon encountering an error to provide the user feedback.
    """

    # Set some defaults
    if beam_set is None or plan is None:
        try:
            beam_set = get_current("BeamSet")
            plan = get_current("Plan")
        except Exception as e:
            if popup:
                MessageBox.Show(
                    "Could not load a beamset, please make sure a beamset is open. Exiting script.")
            print(
                'Failed to load beam_set and plan using get_current(), a beamset may not be open. Failed with the following error message:\n' + str(
                    e))
            return None

    # Make sure an actual beamset and plan are supplied:
    try:
        beam_set.DicomPlanLabel
        plan.BeamSets
    except:
        if popup:
            MessageBox.Show(
                "The plan or beamset supplied are not the correct data type (e.g. perhaps a string was supplied with the plan name rather than a Raystation plan object).")
        print(
            'Failed to access beam_set and plan expected attributes, wrong data type may have been supplied.')
        return None

    # Check that the beamset corresponds to the plan.
    if beam_set not in plan.BeamSets:
        if popup:
            MessageBox.Show("The beamset provided is not part of the specified plan.")
        print(
            'The beamset supplied as an argument to set_opt_parameters(...) is not part of the plan provided')
        return None

    # Check to be sure beamset is set for inverse planning.
    if beam_set.PlanGenerationTechnique != 'Imrt':
        if popup:
            MessageBox.Show(
                "This beamset is not set for inverse planning. Please change the plan treatment technique to SMLC or VMAT and try again.")
        print("Plan generation technique must be 'Imrt' to set optimization parameters.")
        return None

    # Find the correct beam_set settings to modify. This should work with dual-optimized beam sets as well.
    settings = [val.OptimizationParameters for val in plan.PlanOptimizations for val2 in
                val.OptimizationParameters.TreatmentSetupSettings if
                val2.ForTreatmentSetup.DicomPlanLabel == beam_set.DicomPlanLabel]
    assert len(settings) == 1
    settings = settings[0]

    # If compute intermediate dose is not specified, set automatically depending on delivery technique.
    if ComputeIntermediateDose is None:
        if any([treatment_setup_settings.ForTreatmentSetup.DeliveryTechnique == 'SMLC' for
                treatment_setup_settings in settings.TreatmentSetupSettings]):
            ComputeIntermediateDose = True
        else:
            ComputeIntermediateDose = False

    # Set the MaxNumberOfSegments to 10x number of beams, if applicable
    if MaxNumberOfSegments is None:
        if beam_set.Beams.Count == 0:
            MaxNumberOfSegments = 50
        else:
            MaxNumberOfSegments = max(beam_set.Beams.Count * 10, 50)

    # Check that parameters are valid
    assert 0 < MaxNumberOfIterations, 'A valid number must be entered for the maximum number of iterations.'
    assert 0 < IterationsInPreparationsPhase <= MaxNumberOfIterations, 'The number of iterations before conversion must be less than or equal to the total number of iterations.'
    assert MaxNumberOfSegments >= beam_set.Beams.Count, 'The number of segments must be greater than or equal to the number of beams.'
    assert MinSegmentMUPerFraction >= 0.1, 'The minimum segment MU must be 0.1 or greater.'
    assert MinSegmentArea > 0, 'The minimum segment area must be greater than 0.'
    assert 1 <= MinNumberOfOpenLeafPairs <= 10, 'The minimum number of open leaf pairs must be between 1 and 10 inclusive.'
    assert 0 <= MinLeafEndSeparation <= 10, 'The minimum leaf end separation must be between 0 and 10 inclusive.'
    assert MaxLeafSpeed is None or 0 < MaxLeafSpeed < 20, 'The VMAT leaf speed must be between 0 and 20 exclusive.'

    # Check that the plan is not approved / locked
    if beam_set.Review is not None:
        if popup:
            MessageBox.Show(
                "The plan is locked and cannot be changed. Please unlock the plan and try again.")
        print('Plan is approved or locked. Parameters cannot be set.')
        return None

    # Set parameters. All parameters are located in "plan.PlanOptimizations[i].OptimizationParameters" where i = [number of the beam_set]
    try:
        settings.Algorithm.MaxNumberOfIterations = MaxNumberOfIterations
        settings.DoseCalculation.ComputeFinalDose = ComputeFinalDose
        settings.DoseCalculation.ComputeIntermediateDose = ComputeIntermediateDose
        settings.DoseCalculation.IterationsInPreparationsPhase = IterationsInPreparationsPhase
        # Loops through Beam Set Settings if beam sets are being co-optimized
        for treatment_setup_settings in settings.TreatmentSetupSettings:
            if treatment_setup_settings.ForTreatmentSetup.DeliveryTechnique == 'SMLC':
                s = treatment_setup_settings.SegmentConversion
                s.MaxNumberOfSegments = MaxNumberOfSegments
                s.MinSegmentArea = MinSegmentArea
                s.MinSegmentMUPerFraction = MinSegmentMUPerFraction
                s.MinNumberOfOpenLeafPairs = MinNumberOfOpenLeafPairs
                s.MinLeafEndSeparation = MinLeafEndSeparation
            if treatment_setup_settings.ForTreatmentSetup.DeliveryTechnique == 'DynamicArc':
                s = treatment_setup_settings.SegmentConversion.ArcConversionProperties
                s.UseMaxLeafTravelDistancePerDegree = True
                s.MaxLeafTravelDistancePerDegree = 0.3
    except Exception as e:
        if popup:
            MessageBox.Show(
                'Parameters could not be set. Exiting script. Failed with the following error message:\n' + str(
                    e))
        print(
            'Parameters could not be set. Exiting script. Failed with the following error message:\n' + str(
                e))
        return None


# def get_field_border_at_SAD(beam):
#    """Not functioning yet."""
#    orientation = beam.PatientPosition
#    gantry_rot = beam.GantryAngle
#    couch_rot = beam.CouchAngle
#    coll_rot = beam.InitialCollimatorAngle
#    segment = beam.Segments[0]
#    top = segment.JawPositions[3]
#    bottom = segment.JawPositions[2]
#    leftbank = segment.LeafPositions[0]
#    rightbank = segment.LeafPositions[1]
#
#    numleaves = leftbank.Count
#    if numleaves == 80:
#        start = -19.75
#        step = 0.5
#    else:
#        start = -19.5
#        step = 1
#
#    left_result,right_result = [],[]
#    for i in range(numleaves):
#        pos = start+step*i
#        if bottom < pos < top:
#            leftleaf = [leftbank[i],pos,0]
#            rightleaf = [rightbank[i],pos,0]
#            left_result.append(get_leaf_projection_at_iso(leftleaf,coll_rot,gantry_rot,couch_rot))
#            right_result.append(get_leaf_projection_at_iso(rightleaf,coll_rot,gantry_rot,couch_rot))
#            left_result = [cartesian_to_dicom(each,orientation) for each in left_result]
#            right_result = [cartesian_to_dicom(each,orientation) for each in right_result]
#    return left_result+right_result

# def get_leaf_projection_at_iso(leaf_pos,coll_rot,gantry_rot,couch_rot):
#    """NOT WORKING CORRECTLY YET.
#       Find the projection of the leaf position, supplied as a two-entry list (x,y), in DICOM coordinates for the supplied machine parameters for a HFS patient.
#       x and y are defined in Cartesian coordinates at SAD assuming no gantry/collimator/couch rotations. For IEC61217 this is equivalent to the x and y values reported by
#       the machine for each leaf."""
#    leaf_pos = rot_vect(leaf_pos,'z',coll_rot)
#    leaf_pos = rot_vect(leaf_pos,'y',gantry_rot)
#    leaf_pos = rot_vect(leaf_pos,'z',-1*couch_rot)
#    return leaf_pos


def external_contoured(case, exam):
    """Determine if the external ROI has been contoured on the given examination."""
    structure_set = case.PatientModel.StructureSets[exam.Name].RoiGeometries
    external = [each for each in structure_set if each.OfRoi.Type == 'External']
    if len(external) == 0:
        return False
    return external[0].HasContours()


def roi_contoured(case, exam, name):
    """Determine if the roi withe the specified name (supplied as a string) has been contoured on either the supplied case and exam or the current case and exam if none are specified."""
    if name == None:
        return False
    structure_set = case.PatientModel.StructureSets[exam.Name].RoiGeometries
    roi = [each for each in structure_set if each.OfRoi.Name == name]
    if len(roi) == 0:
        return False
    return roi[0].HasContours()


def rename_exams(exams, popup=False, rename=True):
    """Find the series description for each examination. If rename is True, rename the CT in Raystation. If popup is True, show an error dialogue if the renaming can't be completed.
       Return a list of Dicom Series Descriptions and Raystation CT Names [(Desc,Name),...]
       Examinations should be supplied as a list, but if a single standalone examination is supplied the function will handle that as well."""

    if type(exams) != list:
        exams = [exams]
    errors = []
    result = []
    print(exams)
    for each in exams:
        try:
            dicomTag = each.GetStoredDicomTagValueForVerification(Group=0x0008, Element=0x103e)
            if rename:
                each.Name = dicomTag['Series Description']
                if "20" not in each.Name:
                   scandate=each.GetStoredDicomTagValueForVerification(Group=0x0008, Element=0x0021)['Series Date']
                   each.Name=dicomTag['Series Description']+" "+scandate
            result.append((dicomTag['Series Description'], each.Name))
        except Exception as e:
            print(e)
            errors.append("Could not find CT Series Description for CT '%s'." % (each.Name))
    if len(errors) > 0 and popup:
        MessageBox.Show('\n'.join(errors))
    return result


######################
# Plan Setup Scripts #
######################

def import_couch_model(COCUH):
    """Imports selected couch model and moves it to the correct location.  Model is pruned to fit exam. Created by WL, March 2018 """
    patient = get_current("Patient")
    case = get_current("Case")
    examination = get_current("Examination")
    ui = get_current("ui")
    patient_db = get_current('PatientDB')

    couch_templates = {'iBEAM evo': "iBEAM evo", 'Qfix kVue': "Qfix kVue", 'Varian IGRT': "Varian IGRT Couch"}
    couch_template_rois = {'iBEAM evo': ["iBEAM evo Couch Core", "iBEAM evo Couch Shell"],
                           'Qfix kVue': ["Qfix kVue Couch"],
                           'Varian IGRT': ["Varian IGRT Couch Exterior", "Varian IGRT Couch Interior"]}

    approved_plan = False
    approved_roi_names = []
    approved_plans = []

    # Create a list of approved ROIs on the current exam
    try:
        if case.PatientModel.StructureSets[examination.Name].ApprovedStructureSets.Count != 0:
            for each in case.PatientModel.StructureSets[examination.Name].ApprovedStructureSets:
                for approved_roi in each.ApprovedRoiStructures:
                    approved_roi_names.append(approved_roi.OfRoi.Name)
            print('Approved ROIS on this exam are: ' + approved_roi_names)
    except:
        print('Could not store approved ROI names')

    # If the couch model pieces on current exam are approved, set approved_plan to True
    if 'iBEAM evo Couch Core' in approved_roi_names:
        approved_plan = True
    if 'iBEAM evo Couch Shell' in approved_roi_names:
        approved_plan = True
    if 'Qfix kVue Couch' in approved_roi_names:
        approved_plan = True
    if 'Varian IGRT Couch Exterior' in approved_roi_names:
        approved_plan = True
    if 'Varian IGRT Couch Interior' in approved_roi_names:
        approved_plan = True

    # If an approved plan exists on current exam, set approved_plan to True
    for each in case.TreatmentPlans:
        if each.Review == 'Approved':
            approved_plans.append(each.Name)
    if len(approved_plans) > 0:
        approved_plan = True

        # Create a list of Beam Sets with Dose calculated on the current exam
    beamsets_with_dose = []
    try:
        for each in case.TreatmentPlans:
            if each.TreatmentCourse.TotalDose.OnDensity.FromExamination.Name == examination.Name:
                for each_set in each.BeamSets:
                    if each_set.FractionDose.DoseValues != None:
                        beamsets_with_dose.append(each_set.DicomPlanLabel)
    except:
        print('Could Not Determine Number of Beamsets with Dose')

    # If an approved plan does not exist on the current exam and dose has not been calculated
    if not approved_plan and len(beamsets_with_dose) == 0:

        # ---Initialize Couch Model ROI's by deleting them if present---#
        try:
            # Change View to Patient Modeling, Structure Definition
            ui.TitleBar.MenuItem['Patient Modeling'].Button_PatientModeling.Click()
            ui.TabControl_Modules.TabItem['Structure Definition'].Select()
        except:
            print('Unable to change UI')

        # If Couch Shell Geometry exists, delete it

        try:
            if patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries[
                'iBEAM evo Couch Shell'].PrimaryShape != None:
                case.PatientModel.StructureSets[examination.Name].RoiGeometries['iBEAM evo Couch Shell'].DeleteGeometry()
        except:
            print("Couch Shell Geometry Not Present or is Locked")

        # If Couch Core Geometry exists, delete it
        try:
            if patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries[
                'iBEAM evo Couch Core'].PrimaryShape != None:
                case.PatientModel.StructureSets[examination.Name].RoiGeometries['iBEAM evo Couch Core'].DeleteGeometry()
        except:
            print("Couch Core Geometry Not Present or is Locked")

        # If Qfix kVue Couch Geometry exists, delete it
        try:
            if patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries[
                'Qfix kVue Couch'].PrimaryShape != None:
                case.PatientModel.StructureSets[examination.Name].RoiGeometries['Qfix kVue Couch'].DeleteGeometry()
        except:
            print("Qfix kVue Couch Geometry Not Present or is Locked")
            
        # If Varian IGRT Couch Exterior Geometry exists, delete it
        try:
            if patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries[
                    'Varian IGRT Couch Exterior'].PrimaryShape != None:
                case.PatientModel.StructureSets[examination.Name].RoiGeometries['Varian IGRT Couch Exterior'].DeleteGeometry()
        except:
            print("Varian IGRT Couch Exterior Geometry Not Present or is Locked")

         # If Varian IGRT Couch Exterior Geometry exists, delete it
        try:
            if patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries[
                     'Varian IGRT Couch Interior'].PrimaryShape != None:
                case.PatientModel.StructureSets[examination.Name].RoiGeometries['Varian IGRT Couch Interior'].DeleteGeometry()
        except:
            print("Varian IGRT Couch Interior Geometry Not Present or is Locked")
            
        # initialize table height
        th = None

        # ---End Initialization---#

        # ---Begin Import---#

        # Get Table Height from Dicom Header
        try:
            th = case.Examinations[examination.Name].GetStoredDicomTagValueForVerification(Group=0x0018, Element=0x1130)
            th = float(th.get('Table Height'))
            print("Table Height is:", th)
        except:
            print("Could Not Get Table Height")

        # Check to see if patient is prone, import appropriate structure template, find center, and assign appropriate x,y, and z offsets
        x = 0  # right
        z = 0  # sup
        print(COCUH)

        if th is not None:
            # iBEAM evo was selected
            if COCUH == 'iBEAM evo':
                if patient.Cases[case.CaseName].Examinations[examination.Name].PatientPosition == "HFS" or \
                        patient.Cases[case.CaseName].Examinations[examination.Name].PatientPosition == "FFS":
                    template = patient_db.LoadTemplatePatientModel(templateName=couch_templates[COCUH], lockMode='Read')
                    case.PatientModel.CreateStructuresFromTemplate(SourceTemplate=template,
                                                                   SourceExaminationName="CT 1",
                                                                   SourceRoiNames=couch_template_rois[COCUH],
                                                                   SourcePoiNames=[], AssociateStructuresByName=True,
                                                                   TargetExamination=examination,
                                                                   InitializationOption="AlignImageCenters")
                 

                    center_shell = \
                    patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries['iBEAM evo Couch Shell'].GetCenterOfRoi()
                      
                
                    print(th)
                    print(center_shell.y)

                    y = (((th / 10) - center_shell.y)-4.3) 
               

                else:  # patient is prone
                    template = patient_db.LoadTemplatePatientModel(templateName="iBEAM evo Prone",
                                                                   lockMode='Read')
                    case.PatientModel.CreateStructuresFromTemplate(SourceTemplate=template,
                                                                   SourceExaminationName="CT 1",
                                                                   SourceRoiNames=couch_template_rois[COCUH],
                                                                   SourcePoiNames=[], AssociateStructuresByName=True,
                                                                   TargetExamination=examination,
                                                                   InitializationOption="AlignImageCenters")
                 
                    center_shell = \
                    patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries['iBEAM evo Couch Shell'].GetCenterOfRoi()

     
                    y = ((((-1*th) / 10) - center_shell.y) + 4.5) 
                    print(y)


            elif COCUH == 'Varian IGRT':
                if patient.Cases[case.CaseName].Examinations[examination.Name].PatientPosition == "HFS" or \
                        patient.Cases[case.CaseName].Examinations[examination.Name].PatientPosition == "FFS":
                    template = patient_db.LoadTemplatePatientModel(templateName=couch_templates[COCUH], lockMode='Read')
                    case.PatientModel.CreateStructuresFromTemplate(SourceTemplate=template,
                                                                   SourceExaminationName="CT 1",
                                                                   SourceRoiNames=couch_template_rois[COCUH],
                                                                   SourcePoiNames=[], AssociateStructuresByName=True,
                                                                   TargetExamination=examination,
                                                                   InitializationOption="AlignImageCenters")

                    center_shell = \
                    patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries['Varian IGRT Couch Interior'].GetCenterOfRoi()
                    
                    print(th)
                    print(center_shell.y)

                    y = (((th / 10) - center_shell.y)-3.5)  


                else:  # patient is prone
                    template = patient_db.LoadTemplatePatientModel(templateName="Varian IGRT Couch Prone",
                                                                   lockMode='Read')
                    case.PatientModel.CreateStructuresFromTemplate(SourceTemplate=template,
                                                                   SourceExaminationName="CT 1",
                                                                   SourceRoiNames=couch_template_rois[COCUH],
                                                                   SourcePoiNames=[], AssociateStructuresByName=True,
                                                                   TargetExamination=examination,
                                                                   InitializationOption="AlignImageCenters")

                    center_shell = \
                    patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries['Varian IGRT Couch Interior'].GetCenterOfRoi()
                    print(th)
                    print(center_shell.y)
                    y = (((( -1*th) / 10) - center_shell.y)+3.9) 
                    print(y)
                        
            #Qfix kVue was selected
            else:
                if patient.Cases[case.CaseName].Examinations[examination.Name].PatientPosition == "HFS" or \
                        patient.Cases[case.CaseName].Examinations[examination.Name].PatientPosition == "FFS":
                    template = patient_db.LoadTemplatePatientModel(templateName="Qfix kVue", lockMode='Read')
                    case.PatientModel.CreateStructuresFromTemplate(SourceTemplate=template,
                                                                   SourceExaminationName="CT 1",
                                                                   SourceRoiNames=couch_template_rois[COCUH], SourcePoiNames=[],
                                                                   AssociateStructuresByName=True,
                                                                   TargetExamination=examination,
                                                                   InitializationOption="AlignImageCenters")

                    center_shell = \
                    patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries['Qfix kVue Couch'].GetCenterOfRoi()

                    y = (((th / 10) - center_shell.y)-5) 
          

                else:  # patient is prone
                    template = patient_db.LoadTemplatePatientModel(templateName="Qfix kVue", lockMode='Read')
                    case.PatientModel.CreateStructuresFromTemplate(SourceTemplate=template,
                                                                   SourceExaminationName="CT 1",
                                                                   SourceRoiNames=couch_template_rois[COCUH], SourcePoiNames=[],
                                                                   AssociateStructuresByName=True,
                                                                   TargetExamination=examination,
                                                                   InitializationOption="AlignImageCenters")

                    center_shell = \
                    patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries[
                        'Qfix kVue Couch'].GetCenterOfRoi()
                    
                    print(center_shell.y)
                    y = ((((th * -1) / 10) - center_shell.y)+5.4) 
                 
                
            print(y)
       
            # Define the transformation in x,y,z space to be performed
            transform_matrix = {'M11': 1, 'M12': 0, 'M13': 0, 'M14': x,
                                'M21': 0, 'M22': 1, 'M23': 0, 'M24': y,
                                'M31': 0, 'M32': 0, 'M33': 1, 'M34': z,
                                'M41': 0, 'M42': 0, 'M43': 0, 'M44': 1}

            # Apply the transformation to AD couch model
            try:
                for roi in couch_template_rois[COCUH]:
                    patient.Cases[case.CaseName].PatientModel.RegionsOfInterest[roi].TransformROI3D(
                        Examination=examination, TransformationMatrix=transform_matrix)

            except:
                print('Could not move the model')

            # ---End Import---#

            # ---Prune the Couch Model so that it does not extend off examination---#

            # Create Bounding Box so CouchModel can be Pruned to Fit Dataset
            ct_image_bounding_box = examination.Series[0].ImageStack.GetBoundingBox()
            ct_sup = max([each.z for each in ct_image_bounding_box])
            ct_inf = min([each.z for each in ct_image_bounding_box])
            ct_left = max([each.x for each in ct_image_bounding_box])
            ct_right = min([each.x for each in ct_image_bounding_box])
            ct_ant = max([each.y for each in ct_image_bounding_box])
            ct_post = min([each.y for each in ct_image_bounding_box])
            # Determine Size of Box
            control_box_size_x = abs(ct_left - ct_right)
            control_box_size_y = abs(ct_ant - ct_post)
            control_box_size_z = abs(ct_sup - ct_inf)
            control_box_size_xyz = {"x": control_box_size_x, "y": control_box_size_y, "z": control_box_size_z}
            if patient.Cases[case.CaseName].Examinations[examination.Name].PatientPosition == "HFS" or \
                    patient.Cases[case.CaseName].Examinations[examination.Name].PatientPosition == "FFS":
                center_control_box = {"x": (control_box_size_x / 2) + examination.Series[0].ImageStack.Corner.x,
                                      "y": (control_box_size_y / 2) + examination.Series[0].ImageStack.Corner.y, "z":
                                          (control_box_size_z / 2) + examination.Series[0].ImageStack.Corner.z}
            else:
                center_control_box = {"x": (control_box_size_x / 2) - examination.Series[0].ImageStack.Corner.x,
                                      "y": (control_box_size_y / 2) - examination.Series[0].ImageStack.Corner.y, "z":
                                          (control_box_size_z / 2) + examination.Series[0].ImageStack.Corner.z}

            # Create Control Box around Dataset
            with CompositeAction('Create Box ROI (control_box)'):
                retval_0 = case.PatientModel.CreateRoi(Name="control_box", Color="Orange", Type="Control",
                                                       TissueName=None, RbeCellTypeName=None, RoiMaterial=None)
                retval_0.CreateBoxGeometry(Size=control_box_size_xyz, Examination=examination,
                                           Center=center_control_box, VoxelSize=None)

                for roi in couch_template_rois[COCUH]:
                    # Create Temporary Copies of the Couch Model ROIs that are the Intersection of Couch Model and the Control Box.
                    retval_0 = case.PatientModel.CreateRoi(Name=roi + '_temp', Color="Green", Type="Organ",
                                                           TissueName=None, RbeCellTypeName=None, RoiMaterial=None)
                    retval_0.CreateAlgebraGeometry(Examination=examination, Algorithm="Auto",
                                                   ExpressionA={'Operation': "Union", 'SourceRoiNames': [roi],
                                                                'MarginSettings': {'Type': "Expand",
                                                                                   'Superior': 0, 'Inferior': 0,
                                                                                   'Anterior': 0, 'Posterior': 0,
                                                                                   'Right': 0, 'Left': 0}},
                                                   ExpressionB={'Operation': "Union", 'SourceRoiNames': ["control_box"],
                                                                'MarginSettings':
                                                                    {'Type': "Expand", 'Superior': 0, 'Inferior': 0,
                                                                     'Anterior': 0, 'Posterior': 0, 'Right': 0,
                                                                     'Left': 0}}, ResultOperation="Intersection",
                                                   ResultMarginSettings=
                                                   {'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0,
                                                    'Posterior': 0, 'Right': 0, 'Left': 0})

                    # Copy the temporary ROIs to the Couch Model ROIs
                    # case.PatientModel.StructureSets[examination.Name].CopyRoiGeometryToAnotherRoi(FromRoi=roi+'_temp', ToRoi=roi)
                    case.PatientModel.RegionsOfInterest[roi].CreateAlgebraGeometry(Examination=examination,
                                                                                   Algorithm="Auto",
                                                                                   ExpressionA={
                                                                                       'Operation': "Union",
                                                                                       'SourceRoiNames': [
                                                                                           roi + '_temp'],
                                                                                       'MarginSettings': {
                                                                                           'Type': "Expand",
                                                                                           'Superior': 0,
                                                                                           'Inferior': 0,
                                                                                           'Anterior': 0,
                                                                                           'Posterior': 0,
                                                                                           'Right': 0,
                                                                                           'Left': 0}},
                                                                                   ExpressionB={
                                                                                       'Operation': "Union",
                                                                                       'SourceRoiNames': [],
                                                                                       'MarginSettings': {
                                                                                           'Type': "Expand",
                                                                                           'Superior': 0,
                                                                                           'Inferior': 0,
                                                                                           'Anterior': 0,
                                                                                           'Posterior': 0,
                                                                                           'Right': 0,
                                                                                           'Left': 0}},
                                                                                   ResultOperation="None",
                                                                                   ResultMarginSettings=
                                                                                   {'Type': "Expand",
                                                                                    'Superior': 0,
                                                                                    'Inferior': 0,
                                                                                    'Anterior': 0,
                                                                                    'Posterior': 0, 'Right': 0,
                                                                                    'Left': 0})

                    # Delete temp Structures
                    case.PatientModel.RegionsOfInterest[roi + '_temp'].DeleteRoi()

                # Delete the control box
                case.PatientModel.RegionsOfInterest['control_box'].DeleteRoi()

            # ---End Prune Couch---#

            # Attempt to recompute dose for all beamsets in case

        else:
            MessageBox.Show(
                'Unable to set couch model.  Dicom Header does not include table height.  Please import manually.')
    else:
        MessageBox.Show('Unable to add couch model.  ROIs are locked or Dose is calculated on the current exam.')

    return y

def ROI_setup():
    """Assigns ROI type by name.  Turns visibility off when ROIs have no contours on the current examination. Created by WL, June 2018"""
    patient = get_current("Patient")
    case = get_current("Case")
    examination = get_current("Examination")

    name_check_gtv = ['GTV', 'gtv', 'Gtv', 'ITV', 'itv', 'Itv', 'iTV', 'GTV_ITV']
    name_check_ctv = ['CTV', 'ctv', 'Ctv']
    name_check_ptv = ['PTV', 'ptv', 'Ptv']

    approved_roi_names = [
        'Liver-GTV']  # Liver-GTV is assumed to be approved even if it is not.  Forces it to be an OAR

    # Create a list of all ROIs on all Exams that have been approved (regardless of whether it has contours)
    try:
        for each in case.PatientModel.StructureSets:
            if each.ApprovedStructureSets.Count != 0:
                for approved_ss in each.ApprovedStructureSets:
                    for approved_roi in approved_ss.ApprovedRoiStructures:
                        approved_roi_names.append(approved_roi.OfRoi.Name)
            print(approved_roi_names)
    except:
        print('Could not store approved ROI names')

    for each in case.PatientModel.RegionsOfInterest:

        if each.Name == 'Liver-GTV':
            try:
                case.PatientModel.RegionsOfInterest[each.Name].OrganData.OrganType = "OrganAtRisk"
                case.PatientModel.RegionsOfInterest[each.Name].Type = "Organ"
            except:
                print("Could not Set ROI Type for Liver-GTV")

        for each_namecheck in name_check_gtv:
            if each_namecheck in each.Name:
                print(each.Name + " Is a GTV or ITV")
                if each.Name not in approved_roi_names:  # check to see if ROI is not approved
                    try:
                        case.PatientModel.RegionsOfInterest[each.Name].Type = "Gtv"
                        case.PatientModel.RegionsOfInterest[
                            each.Name].OrganData.OrganType = "Target"
                    except:
                        print("Could not Set ROI Type for " + each.Name)
        for each_namecheck in name_check_ctv:
            if each_namecheck in each.Name:
                print(each.Name + " Is a CTV")
                if each.Name not in approved_roi_names:
                    try:
                        case.PatientModel.RegionsOfInterest[each.Name].Type = "Ctv"
                        case.PatientModel.RegionsOfInterest[
                            each.Name].OrganData.OrganType = "Target"
                    except:
                        print("Could not Set ROI Type for " + each.Name)
        for each_namecheck in name_check_ptv:
            if each_namecheck in each.Name:
                print(each.Name + " Is a PTV")
                if each.Name not in approved_roi_names:
                    try:
                        case.PatientModel.RegionsOfInterest[each.Name].Type = "Ptv"
                        case.PatientModel.RegionsOfInterest[
                            each.Name].OrganData.OrganType = "Target"
                    except:
                        print("Could not Set ROI Type for " + each.Name)

        if case.PatientModel.RegionsOfInterest[
            each.Name].Type == 'Support':  # If an ROI is a Support, set its type to Other
            print(each.Name + " Is a Support, Organ Type set to Other")
            if each.Name not in approved_roi_names:
                case.PatientModel.RegionsOfInterest[each.Name].OrganData.OrganType = "Other"

        elif case.PatientModel.RegionsOfInterest[
            each.Name].OrganData.OrganType == "Target":  # If an ROI is already labelled as a Target, leave it alone
            print(each.Name + " Is already a Target, nothing changed")

        elif case.PatientModel.RegionsOfInterest[
            each.Name].RoiMaterial is not None:  # If an ROI is overriden, set its type to Other
            print(each.Name + " Has a denisty override, Organ Type set to Other")
            if each.Name not in approved_roi_names:
                try:
                    case.PatientModel.RegionsOfInterest[each.Name].OrganData.OrganType = "Other"
                except:
                    print("Could not Set ROI Type for " + each.Name)

        else:  # Set remaining Organ Types as OAR.  Note: External ROI type will remain External
            if case.PatientModel.RegionsOfInterest[each.Name].Type != 'External':
                print(each.Name + " Is an OAR")
                if each.Name not in approved_roi_names:
                    try:
                        case.PatientModel.RegionsOfInterest[
                            each.Name].OrganData.OrganType = "OrganAtRisk"
                        case.PatientModel.RegionsOfInterest[each.Name].Type = "Organ"
                    except:
                        print("Could not Set ROI Type for " + each.Name)

    for each in case.PatientModel.RegionsOfInterest:  # Turn visibility on if ROI has contours, or off if it does not.  This helps prevent a known bug in RS with ROI with empty geometry.  It's also just helpful to see which are empty.
        if not case.PatientModel.StructureSets[examination.Name].RoiGeometries[
            each.Name].HasContours():
            print(each.Name + " Is EMPTY, turning visibility off")
            try:
                patient.SetRoiVisibility(RoiName=each.Name, IsVisible=False)
            except:
                print('Could not change visibility, ROI may have been a donut with a line')
        else:
            print(each.Name + " is NOT empty, turning visibility on.")
            try:
                patient.SetRoiVisibility(RoiName=each.Name, IsVisible=True)
            except:
                print('Could not change visibility, ROI may have been a donut with a line')


def create_external():
    """Creates a new External Contour or Geometry as long as the External ROI is not locked and the current examination does not have dose calculated on it. Removes contours smaller than 20cc. Created by WL, June 2018"""
    patient = get_current("Patient")
    case = get_current("Case")
    examination = get_current("Examination")

    # Create a list of all ROIs on all Exams that have been approved (regardless of whether it has contours)
    approved_roi_names = []

    # Create a list of approved ROIs on the current exam
    try:
        if case.PatientModel.StructureSets[examination.Name].ApprovedStructureSets.Count != 0:
            for each in case.PatientModel.StructureSets[examination.Name].ApprovedStructureSets:
                for approved_roi in each.ApprovedRoiStructures:
                    approved_roi_names.append(approved_roi.OfRoi.Name)
            print('Approved ROIS on this exam are: ' + approved_roi_names)
    except:
        print('Could not store approved ROI names')

    # Create a list of approved ROIs on the all exams
    approved_roi_names_all_exams = []
    try:
        for each in case.PatientModel.StructureSets:
            if each.ApprovedStructureSets.Count != 0:
                for approved_ss in each.ApprovedStructureSets:
                    for approved_roi in approved_ss.ApprovedRoiStructures:
                        approved_roi_names_all_exams.append(approved_roi.OfRoi.Name)
            print(approved_roi_names_all_exams)
    except:
        print('Could not store approved ROI names')

    # Create a list of Beam Sets with Dose calculated on the current exam
    beamsets_with_dose = []
    try:
        for each in case.TreatmentPlans:
            if each.TreatmentCourse.TotalDose.OnDensity.FromExamination.Name == examination.Name:
                for each_set in each.BeamSets:
                    if each_set.FractionDose.DoseValues is not None:
                        beamsets_with_dose.append(each_set.DicomPlanLabel)
    except:
        print('Could Not Determine Number of Beamsets with Dose')

    # Create New External Geometry if External ROI is not locked and dose is not calculated on the current exam
    if 'External' not in approved_roi_names and len(beamsets_with_dose) == 0:

        try:
            patient.Cases[case.CaseName].PatientModel.StructureSets[examination.Name].RoiGeometries[
                'External'].DeleteGeometry()
            print('External Geometry Was Deleted')
        except:
            print('Unable to delete External Geometry')

    for each in case.PatientModel.RegionsOfInterest:
        try:
            if each.Name == 'External' and not \
            case.PatientModel.StructureSets[examination.Name].RoiGeometries[
                each.Name].HasContours():  # If the name External exists but there are no contours:
                print(
                    'ROI named External already exists but has no contours on current exam, create new External Geometry')
                external = patient.Cases[case.CaseName].PatientModel.RegionsOfInterest['External']
                external.CreateExternalGeometry(Examination=examination,
                                                ThresholdLevel=-250)  # Create New Geometry
                case.PatientModel.StructureSets[examination.Name].SimplifyContours(
                    RoiNames=["External"], RemoveHoles3D=False, RemoveSmallContours=True,
                    AreaThreshold=20,
                    ReduceMaxNumberOfPointsInContours=False, MaxNumberOfPoints=None,
                    CreateCopyOfRoi=False)  # Remove Contours less than 20cc
                if 'External' not in approved_roi_names_all_exams:
                    case.PatientModel.RegionsOfInterest['External'].Color = "Olive"
                return

            elif each.Type == 'External' and not \
            case.PatientModel.StructureSets[examination.Name].RoiGeometries[
                each.Name].HasContours():  # If the type External exists but there are no contours:
                external = patient.Cases[case.CaseName].PatientModel.RegionsOfInterest[each.Name]
                external.CreateExternalGeometry(Examination=examination,
                                                ThresholdLevel=-250)  # Create New Geometry
                case.PatientModel.StructureSets[examination.Name].SimplifyContours(
                    RoiNames=["External"], RemoveHoles3D=False, RemoveSmallContours=True,
                    AreaThreshold=20,
                    ReduceMaxNumberOfPointsInContours=False, MaxNumberOfPoints=None,
                    CreateCopyOfRoi=False)  # Remove Contours less than 20cc
                print(
                    'ROI with type External already exists but has no contours on current exam, create new External Geometry')
                if 'External' not in approved_roi_names_all_exams:
                    case.PatientModel.RegionsOfInterest['External'].Color = "Olive"
                return

        except:
            print('Could not create new couch geometries')

        # Create New External ROI if dose is not calculated on the current exam
    if len(beamsets_with_dose) == 0:
        try:
            with CompositeAction('Create external (External)'):
                external = case.PatientModel.CreateRoi(Name="External", Color="Olive",
                                                       Type="External", TissueName="",
                                                       RbeCellTypeName=None, RoiMaterial=None)
                external.CreateExternalGeometry(Examination=examination, ThresholdLevel=-250)
                case.PatientModel.StructureSets[examination.Name].SimplifyContours(
                    RoiNames=["External"], RemoveHoles3D=False, RemoveSmallContours=True,
                    AreaThreshold=20,
                    ReduceMaxNumberOfPointsInContours=False, MaxNumberOfPoints=None,
                    CreateCopyOfRoi=False)  # Remove Contours less than 20cc
        except:
            print('Unable to Generate a new External Contour')
            MessageBox.Show('Unable to Generate a new External Contour')

    else:
        MessageBox.Show('Unable to create new External. Dose is calculated on the current exam.')
