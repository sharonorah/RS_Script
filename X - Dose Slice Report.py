# -------------------------------------------------------------------------------
# Name:        Dose Slice Report (v1.03)
#
# Written for RS Version: 6.1.1.2
#
# Validated for RS Version: 8.1.1.8 - Imported modules rewritten to function in RS8, not completely backwards compatible.
#
# Purpose:     Generates a slice-report for the composite dose distribution of all beam-sets in a plan. Automatically determines start/stop points.
#              Also generates a page on the pdf printout featuring the maximum dose.
#
# Note:        Generates a dummy-plan if multiple beamsets are present and assigns the plan dose to a dummy-beam set in the dummy plan so the standard beamset
#              dose reporting method can be used. Start/stop points automatically determined by dose. Dummy plan must be manually deleted.
#              The script will try to write directly to //viptier1/radonc/pcc/RAYSEARCH but if it fails it will automatically start the remote plan report routine.
#
# Author:      LSW (UWMC)
#
# Created:     10 December 2018 (v1.00)
#
# Updated:     12 February 2019 (v1.01) - Fixed a bug where the wrong examination may be used for the composite dose report.
#                                         Changed the composite beam set name from 'Not For Export' to 'Composite Dose' to avoid confusion.
#                                         Fixed a bug in the error message when a plan was not already open / when the wrong dose reference style was used.
#                                         Added a check for a plan already named 'Comp-Delete' causing the script to exit cleanly rather than crashing.
#              01 July 2019 (v1.02)     - Rewrote for Raystation 8.1.1.2. Output file only has patient name to avoid 'Delete' in plan document name if pdfs are brought in in wrong order.
#              17 October 2019 (v1.03)  - Turns beam contours off so they are not included in the axial printout.
#              11/22/23                 - Updated to run in RS2023B. Still needs to be clinically validated. -AJE
# -------------------------------------------------------------------------------

from connect import get_current
import xUWScriptingUtilities as su
from System import Windows


def process_dose(plan, dosearray, margin=1):
    """Composite dose report specific function. It converts a dosearray into a python list, then finds the location of the maximum dose and it's value,
    and determines all slices with dose > 15% of the maximum calculated dose."""

    # Retrieve Dose Grid Paramaters
    dose_grid = plan.BeamSets[0].FractionDose.InDoseGrid
    corner, numvx, voxsz = dose_grid.Corner, dose_grid.NrVoxels, dose_grid.VoxelSize
    x0, y0, z0 = corner.x, corner.y, corner.z
    xn, yn, zn = numvx.x, numvx.y, numvx.z
    xr, yr, zr = voxsz.x, voxsz.y, voxsz.z

    # Convert Dose Array to Python Datatype
    doseli = list(dosearray.flatten())

    # Find Magnitude and Location of Max Dose
    max_dose = max(doseli)
    max_dose_i = doseli.index(max_dose)
    md_x, md_y, md_z = max_dose_i % (xn), int(max_dose_i % (xn * yn) / xn), int(
        max_dose_i / (xn * yn))
    md_x, md_y, md_z = x0 + (md_x + 0.5) * xr, y0 + (md_y + 0.5) * yr, z0 + (
            md_z + 0.5) * zr  # Convert indices to coordinates. Add half a voxel width to generate the center of the voxel and not the corner.

    # Find maximum dose per slice and determine whether it is greater than 15% of the global maximum dose.
    max_slice_dose = [max(doseli[i * xn * yn:i * xn * yn + xn * yn - 1]) for i in range(zn)]
    max_slice_dose = [each / max_dose for each in max_slice_dose]
    max_slice_dose = [float(each) > 0.15 for each in max_slice_dose]

    # Find start/stop locations for report printout based on slices above threshold.
    start, stop = [], []
    trueflag = False
    for i, v in enumerate(max_slice_dose):
        if v is True and trueflag is False:
            start.append(i)
            trueflag = True
        if v is False and trueflag is True:
            stop.append(i)
            trueflag = False
    if trueflag is True:  # Make sure there is a final stop point.
        stop.append(i)
    start = [z0 + zr * each - 1 for each in
             start]  # Convert indices into coordinates, add a margin.
    stop = [z0 + zr * each + 1 for each in stop]
    remove = []
    ### Remove small gaps (<3cm) to provide continuity for nearby targets.
    for i, v in enumerate(start):
        if i == 0:
            continue
        distance = v - stop[i - 1]
        if distance < 3:
            remove.append(i)
    remove.reverse()
    for each in remove:
        del start[i]
        del stop[i - 1]
    assert len(start) == len(
        stop), "Start and stop lists are not the same length, probable bug in automatic start stop determination."
    startstop = [(start[i], stop[i]) for i in range(len(start))]
    return startstop, max_dose, md_x, md_y, md_z


def run_dose_report(patient, case, plan):
    exam = plan.BeamSets[0].PatientSetup.OfTreatmentSetup.GetPlanningExamination()

    if plan.BeamSets.Count == 1:
        dose = plan.TreatmentCourse.TotalDose.DoseValues.DoseData
        startstop, max_dose, md_x, md_y, md_z = process_dose(plan, dose)
        startstop = [[each[0], each[1], 0, 0] for each in startstop]
        for i, v in enumerate(startstop):
            if v[0] > v[1]:  # Start should be less than stop.
                startstop[i][0], startstop[i][1] = startstop[i][1], startstop[i][0]
        plan.BeamSets[0].EditShowBeamVisualization(ShowBeams=False, ShowContour=False,
                                                   ShowCenterLine=False,
                                                   ShowBeamsFromAllBeamSets=False,
                                                   ShowIsocenterNames=False)  # Turn off and result in no beams in plan document?
        su.generate_slice_report(startstopfocus=startstop,
                                 maxdose=[round(max_dose), md_x, md_y, md_z])
        plan.BeamSets[0].EditShowBeamVisualization(ShowBeams=True, ShowContour=False,
                                                   ShowCenterLine=False,
                                                   ShowBeamsFromAllBeamSets=False,
                                                   ShowIsocenterNames=False)  # Turn off and result in no beams in plan document?
    else:
        plan_names = [each.Name for each in case.TreatmentPlans]
        if 'Delete-CompDose' in plan_names:
            Windows.MessageBox.Show(
                "Please delete the plan named 'Delete-CompDose' before running this script.")
            return False
        dgparams = plan.GetTotalDoseGrid()
        total_dose = plan.TreatmentCourse.TotalDose.DoseValues.DoseData
        example_beamset = plan.BeamSets[0]

        dcm = case.CaseSettings.DoseColorMap
        print(dcm.ColorMapReferenceType)
        if dcm.ColorMapReferenceType != "ReferenceValue":
            Windows.MessageBox.Show(
                "The isodose display 100%s definition must be based on 'Reference Value' for composite dose reports but is currently '%s'.\nPlease set to 'Reference Value' and check that entered value is appropriate for the composite dose distribution. Exiting script." % (
                    '%', dcm.ColorMapReferenceType))
            return False
        newplan = case.AddNewPlan(
            PlanName='Delete-CompDose',
            PlannedBy='Generated Automatically',
            Comment='For composite dose report generation only.',
            ExaminationName=exam.Name,
            AllowDuplicateNames=False)

        bs = newplan.AddNewBeamSet(
            Name='CompositeDose',
            ExaminationName=exam.Name,
            MachineName=example_beamset.MachineReference.MachineName,
            Modality=example_beamset.Modality,
            TreatmentTechnique=example_beamset.GetTreatmentTechniqueType(),
            PatientPosition=example_beamset.PatientPosition,
            NumberOfFractions=1,
            CreateSetupBeams=False,
            UseLocalizationPointAsSetupIsocenter=True,
            Comment='For composite dose report only.')

        bs.UpdateDoseGrid(
            Corner={
                'x': dgparams.Corner.x,
                'y': dgparams.Corner.y,
                'z': dgparams.Corner.z},
            VoxelSize={
                'x': dgparams.VoxelSize.x,
                'y': dgparams.VoxelSize.y,
                'z': dgparams.VoxelSize.z},
            NumberOfVoxels={
                'x': dgparams.NrVoxels.x,
                'y': dgparams.NrVoxels.y,
                'z': dgparams.NrVoxels.z})

        bs.FractionDose.SetDoseValues(
            Dose=total_dose,
            CalculationInfo='Composite')

        startstop, max_dose, md_x, md_y, md_z = process_dose(plan, total_dose)
        startstop = [[each[0], each[1], 0, 0] for each in startstop]
        for i, v in enumerate(startstop):
            if v[0] > v[1]:  # Start should be less than stop
                startstop[i][0], startstop[i][1] = startstop[i][1], startstop[i][0]

        patient.Save()
        newplan.SetCurrent()
        su.generate_slice_report(
            startstopfocus=startstop,
            maxdose=[round(max_dose), md_x, md_y, md_z])
        Windows.MessageBox.Show("Script complete. Please delete the automatically generated plan.")

    return True


if __name__ == '__main__':
    skip = False
    try:
        patient = get_current('Patient')
        case = get_current('Case')
        plan = get_current('Plan')
    except:
        Windows.MessageBox.Show("A plan must be open to run this script.")
        skip = True
    if not skip:
        run_dose_report(patient, case, plan)
