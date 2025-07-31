# -------------------------------------------------------------------------------
# Name:        Dose Slice Report (v1.03) - Final Clean Version
#
# Purpose:     Generates a slice-report for the composite dose distribution.
#              Automatically prompts the user to save the PDF (RayStation handles it).
# Author:      Original: LSW (UWMC), Fixes by ChatGPT
# -------------------------------------------------------------------------------

from connect import get_current
import xUWScriptingUtilities as su
from System import Windows


def process_dose(plan, dosearray, margin=1):
    dose_grid = plan.BeamSets[0].FractionDose.InDoseGrid
    corner, numvx, voxsz = dose_grid.Corner, dose_grid.NrVoxels, dose_grid.VoxelSize
    x0, y0, z0 = corner.x, corner.y, corner.z
    xn, yn, zn = numvx.x, numvx.y, numvx.z
    xr, yr, zr = voxsz.x, voxsz.y, voxsz.z

    doseli = list(dosearray.flatten())

    max_dose = max(doseli)
    max_dose_i = doseli.index(max_dose)
    md_x = x0 + ((max_dose_i % xn) + 0.5) * xr
    md_y = y0 + (int(max_dose_i % (xn * yn) / xn) + 0.5) * yr
    md_z = z0 + (int(max_dose_i / (xn * yn)) + 0.5) * zr

    max_slice_dose = [max(doseli[i * xn * yn:i * xn * yn + xn * yn - 1]) for i in range(zn)]
    max_slice_dose = [each / max_dose for each in max_slice_dose]
    max_slice_dose = [float(each) > 0.15 for each in max_slice_dose]

    start, stop = [], []
    trueflag = False
    for i, v in enumerate(max_slice_dose):
        if v and not trueflag:
            start.append(i)
            trueflag = True
        if not v and trueflag:
            stop.append(i)
            trueflag = False
    if trueflag:
        stop.append(i)

    start = [z0 + zr * each - 1 for each in start]
    stop = [z0 + zr * each + 1 for each in stop]
    remove = []

    for i in range(1, len(start)):
        distance = start[i] - stop[i - 1]
        if distance < 3:
            remove.append(i)
    remove.reverse()
    for each in remove:
        del start[each]
        del stop[each - 1]

    assert len(start) == len(stop), "Start and stop lists not equal. Likely bug."
    return [(start[i], stop[i]) for i in range(len(start))], max_dose, md_x, md_y, md_z


def run_dose_report(patient, case, plan):
    exam = plan.BeamSets[0].PatientSetup.OfTreatmentSetup.GetPlanningExamination()

    if plan.BeamSets.Count == 1:
        dose = plan.TreatmentCourse.TotalDose.DoseValues.DoseData
        startstop, max_dose, md_x, md_y, md_z = process_dose(plan, dose)
        startstop = [[each[0], each[1], 0, 0] for each in startstop]
        for i, v in enumerate(startstop):
            if v[0] > v[1]:
                startstop[i][0], startstop[i][1] = startstop[i][1], startstop[i][0]

        plan.BeamSets[0].EditShowBeamVisualization(ShowBeams=False, ShowContour=False,
                                                   ShowCenterLine=False,
                                                   ShowBeamsFromAllBeamSets=False,
                                                   ShowIsocenterNames=False)

        Windows.MessageBox.Show("Select a location to save the Dose Slice Report PDF when prompted.")
        su.generate_slice_report(startstopfocus=startstop,
                                 maxdose=[round(max_dose), md_x, md_y, md_z])

        plan.BeamSets[0].EditShowBeamVisualization(ShowBeams=True, ShowContour=False,
                                                   ShowCenterLine=False,
                                                   ShowBeamsFromAllBeamSets=False,
                                                   ShowIsocenterNames=False)

    else:
        plan_names = [each.Name for each in case.TreatmentPlans]
        if 'Delete-CompDose' in plan_names:
            Windows.MessageBox.Show("Please delete the plan named 'Delete-CompDose' before running this script.")
            return False

        dgparams = plan.GetTotalDoseGrid()
        total_dose = plan.TreatmentCourse.TotalDose.DoseValues.DoseData
        example_beamset = plan.BeamSets[0]

        dcm = case.CaseSettings.DoseColorMap
        if dcm.ColorMapReferenceType != "ReferenceValue":
            Windows.MessageBox.Show("Isodose display must use 'Reference Value' mode.")
            return False

        newplan = case.AddNewPlan(PlanName='Delete-CompDose',
                                  PlannedBy='Generated Automatically',
                                  Comment='For composite dose report generation only.',
                                  ExaminationName=exam.Name,
                                  AllowDuplicateNames=False)

        bs = newplan.AddNewBeamSet(Name='CompositeDose',
                                   ExaminationName=exam.Name,
                                   MachineName=example_beamset.MachineReference.MachineName,
                                   Modality=example_beamset.Modality,
                                   TreatmentTechnique=example_beamset.GetTreatmentTechniqueType(),
                                   PatientPosition=example_beamset.PatientPosition,
                                   NumberOfFractions=1,
                                   CreateSetupBeams=False,
                                   UseLocalizationPointAsSetupIsocenter=True,
                                   Comment='For composite dose report only.')

        bs.UpdateDoseGrid(Corner={'x': dgparams.Corner.x, 'y': dgparams.Corner.y, 'z': dgparams.Corner.z},
                          VoxelSize={'x': dgparams.VoxelSize.x, 'y': dgparams.VoxelSize.y, 'z': dgparams.VoxelSize.z},
                          NumberOfVoxels={'x': dgparams.NrVoxels.x, 'y': dgparams.NrVoxels.y,
                                          'z': dgparams.NrVoxels.z})

        bs.FractionDose.SetDoseValues(Dose=total_dose, CalculationInfo='Composite')

        startstop, max_dose, md_x, md_y, md_z = process_dose(plan, total_dose)
        startstop = [[each[0], each[1], 0, 0] for each in startstop]
        for i, v in enumerate(startstop):
            if v[0] > v[1]:
                startstop[i][0], startstop[i][1] = startstop[i][1], startstop[i][0]

        patient.Save()
        newplan.SetCurrent()

        Windows.MessageBox.Show("Select a location to save the Dose Slice Report PDF when prompted.")
        su.generate_slice_report(startstopfocus=startstop,
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
