from connect import get_current
import os

case = get_current("Case")
examination = get_current("Examination")
ui = get_current("ui")
plan = case.TreatmentPlans[0]

# --- Settings ---
output_dir = os.path.expanduser("~/RayStationScreenshots")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

visible_rois = ['Liver', 'GTV', 'CTV', 'PTV', 'SpinalCord', 'Kidneys']

# --- Step 1: Set ROI visibility ---
for roi in case.PatientModel.RegionsOfInterest:
    try:
        show = roi.Name in visible_rois
        roi.ReviewDisplayColor = 'Red' if show else 'Gray'
        for view in ['Transversal', 'Sagittal', 'Coronal']:
            case.SetRoiVisibility(RoiName=roi.Name, View=view, Visible=show)
    except:
        print(f"Failed ROI visibility change: {roi.Name}")

# --- Step 2: Turn on dose display ---
try:
    dose_name = plan.TreatmentCourse.TotalDose.OnStructureSet.DicomPlanLabel
    for view in ['Transversal', 'Sagittal', 'Coronal']:
        case.SetDoseVisibility(DoseName=dose_name, View=view, Visible=True)
except Exception as e:
    print(f"Failed to access dose or set dose visibility: {e}")

# --- Step 3: Screenshot the current screen (user manually selects view) ---
screenshot_path = os.path.join(output_dir, "CurrentView.png")
try:
    ui.SaveScreenShot(FilePath=screenshot_path)
    print(f"Screenshot saved: {screenshot_path}")
except:
    print("Could not save screenshot from current view.")

# --- Step 4: 3D View Screenshots ---
try:
    ui.Open3DView()
    ui.ThreeDView.SetDoseVisibility(True)
    ui.ThreeDView.SetRoiVisibility(True)

    angles = [(0, 0), (90, 0), (0, 90), (45, 45)]
    for idx, (yaw, pitch) in enumerate(angles):
        ui.ThreeDView.SetCameraOrientation(Yaw=yaw, Pitch=pitch)
        screenshot_path = os.path.join(output_dir, f"3D_View_{idx}.png")
        ui.ThreeDView.SaveScreenshot(FilePath=screenshot_path)

    print(f"\n3D screenshots saved to: {output_dir}")
except:
    print("3D View or screenshot failed.")
