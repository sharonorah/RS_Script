from connect import get_current
import os
import clr

# --- Windows Folder Selection ---
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import FolderBrowserDialog, DialogResult

folder_dialog = FolderBrowserDialog()
folder_dialog.Description = "Select folder to save screenshots"
dialog_result = folder_dialog.ShowDialog()

if dialog_result != DialogResult.OK:
    raise Exception("Screenshot export cancelled by user.")

output_dir = folder_dialog.SelectedPath

# Confirm folder exists
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# --- RayStation Setup ---
case = get_current("Case")
examination = get_current("Examination")
ui = get_current("ui")
plan = case.TreatmentPlans[0]

visible_rois = ['Liver', 'GTV', 'CTV', 'PTV', 'SpinalCord', 'Kidneys']

# --- Set ROI Visibility ---
for roi in case.PatientModel.RegionsOfInterest:
    try:
        show = roi.Name in visible_rois
        roi.ReviewDisplayColor = 'Red' if show else 'Gray'
        for view in ['Transversal', 'Sagittal', 'Coronal']:
            case.SetRoiVisibility(RoiName=roi.Name, View=view, Visible=show)
    except:
        print(f"Failed ROI visibility change: {roi.Name}")

# --- Turn on Dose Display ---
try:
    dose_name = plan.TreatmentCourse.TotalDose.OnStructureSet.DicomPlanLabel
    for view in ['Transversal', 'Sagittal', 'Coronal']:
        case.SetDoseVisibility(DoseName=dose_name, View=view, Visible=True)
except Exception as e:
    print(f"Failed to access dose or set dose visibility: {e}")

# --- Screenshot current 2D view ---
try:
    screenshot_path = os.path.join(output_dir, "CurrentView.png")
    ui.SaveScreenShot(FilePath=screenshot_path)
    print(f"Saved: {screenshot_path}")
except:
    print("Could not save screenshot from current view.")

# --- 3D View Screenshots ---
try:
    ui.Open3DView()
    ui.ThreeDView.SetDoseVisibility(True)
    ui.ThreeDView.SetRoiVisibility(True)

    angles = [(0, 0), (90, 0), (0, 90), (45, 45)]
    for idx, (yaw, pitch) in enumerate(angles):
        ui.ThreeDView.SetCameraOrientation(Yaw=yaw, Pitch=pitch)
        screenshot_path = os.path.join(output_dir, f"3D_View_{idx}.png")
        ui.ThreeDView.SaveScreenshot(FilePath=screenshot_path)

    print(f"3D screenshots saved to: {output_dir}")
except:
    print("3D View or screenshot failed.")

# --- Optionally open folder ---
try:
    os.startfile(output_dir)
except:
    pass
