from connect import get_current
import os
import clr
import datetime

# ---------- Select folder with file dialog ----------
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import FolderBrowserDialog, DialogResult, MessageBox

folder_dialog = FolderBrowserDialog()
folder_dialog.Description = "Select folder to save screenshots"
dialog_result = folder_dialog.ShowDialog()

if dialog_result != DialogResult.OK:
    raise Exception("Screenshot export cancelled by user.")

output_dir = folder_dialog.SelectedPath

if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# ---------- Logging setup ----------
log_path = os.path.join(output_dir, "screenshot_log.txt")
def log(msg):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, 'a') as log_file:
        log_file.write(f"[{timestamp}] {msg}\n")

log("=== Starting Screenshot Script ===")

# ---------- RayStation setup ----------
try:
    case = get_current("Case")
    examination = get_current("Examination")
    ui = get_current("ui")
    plan = case.TreatmentPlans[0]
    log("Accessed current case and plan.")
except Exception as e:
    log(f"Failed to access RayStation context: {e}")
    raise

# ---------- Set ROI visibility ----------
visible_rois = ['Liver', 'GTV', 'CTV', 'PTV', 'SpinalCord', 'Kidneys']

for roi in case.PatientModel.RegionsOfInterest:
    try:
        show = roi.Name in visible_rois
        roi.ReviewDisplayColor = 'Red' if show else 'Gray'
        for view in ['Transversal', 'Sagittal', 'Coronal']:
            case.SetRoiVisibility(RoiName=roi.Name, View=view, Visible=show)
        log(f"Set visibility for ROI: {roi.Name} - {'Visible' if show else 'Hidden'}")
    except Exception as e:
        log(f"Failed to set ROI visibility for {roi.Name}: {e}")

# ---------- Turn on dose display ----------
try:
    dose_name = plan.TreatmentCourse.TotalDose.OnStructureSet.DicomPlanLabel
    for view in ['Transversal', 'Sagittal', 'Coronal']:
        case.SetDoseVisibility(DoseName=dose_name, View=view, Visible=True)
        log(f"Dose visibility ON for view: {view}")
except Exception as e:
    log(f"Failed to set dose visibility: {e}")

# ---------- Screenshot 2D current view ----------
try:
    screenshot_path = os.path.join(output_dir, "CurrentView.png")
    log(f"Saving 2D screenshot to: {screenshot_path}")
    ui.SaveScreenShot(FilePath=screenshot_path)
    log(f"Saved 2D screenshot: {screenshot_path}")
except Exception as e:
    log(f"2D screenshot failed: {e}")

# ---------- 3D View screenshots ----------
try:
    ui.Open3DView()
    ui.ThreeDView.SetDoseVisibility(True)
    ui.ThreeDView.SetRoiVisibility(True)
    log("Opened 3D view and enabled dose/ROI display.")

    angles = [(0, 0), (90, 0), (0, 90), (45, 45)]
    for idx, (yaw, pitch) in enumerate(angles):
        try:
            ui.ThreeDView.SetCameraOrientation(Yaw=yaw, Pitch=pitch)
            screenshot_path = os.path.join(output_dir, f"3D_View_{idx}.png")
            ui.ThreeDView.SaveScreenshot(FilePath=screenshot_path)
            log(f"Saved 3D screenshot {idx}: {screenshot_path}")
        except Exception as e:
            log(f"3D screenshot {idx} failed: {e}")
except Exception as e:
    log(f"3D view setup failed: {e}")

# ---------- Completion popup ----------
MessageBox.Show("Screenshots completed.\nCheck:\n" + output_dir, "Done")
log("=== Screenshot script complete ===")
