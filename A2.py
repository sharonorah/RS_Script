from connect import *
from System.Drawing import Color
from System.Windows.Forms import (
    Application, Form, Label, Button, CheckBox, ComboBox, DockStyle
)

# Load current case and patient
case = get_current("Case")
patient = get_current("Patient")

class RoiViewerForm(Form):
    def __init__(self):
        self.Text = "Change ROI Visualization"
        self.Width = 400
        self.Height = 350

        # ROI category checkboxes
        self.all_rois_check = CheckBox(Text="All ROIs", Left=20, Top=20, Width=150)
        self.target_check = CheckBox(Text="Target ROIs", Left=20, Top=50, Width=150)
        self.oar_check = CheckBox(Text="OARs", Left=20, Top=80, Width=150)
        self.other_check = CheckBox(Text="Other ROIs", Left=20, Top=110, Width=150)
        self.support_check = CheckBox(Text="Support Structures", Left=20, Top=140, Width=200)

        # Visualization options
        self.visibility_check = CheckBox(Text="Visible", Left=200, Top=20, Width=100)
        self.drr_check = CheckBox(Text="Show DRR Contours", Left=200, Top=50, Width=150)

        self.mode2d_label = Label(Text="2D Mode:", Left=200, Top=90, Width=100)
        self.mode2d_combo = ComboBox(Left=270, Top=90, Width=100)
        self.mode2d_combo.Items.AddRange(["Solid", "Outline", "None"])
        self.mode2d_combo.SelectedIndex = 0

        self.mode3d_label = Label(Text="3D Mode:", Left=200, Top=130, Width=100)
        self.mode3d_combo = ComboBox(Left=270, Top=130, Width=100)
        self.mode3d_combo.Items.AddRange(["Solid", "Wireframe", "None"])
        self.mode3d_combo.SelectedIndex = 0

        # Apply button
        self.apply_button = Button(Text="Apply", Left=150, Top=200, Width=100)
        self.apply_button.Click += self.apply_settings

        # Status label
        self.status_label = Label(Text="", Left=20, Top=240, Width=340, Height=40)

        # Add all controls to the form
        for control in [
            self.all_rois_check, self.target_check, self.oar_check,
            self.other_check, self.support_check, self.visibility_check,
            self.drr_check, self.mode2d_label, self.mode2d_combo,
            self.mode3d_label, self.mode3d_combo, self.apply_button,
            self.status_label
        ]:
            self.Controls.Add(control)

    def apply_settings(self, sender, event):
        try:
            count = 0

            process_all = self.all_rois_check.Checked
            process_targets = self.target_check.Checked or process_all
            process_oars = self.oar_check.Checked or process_all
            process_other = self.other_check.Checked or process_all
            process_support = self.support_check.Checked or process_all

            is_visible = self.visibility_check.Checked
            show_drr = self.drr_check.Checked
            mode_2d = self.mode2d_combo.SelectedItem
            mode_3d = self.mode3d_combo.SelectedItem

            target_keywords = ['target', 'ptv', 'ctv', 'gtv']
            oar_keywords = ['organ', 'organ at risk', 'oar']
            support_keywords = ['support']

            for roi in case.PatientModel.RegionsOfInterest:
                roi_name = roi.Name
                roi_type = roi.Type.lower()

                should_process = (
                    process_all or
                    (process_targets and any(k in roi_type for k in target_keywords)) or
                    (process_oars and any(k in roi_type for k in oar_keywords)) or
                    (process_support and any(k in roi_type for k in support_keywords)) or
                    (process_other and not any(k in roi_type for k in target_keywords + oar_keywords + support_keywords))
                )

                if not should_process:
                    continue

                try:
                    vis = roi.RoiVisualizationSettings
                    vis.IsVisible = is_visible
                    vis.ShowDRRContours = show_drr
                    vis.VisualizationMode2D = mode_2d
                    vis.VisualizationMode3D = mode_3d
                    count += 1
                except Exception as e:
                    print(f"Error applying settings to {roi_name}: {e}")

            patient.Save()
            self.status_label.Text = f"Successfully updated {count} ROIs!"
            self.status_label.ForeColor = Color.Green

        except Exception as e:
            self.status_label.Text = f"Error: {str(e)}"
            self.status_label.ForeColor = Color.Red

# Run the form
form = RoiViewerForm()
Application.Run(form)
