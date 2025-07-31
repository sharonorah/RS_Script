from connect import *
import clr

# Add necessary .NET references for the GUI
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Drawing import Point, Color, Size
from System.Windows.Forms import (Application, Button, Form, Label, CheckBox, 
                                 ComboBox, Panel, RadioButton, GroupBox)

# Get current context
case = get_current('Case')
patient = get_current('Patient')

class ROIVisualizationForm(Form):
    def __init__(self):
        # Form settings
        self.Text = "ROI Visualization Settings"
        self.Width = 400
        self.Height = 350
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        
        # Visibility settings
        y_pos = 20
        
        self.visibility_check = CheckBox()
        self.visibility_check.Text = "Make ROIs Visible"
        self.visibility_check.Checked = True
        self.visibility_check.Location = Point(20, y_pos)
        self.visibility_check.Width = 250
        self.Controls.Add(self.visibility_check)
        
        y_pos += 30
        
        self.drr_check = CheckBox()
        self.drr_check.Text = "Show DRR Contours"
        self.drr_check.Checked = True
        self.drr_check.Location = Point(20, y_pos)
        self.drr_check.Width = 250
        self.Controls.Add(self.drr_check)
        
        # 2D visualization mode
        y_pos += 40
        
        label_2d = Label()
        label_2d.Text = "2D Visualization Mode:"
        label_2d.Location = Point(20, y_pos)
        label_2d.Width = 150
        self.Controls.Add(label_2d)
        
        self.mode2d_combo = ComboBox()
        self.mode2d_combo.Location = Point(180, y_pos)
        self.mode2d_combo.Width = 180
        self.mode2d_combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.mode2d_combo.Items.AddRange(["Contour", "Filled", "FilledContour", "Off"])
        self.mode2d_combo.SelectedIndex = 0
        self.Controls.Add(self.mode2d_combo)
        
        # 3D visualization mode
        y_pos += 40
        
        label_3d = Label()
        label_3d.Text = "3D Visualization Mode:"
        label_3d.Location = Point(20, y_pos)
        label_3d.Width = 150
        self.Controls.Add(label_3d)
        
        self.mode3d_combo = ComboBox()
        self.mode3d_combo.Location = Point(180, y_pos)
        self.mode3d_combo.Width = 180
        self.mode3d_combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.mode3d_combo.Items.AddRange(["Shaded", "SemiTransparent", "ShadedWireframe", "Off"])
        self.mode3d_combo.SelectedIndex = 2
        self.Controls.Add(self.mode3d_combo)
        
        # ROI type filter (optional)
        y_pos += 40
        
        group_roi_types = GroupBox()
        group_roi_types.Text = "Apply to ROI Types:"
        group_roi_types.Location = Point(20, y_pos)
        group_roi_types.Size = Size(350, 80)
        self.Controls.Add(group_roi_types)
        
        self.all_rois_check = CheckBox()
        self.all_rois_check.Text = "All ROIs"
        self.all_rois_check.Checked = True
        self.all_rois_check.Location = Point(20, 20)
        self.all_rois_check.Width = 100
        self.all_rois_check.CheckedChanged += self.all_rois_checked_change
        group_roi_types.Controls.Add(self.all_rois_check)
        
        self.target_check = CheckBox()
        self.target_check.Text = "Targets"
        self.target_check.Enabled = False
        self.target_check.Location = Point(130, 20)
        self.target_check.Width = 100
        group_roi_types.Controls.Add(self.target_check)
        
        self.oar_check = CheckBox()
        self.oar_check.Text = "OARs"
        self.oar_check.Enabled = False
        self.oar_check.Location = Point(230, 20)
        self.oar_check.Width = 100
        group_roi_types.Controls.Add(self.oar_check)
        
        self.other_check = CheckBox()
        self.other_check.Text = "Other"
        self.other_check.Enabled = False
        self.other_check.Location = Point(20, 50)
        self.other_check.Width = 100
        group_roi_types.Controls.Add(self.other_check)
        
        self.support_check = CheckBox()
        self.support_check.Text = "Support"
        self.support_check.Enabled = False
        self.support_check.Location = Point(130, 50)
        self.support_check.Width = 100
        group_roi_types.Controls.Add(self.support_check)
        
        # Apply button
        y_pos += 100
        
        self.apply_button = Button()
        self.apply_button.Text = "Apply Settings"
        self.apply_button.Location = Point(120, y_pos)
        self.apply_button.Width = 150
        self.apply_button.Height = 30
        self.apply_button.BackColor = Color.LightGreen
        self.apply_button.Click += self.apply_settings
        self.Controls.Add(self.apply_button)
        
        # Status label
        y_pos += 40
        
        self.status_label = Label()
        self.status_label.Text = ""
        self.status_label.Location = Point(20, y_pos)
        self.status_label.Width = 350
        self.status_label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        self.Controls.Add(self.status_label)
    
    def all_rois_checked_change(self, sender, event):
        """Enable/disable ROI type filters based on All ROIs checkbox"""
        self.target_check.Enabled = not self.all_rois_check.Checked
        self.oar_check.Enabled = not self.all_rois_check.Checked
        self.other_check.Enabled = not self.all_rois_check.Checked
        self.support_check.Enabled = not self.all_rois_check.Checked
    
    def apply_settings(self, sender, event):
        """Apply the selected visualization settings to ROIs"""
        try:
            count = 0
            
            # Determine which ROI types to process
            process_all = self.all_rois_check.Checked
            process_targets = self.target_check.Checked or process_all
            process_oars = self.oar_check.Checked or process_all
            process_other = self.other_check.Checked or process_all
            process_support = self.support_check.Checked or process_all
            
            # Get visualization mode settings
            is_visible = self.visibility_check.Checked
            show_drr = self.drr_check.Checked
            mode_2d = self.mode2d_combo.SelectedItem
            mode_3d = self.mode3d_combo.SelectedItem
            
            for roi in case.PatientModel.RegionsOfInterest:
                roi_type = roi.Type.lower()
                
                # Check if this ROI type should be processed
                should_process = False
                if process_all:
                    should_process = True
                elif process_targets and ('Target' in roi_type or 'ptv' in roi_type or 'ctv' in roi_type or 'gtv' in roi_type):
                    should_process = True
                elif process_oars and ('Organ' in roi_type or 'Organ at risk' in roi_type):
                    should_process = True
                elif process_support and 'Support' in roi_type:
                    should_process = True
                elif process_other and not ('Target' in roi_type or 'ptv' in roi_type or 'ctv' in roi_type or 'gtv' in roi_type or 'organ' in roi_type or 'oar' in roi_type or 'support' in roi_type):
                    should_process = True
                
                if should_process:
                    try:
                        vis = roi.RoiVisualizationSettings
                        vis.IsVisible = is_visible
                        vis.ShowDRRContours = show_drr
                        vis.VisualizationMode2D = mode_2d
                        vis.VisualizationMode3D = mode_3d
                        count += 1
                        print(f"Applying settings to {roi.Name}")
                    except Exception as e:
                        print(f"Error applying settings to {roi.Name}: {e}")
            
            # Save changes
            patient.Save()
            
            self.status_label.Text = f"Successfully updated {count} ROIs!"
            self.status_label.ForeColor = Color.Green
        
        except Exception as e:
            self.status_label.Text = f"Error: {str(e)}"
            self.status_label.ForeColor = Color.Red

# Run the application
form = ROIVisualizationForm()
Application.Run(form)