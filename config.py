import os

# Load the IFC file
file_path = r'C:\Users\Harshal.Gunjal\OneDrive - ILF Group Holding GmbH\Dokumente\Blender-BIM_Scripts\pset_scripts\Kißlegg_OLA_Mast_mit_Ausleger.ifc'
model_name, _ = os.path.splitext(os.path.basename(file_path))


# Define the Excel file directory as the same as the IFC file directory
ifc_folder_path = os.path.dirname(file_path)

# Construct the output file path
output_excel_path = os.path.join(ifc_folder_path, f"{model_name}_ifc_param_manager.xlsx")