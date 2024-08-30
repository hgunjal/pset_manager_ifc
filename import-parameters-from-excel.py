import ifcopenshell
import ifcopenshell.api.pset
import pandas as pd
import os
from config import file_path, model_name, ifc_folder_path, output_excel_path

# Load the IFC file
ifc_file = ifcopenshell.open(file_path)

# Load the Excel file
df = pd.read_excel(output_excel_path)

# Iterate through the DataFrame and add property sets
for index, row in df.iterrows():
    guid = row['GUID']
    element = ifc_file.by_guid(guid)

    if element:
        # Create or retrieve the Pset
        pset = ifcopenshell.api.pset.add_pset(ifc_file, product=element, name="Allgemein")

        # Prepare properties and filter out invalid ones
        properties = {}

        for key in [
            "4D-Vorgangs-ID", "Terminplan-ID", "Objekt", "Strecke", "Bauabschnitt",
            "Bauwerk", "Gewerk", "Gleis", "Stationierung", "StationierungBis",
            "StationierungVon", "PSPElement", "BauphaseAbgebrochen", "BauphaseErstellt",
            "Detail", "Eigentuemer", "Kommentar", "Regelwerksnummer",
            "Richtzeichnung", "Zustand"
        ]:
            value = row.get(key)
            if value and pd.notna(value):  # Check if value is not None or NaN
                properties[key] = value

        # Add properties to the Pset
        if properties:  # Ensure there are valid properties to add
            ifcopenshell.api.pset.edit_pset(ifc_file, pset=pset, properties=properties)

# Construct the output file path
output_excel_path = os.path.join(ifc_folder_path, f"{model_name}_ifc_param_manager.xlsx")

# Save the modified IFC file
output_file = os.path.join(ifc_folder_path, f"{model_name}_attb.ifc")
ifc_file.write(output_file)

print(f"Property sets added and IFC file saved as {output_file}")
