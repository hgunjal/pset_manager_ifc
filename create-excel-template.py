import ifcopenshell
import pandas as pd
import os
from config import file_path, model_name, ifc_folder_path, output_excel_path

# Load the IFC file
ifc_file = ifcopenshell.open(file_path)

# Extract entities (GUID, Name, and Type)
data = []
for element in ifc_file.by_type('IfcProduct'):
    if element.is_a('IfcProduct') and element.Name:
        data.append({
            "GUID": element.GlobalId,
            "Name": element.Name,
            "Type": element.is_a(),
            "4D-Vorgangs-ID": "X",      # Placeholder for 4D-Vorgangs-ID
            "Terminplan-ID": "X",       # Placeholder for Terminplan-ID
            "Objekt": "X",              # Placeholder for Objekt
            "Strecke": "X",             # Placeholder for Strecke
            "Bauabschnitt": "X",        # Placeholder for Bauabschnitt
            "Bauwerk": "X",             # Placeholder for Bauwerk
            "Gewerk": "X",              # Placeholder for Gewerk
            "Gleis": "X",               # Placeholder for Gleis
            "Stationierung": "X",       # Placeholder for Stationierung
            "StationierungBis": "X",    # Placeholder for StationierungBis
            "StationierungVon": "X",    # Placeholder for StationierungVon
            "PSPElement": "X",          # Placeholder for PSPElement
            "BauphaseAbgebrochen": "X", # Placeholder for BauphaseAbgebrochen
            "BauphaseErstellt": "X",    # Placeholder for BauphaseErstellt
            "Detail": "X",              # Placeholder for Detail
            "Eigentuemer": "X",         # Placeholder for Eigentuemer
            "Kommentar": "X",           # Placeholder for Kommentar
            "Regelwerksnummer": "X",    # Placeholder for Regelwerksnummer
            "Richtzeichnung": "X",      # Placeholder for Richtzeichnung
            "Zustand": "X"              # Placeholder for Zustand
        })

# Convert data to DataFrame
df = pd.DataFrame(data)



# Export to Excel
df.to_excel(output_excel_path, index=False)

print(f"IFC data exported to {output_excel_path} with relevant columns for property sets.")
