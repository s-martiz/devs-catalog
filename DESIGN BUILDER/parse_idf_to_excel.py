import os
import pandas as pd
from eppy.modeleditor import IDF

# Fonction pour charger le fichier IDF
def load_idf(file_path):
    IDF.setiddname("C:/EnergyPlusV9-4-0/Energy+.idd")  # Chemin vers votre fichier EnergyPlus IDD
    return IDF(file_path)

# Fonction pour récupérer le premier fichier .idf du répertoire
def find_idf_file(directory):
    for file in os.listdir(directory):
        if file.endswith(".idf"):
            return os.path.join(directory, file)
    return None

# Récupérer le répertoire actuel
current_dir = os.getcwd()

# Trouver le premier fichier .idf dans le répertoire
idf_file = find_idf_file(current_dir)

# Vérifier si un fichier .idf a été trouvé
idf_model = load_idf(idf_file)

# Extraction des propriétés thermiques et hygroscopiques de tous les matériaux dans les constructions
def extract_construction_properties(idf_model):
    construction = idf_model.idfobjects["CONSTRUCTION"]
    construction_data = []

    # Parcourir toutes les constructions
    for const in construction:
        construction_dict = {
            "Construction Name": const.Name,
            "Layer Names": []
        }

        # Traiter l'extérieur et les couches internes
        layer_names = [const.Outside_Layer] + [getattr(const, f"Layer_{i}", None) for i in range(1, 6)]

        for layer_name in layer_names:
            if layer_name:
                mat_layer = idf_model.getobject(key="MATERIAL", name=layer_name)
                if mat_layer:  # Vérifier si le matériau existe
                    layer_data = extract_layer_data(idf_model, layer_name)
                    construction_dict["Layer Names"].append(layer_data)

        construction_data.append(construction_dict)

    return construction_data

# Extraction des propriétés de chaque matériau (thermiques et hygroscopiques)
def extract_layer_data(idf_model, layer_name):
    mat_layer = idf_model.getobject(key="MATERIAL", name=layer_name)

    # Extraction des propriétés de base
    layer_data = {
        "Material Name": mat_layer.Name,
        "Thickness": mat_layer.Thickness,
        "Conductivity": mat_layer.Conductivity,
        "Density": mat_layer.Density,
        "Specific_Heat": mat_layer.Specific_Heat,
        "ThermalConductivity_x": extract_hygro_data(idf_model, layer_name, "THERMALCONDUCTIVITY", "x"),
        "ThermalConductivity_y": extract_hygro_data(idf_model, layer_name, "THERMALCONDUCTIVITY", "y"),
        "Diffusion_x": extract_hygro_data(idf_model, layer_name, "DIFFUSION", "x"),
        "Diffusion_y": extract_hygro_data(idf_model, layer_name, "DIFFUSION", "y"),
        "Redistribution_x": extract_hygro_data(idf_model, layer_name, "REDISTRIBUTION", "x"),
        "Redistribution_y": extract_hygro_data(idf_model, layer_name, "REDISTRIBUTION", "y"),
        "Suction_x": extract_hygro_data(idf_model, layer_name, "SUCTION", "x"),
        "Suction_y": extract_hygro_data(idf_model, layer_name, "SUCTION", "y"),
        "Sorption_x": extract_hygro_data(idf_model, layer_name, "SORPTIONISOTHERM", "x"),
        "Sorption_y": extract_hygro_data(idf_model, layer_name, "SORPTIONISOTHERM", "y")
    }

    return layer_data

# Extraction des données hygroscopiques pour chaque propriété (ThermalConductivity, Diffusion, etc.)
def extract_hygro_data(idf_model, layer_name, property_name, axis):
    hygro_data = []
    coords = {
        "THERMALCONDUCTIVITY": {"x": "Moisture_Content", "y": "Thermal_Conductivity"},
        "DIFFUSION": {"x": "Relative_Humidity_Fraction", "y": "Water_Vapor_Diffusion_Resistance_Factor"},
        "REDISTRIBUTION": {"x": "Moisture_Content", "y": "Liquid_Transport_Coefficient"},
        "SUCTION": {"x": "Moisture_Content", "y": "Liquid_Transport_Coefficient"},
        "SORPTIONISOTHERM": {"x": "Relative_Humidity_Fraction", "y": "Moisture_Content"}
    }

    if property_name in coords:
        coords_dict = coords[property_name]
        to_list = idf_model.getobject(key=f"MATERIALPROPERTY:HEATANDMOISTURETRANSFER:{property_name}", name=layer_name)
        
        for key, val in zip(to_list.objls, to_list.obj):
            if coords_dict["x"] in key and axis == "x":
                hygro_data.append(val)  # Ajouter les valeurs de Moisture_Content ou Relative_Humidity_Fraction
            elif coords_dict["y"] in key and axis == "y":
                hygro_data.append(val)  # Ajouter les valeurs de Thermal_Conductivity ou Water_Vapor_Diffusion_Resistance_Factor

    return hygro_data

# Enregistrement des données dans un fichier Excel avec répartition des listes sur des lignes séparées
def save_to_excel(construction_data, output_filename):
    with pd.ExcelWriter(output_filename) as writer:
        for construction in construction_data:
            construction_name = construction["Construction Name"]
            layers_data = construction["Layer Names"]
            
            # Préparer un DataFrame pour chaque construction
            data_rows = []

            # Trouver la longueur maximale des listes pour éviter les erreurs
            max_len = 0
            for layer in layers_data:
                for prop in ["Sorption_x", "Sorption_y", "ThermalConductivity_x", "ThermalConductivity_y", "Diffusion_x", "Diffusion_y", "Redistribution_x", "Redistribution_y", "Suction_x", "Suction_y"]:
                    if len(layer.get(prop, [])) > max_len:
                        max_len = len(layer.get(prop, []))

            for i in range(max_len):
                for layer in layers_data:
                    row = {
                        "Material Name": layer["Material Name"],
                        "Thickness": layer["Thickness"],
                        "Conductivity": layer["Conductivity"],
                        "Density": layer["Density"],
                        "Specific_Heat": layer["Specific_Heat"],
                    }

                    # Remplir les colonnes pour chaque propriété hygroscopique (avec valeurs par défaut si la liste est plus courte)
                    for prop in ["Sorption_x", "Sorption_y", "ThermalConductivity_x", "ThermalConductivity_y", "Diffusion_x", "Diffusion_y", "Redistribution_x", "Redistribution_y", "Suction_x", "Suction_y"]:
                        if i < len(layer.get(prop, [])):
                            row[prop] = layer[prop][i]
                        else:
                            row[prop] = None  # Valeur manquante si la liste est plus courte

                    data_rows.append(row)

            # Trier les matériaux et propriétés par "Material Name" et par épaisseur
            layer_df = pd.DataFrame(data_rows)
            
            # Vérification avant tri, éviter de trier si les colonnes n'existent pas
            if "Material Name" in layer_df.columns and "Thickness" in layer_df.columns:
                layer_df = layer_df.sort_values(by=["Material Name", "Thickness"])  # Trier par Material Name et par épaisseur
            
            layer_df.to_excel(writer, sheet_name=construction_name[:31], index=False)

# Récupérer le répertoire actuel
current_dir = os.getcwd()

# Trouver le premier fichier .idf dans le répertoire
idf_file = find_idf_file(current_dir)

# Vérifier si un fichier .idf a été trouvé
if idf_file:
    # Charger le modèle IDF
    idf_model = load_idf(idf_file)

    # Extraire les propriétés des constructions
    construction_data = extract_construction_properties(idf_model)

    # Sauvegarder les données dans un fichier Excel
    save_to_excel(construction_data, "Construction_Properties.xlsx")
    print("Les données ont été enregistrées dans 'Construction_Properties.xlsx'.")
else:
    print("Aucun fichier .idf trouvé dans le répertoire.")
