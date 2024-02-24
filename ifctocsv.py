import csv
import os
import sys
import inspect


attributes = ["ADIF_00_Codigo_Producto",
    "ADIF_00_Descripcion_Producto",
    "ADIF_00_Disciplina",
    "ADIF_00_Id_Objeto",
    "ADIF_00_Subdisciplina",
    "ADIF_03_Clasificacion_Elemento",
    "ADIF_03_Codigo_Estudio",
    "ADIF_04_Codigo_Actividad",
    "ADIF_04_Fase",
    "ADIF_05_01_Capitulo",
    "ADIF_05_01_Coeficiente_Certificado",
    "ADIF_05_01_Partida",
    "ADIF_05_01_Unid",
    "ADIF_07_Certificado",
    "ADIF_07_Ejecutado",
    "ADIF_07_Fecha_Ejecucion",
    "ADIF_07_Verificado"
]

def count_elements(ifc_file):
    count = 0
    for _ in ifc_file.by_type('IfcProduct'):
        count += 1
    return count


def export_ifc_to_csv():
    # Getting the path to the current executing script
    script_dir = os.path.dirname(inspect.getfile(inspect.currentframe()))
    # Form the path to the ifcopenshell library relative to the path to the Python executable file
    parent_dir = os.path.dirname(script_dir)
    # Add the path to the ifcopenshell library to sys.path
    resources_dir = os.path.join(parent_dir, "Resources")
    sys.path.append(resources_dir)
    import ifcopenshell
    # Getting a list of .ifc files in the current directory
    ifc_files = [f for f in os.listdir(script_dir) if f.endswith('.ifc')]

    if not ifc_files:
        print("No IFC files")
        return
    total_files = len(ifc_files)
    current_file_num = 0

    for ifc_file_name in ifc_files:
        ifc_file_path = os.path.join(script_dir, ifc_file_name)
        ifc_file = ifcopenshell.open(ifc_file_path)
       
        

        # Creating a name for the .csv file based on the name of the IFC file
        csv_filename = os.path.splitext(ifc_file_name)[0] + ".csv"
        csv_file_path = os.path.join(script_dir, csv_filename)

        total_elements = count_elements(ifc_file)
        current_element = 0

        with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            # Writing the header of the CSV file
            header = ['IfcEntity','GUID', 'Name'] + attributes
            writer.writerow(header)

            elements = ifc_file.by_type('IfcProduct')

            for element in elements:
                IfcEntity = element.is_a()
                guid = element.GlobalId
                name = element.Name if hasattr(element, 'Name') else ''
                psets = ifcopenshell.util.element.get_psets(element)
                
                # # Creating an empty list for attribute values
                attribute_values = []
                # Iterating through each attribute from the header and checking its presence in psets

                for attribute in attributes:
                    found_value = None
                    for group_name, parameters in psets.items():
                        if attribute in parameters:
                            # If the attribute is found, we save its value and exit the loop.
                            found_value = parameters[attribute]
                            break
                    
                    # Adding the attribute value to the list.
                    attribute_values.append(found_value if found_value is not None else '')   
                writer.writerow([IfcEntity, guid, name] + attribute_values)
                current_element += 1
                progress_percentage = (current_element / total_elements) * 100
                print(f"File processing progress {ifc_file_name}: {progress_percentage:.0f} %", end='\r', flush = True)
        current_file_num += 1
        print(f"Export completed for file {ifc_file_name}, {current_file_num} / {total_files}")

export_ifc_to_csv()



