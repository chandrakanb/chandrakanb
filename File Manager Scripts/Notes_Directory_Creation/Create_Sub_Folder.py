import os

def create_folder_structure(base_folder, subfolders):
    # Create the base folder in the current directory
    if not os.path.exists(base_folder):
        os.makedirs(base_folder)
        print(f"Created base folder: {base_folder}")
    else:
        print(f"Base folder already exists: {base_folder}")

    # Create subfolders within the base folder
    for folder in subfolders:
        subfolder_path = os.path.join(base_folder, folder)
        if not os.path.exists(subfolder_path):
            os.makedirs(subfolder_path)
            print(f"Created subfolder: {subfolder_path}")
        else:
            print(f"Subfolder already exists: {subfolder_path}")

# Example usage for each subject:

# Folder structure for "Fundamentals_of_Data_Structures"
base_folder_name = "Fundamentals_of_Data_Structures(210242)"
subfolder_list = [
    "Unit_I_Introduction_to_Algorithm_and_Data_Structures",
    "Unit_II_Linear_Data_Structure_Using_Sequential_Organization",
    "Unit_III_Searching_and_Sorting",
    "Unit_IV_Linked_List",
    "Unit_V_Stack",
    "Unit_VI_Queue"
]
create_folder_structure(base_folder_name, subfolder_list)

# Folder structure for "Discrete_Mathematics"
base_folder_name = "Discrete_Mathematics(210241)"
subfolder_list = [
    "Unit_I_Set_Theory_and_Logic",
    "Unit_II_Relations_and_Functions",
    "Unit_III_Counting_Principle",
    "Unit_IV_Graph_Theory_",
    "Unit_V_Trees",
    "Unit_VI_Algebraic_Structures_and_Coding_Theory"
]
create_folder_structure(base_folder_name, subfolder_list)

# Folder structure for "Digital_Electronics_and_Logic_Design"
base_folder_name = "Digital_Electronics_and_Logic_Design(210245)"
subfolder_list = [
    "Unit_I_Minimization_Technique",
    "Unit_II_Combinational_Logic_Design",
    "Unit_III_Sequential_Logic_Design",
    "Unit_IV_Algorithmic_State_Machines_and_Programmable_Logic_Devices",
    "Unit_V_Logic_Families",
    "Unit_VI_Introduction_to_Computer_Architecture"
]
create_folder_structure(base_folder_name, subfolder_list)

# Folder structure for "Object_Oriented_Programming"
base_folder_name = "Object_Oriented_Programming(210243)"
subfolder_list = [
    "Unit_I_Fundamentals_of_Object_Oriented_Programming",
    "Unit_II_Inheritance_and_Pointers",
    "Unit_III_Polymorphism",
    "Unit_IV_Files_and_Streams",
    "Unit_V_Exception_Handling_&_Templates",
    "Unit_VI_Standard_Template_Library_(STL)"
]
create_folder_structure(base_folder_name, subfolder_list)

# Folder structure for "Computer_Graphics"
base_folder_name = "Computer_Graphics(210244)"
subfolder_list = [
    "Unit_I_Graphics_Primitives_and_Scan_Conversion_Algorithms",
    "Unit_II_Polygon,_Windowing_and_Clipping",
    "Unit_III_2D,_3D_Transformations_and_Projections",
    "Unit_IV_Light,_Colour,_Shading_and_Hidden_Surfaces",
    "Unit_V_Curves_and_Fractals",
    "Unit_VI_Introduction_to_Animation_and_Gaming"
]
create_folder_structure(base_folder_name, subfolder_list)
