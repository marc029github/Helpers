"""
    Read an xls file that contains two columns: name_image, type. Then
    creates a folder for each type and moves the images in another
    folder called "images" to the corresponding folder of the type.

    Author: Marcos Fernandez
    Date: 10/07/2020

"""

import math
import os
import pandas as pd


def SortImagesinFoldersFromXLS(xlsfile, inputImagesFolder, outputImagesFolder, clnImageName, clnType):

    try:
        # Read the xls file and get the data
        df = pd.read_excel(xlsfile)
        data = df.to_dict('records')

        # Define constant variable
        TIPO_MATERIAL = clnType
        TIPO_MATERIAL_FOLDER = clnType.replace(" ", "_")

        # Create the folders
        for d in data:
            if not math.isnan(d[TIPO_MATERIAL]):
                if not os.path.exists(f'{outputImagesFolder}/{TIPO_MATERIAL_FOLDER}/{str(int(d[TIPO_MATERIAL]))}'):
                    os.makedirs(f'{outputImagesFolder}/{TIPO_MATERIAL_FOLDER}/{str(int(d[TIPO_MATERIAL]))}')

        # Copy image to the corresponding folder
        for d in data:
            if not math.isnan(d[TIPO_MATERIAL]):
                if os.path.exists(f'{inputImagesFolder}/{d[clnImageName]}'):
                    os.system(f'cp {inputImagesFolder}/{d[clnImageName]} {outputImagesFolder}/{TIPO_MATERIAL_FOLDER}/{str(int(d[TIPO_MATERIAL]))}') 

        return "Done!"
    
    except Exception as e:
        print(f'An error occurred: {str(e)}')
        return f'Error: {str(e)}'


cln1 = SortImagesinFoldersFromXLS(xlsfile='caracteristicas.xlsx', 
                           inputImagesFolder='imagenes_apoyos', 
                           outputImagesFolder='img_sorted', 
                           clnImageName='Nombre',
                            clnType= 'Tipo Material')
print("Resultado campo 1: ", cln1)

cln2 = SortImagesinFoldersFromXLS(xlsfile='caracteristicas.xlsx',                        
                           inputImagesFolder='imagenes_apoyos', 
                           outputImagesFolder='img_sorted', 
                           clnImageName='Nombre',
                            clnType= 'Subtipo Material')
print("Resultado campo 2: ", cln2)

cln3 = SortImagesinFoldersFromXLS(xlsfile='caracteristicas.xlsx',                        
                           inputImagesFolder='imagenes_apoyos', 
                           outputImagesFolder='img_sorted', 
                           clnImageName='Nombre',
                            clnType= 'Función Normal')
print("Resultado campo 3 ", cln3)

cln4 = SortImagesinFoldersFromXLS(xlsfile='caracteristicas.xlsx',                        
                           inputImagesFolder='imagenes_apoyos', 
                           outputImagesFolder='img_sorted', 
                           clnImageName='Nombre',
                            clnType= 'Función Especial')
print("Resultado campo 4: ", cln4)