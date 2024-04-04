import pandas as pd
import os
import re
from tkinter import filedialog, messagebox
from tkinter import *


root = Tk()
root.withdraw()  # Ocultar la ventana principal de Tkinter
def Reactor():
    
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    # Salir del programa si no se selecciona ningún archivo
    if not file_path:
        exit()

    try:
        # Cargar el archivo Excel en un DataFrame y se elimina las 3 primeras filas    
        df = pd.read_excel(file_path, sheet_name='XXMU___Reporte_de_gastos_de_Ma_',skiprows=3)
        # Se especifica que las columnas 'VR Transaccion' y 'Fecha Transacción' sean de tipo str directamente al cargar el archivo
        df = df.astype({'VR Transaccion': str, 'Fecha Transacción': str})

    
        # Imprimir los nombres de las columnas y los tipos de datos
        print("Nombres de columna:", list(df.columns))
        print(df.dtypes)

        # Patrones a buscar en la columna "Descripción"
        patrones = [r'R1\d\d', r'R2\d\d', r'CR[1-6]',r'MZ11',r'MZ[1-4]',r'MZ',r'CW[1-3]']

        # Compilar los patrones en expresiones regulares
        regex_patrones = [re.compile(patron) for patron in patrones]

        # Verificar si la columna "Descripción" existe en el DataFrame
        if "Desc. Activo" in df.columns:
            # Crear una nueva columna "Equipo"
            df["Reactor"] = ""

            # Buscar los patrones en la columna "Descripción" e insertar el resultado en la nueva columna "Equipo"
            for index, row in df.iterrows():
                descripcion = row["Desc. Activo"]
                for patron in regex_patrones:
                    if not isinstance(descripcion, str):
                        descripcion = str(descripcion)
                    match = patron.search(descripcion)
                    if match:
                        df.at[index, "Reactor"] = match.group()
                        break  # Detener la búsqueda después de encontrar el primer patrón coincidente

        else:
            messagebox.showerror(message="La columna 'Desc. Activo' no existe en el archivo.", title="ERROR")
        
        
    
        # Cambiar el tipo de datos de la columna 
        df['VR Transaccion'] = df['VR Transaccion'].astype('int64')
        print("Nombres de columna:", list(df.columns))
        print(df.dtypes)



        #SE GUARDA EL RESULTADO
        with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="XXMU___Reporte_de_gastos_de_Ma_", index=False)
        messagebox.showinfo(message="El archivo ha sido actualizado correctamente, seleccione el siguiente", title="EXITO")


        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

        if not file_path:
            exit()


        # Cargar el archivo Excel en un DataFrame    
        dt = pd.read_excel(file_path)
        
        # Concatenar los DataFrames en un solo DataFrame
        excel_merged = pd.concat([df, dt], ignore_index=True)

        #SE GUARDA EL RESULTADO
        with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
            excel_merged.to_excel(writer, index=False)
        messagebox.showinfo(message="El archivo ha sido actualizado correctamente.", title="EXITO")

    
    except FileNotFoundError:
        messagebox.showerror(message="El archivo seleccionado no se encontró.", title="ERROR")
    except :
        messagebox.showerror(message="Error", title="ERROR")

   

def repetir_ciclo():
    continuar = True
    while continuar:
        print(Reactor())

        respuesta = messagebox.askyesno("Pregunta", "¿Deseas continuar?")
        if respuesta:
            # Continuar el ciclo
            pass
        else:
            continuar = False  # Salir del ciclo


# Llamar a la función repetir_ciclo() cuando se inicie la aplicación
repetir_ciclo()
    
    



