from openpyxl import Workbook
import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser
from modules.file_handler import leer_archivo_txt
from modules.excel_formatter import ajustar_tamanio_columnas, formatear_celdas
from modules.excel_writer import procesar_datos

def generar_excel(input_path, output_path):
    """
    Función máscara para generar un archivo Excel a partir de un archivo de texto.

    Args:
        input_path (str): Ruta del archivo de texto
        output_path (str): Ruta del archivo Excel

    Returns:
        e (bool): Falso si falló algo en la ejecución, Verdadero si se generó el archivo correctamente
    """
    try:
        wb = Workbook()
        ws = wb.active

        lineas = leer_archivo_txt(input_path)

        procesar_datos(lineas, ws, wb)
        
        ajustar_tamanio_columnas(ws)

        formatear_celdas(ws)
        
        wb.save(output_path)
        print("Documento output.xlsx generado correctamente.")
        return True
    except Exception as e:
        print("Error:", e)
        return False

def seleccionar_archivo_entrada(entry):
    """
    Abre un cuadro de diálogo para seleccionar un archivo de texto.

    Args:
        entry (tk.Entry): Campo de texto de la ventana principal
    """
    ruta = filedialog.askopenfilename(title="Seleccionar archivo de texto", filetypes=[("Archivos de texto", "*.txt")])
    if ruta:
        entry.delete(0, tk.END)
        entry.insert(0, ruta)

def seleccionar_archivo_salida(entry):
    """
    Abre un cuadro de diálogo para seleccionar un archivo de Excel.

    Args:
        entry (tk.Entry): Campo de texto de la ventana principal
    """
    ruta = filedialog.asksaveasfilename(title="Guardar archivo de Excel", filetypes=[("Archivos de Excel", "*.xlsx")])
    if ruta:
        entry.delete(0, tk.END)
        entry.insert(0, ruta)

def iniciar_proceso(input_entry, output_entry):
    """
    Inicia el proceso de generación de un archivo Excel a partir de un archivo de texto.

    Args:
        input_entry (tk.Entry): Campo de texto de la ventana principal
        output_entry (tk.Entry): Campo de texto de la ventana principal
    """
    input_path = input_entry.get()
    output_path = output_entry.get()

    if not input_path or not output_path:
        messagebox.showerror("Error", "Debes seleccionar un archivo de entrada y otro de salida.")
        return

    messagebox.showinfo("Proceso en curso", "El proceso de generación del archivo de Excel ha comenzado, por favor espera. Este proceso puede tardar unos minutos.")

    if generar_excel(input_path, output_path):
        messagebox.showinfo("Proceso completado", "El archivo de Excel se generó correctamente.")
        return
    else:
        messagebox.showerror("Error", "Ocurrió un error al generar el archivo de Excel.")

def abrir_repositorio():
    webbrowser.open("https://github.com/DarioSuarezL/txt_to_xlsx_converter")

def acerca_de():
    frame = tk.Toplevel()
    frame.title("Acerca de")
    tk.Label(frame, text="Versión: 0.4.0").pack(padx=10, pady=10)
    tk.Label(frame, text="Autor: Univ. Darío Suárez Lazarte").pack(padx=10)
    tk.Label(frame, text="Correo: dsuarezlazarte@gmail.com").pack(padx=10)
    tk.Button(frame, text="Ir al repositorio", command=lambda: abrir_repositorio()).pack(padx=10, pady=10)



def iniciar_gui():
    root = tk.Tk()
    root.title("Conversor de archivos de texto a Excel")

    # Configuración de interfaz
    tk.Label(root, text="Archivo de entrada:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    input_entry = tk.Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=10, pady=10)
    input_button = tk.Button(root, text="Seleccionar archivo txt", command=lambda: seleccionar_archivo_entrada(input_entry))
    input_button.grid(row=0, column=2, padx=10, pady=10)

    tk.Label(root, text="Archivo de salida:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=10, pady=10)
    output_button = tk.Button(root, text="Seleccionar archivo xlsx", command=lambda: seleccionar_archivo_salida(output_entry))
    output_button.grid(row=1, column=2, padx=10, pady=10)

    process_button = tk.Button(root, text="Generar archivo Excel", command=lambda: iniciar_proceso(input_entry, output_entry))
    process_button.grid(row=2, column=1, padx=10, pady=10)

    #TODO: Actualizar versión cada que se vea necesario
    about_me_button = tk.Button(root, text="Acerca de", command=lambda: acerca_de())
    about_me_button.grid(row=2, column=2, padx=5, pady=5)

    root.mainloop()

if __name__ == "__main__":
    iniciar_gui()