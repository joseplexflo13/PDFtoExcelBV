import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pdfplumber
import os
import subprocess
import re

def extraer_campos_especificos(pdf_path):
    texto = ""
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        texto = page.extract_text() or ""

    lines = texto.splitlines()
    joined_text = "\n".join(lines)

    datos = {
        "Market Channel Reference #": "",
        "Destination Purchase Order #": "",
        "Market": "",
        "PO Channel": "",
        "Ship To": "",
        "Ship Cancel Date": "",
        "In DC Date": ""
    }

    for line in lines:
        if "Market Channel Reference #" in line:
            match = re.search(r"Market Channel Reference #\s+([^\s]+)", line)
            if match:
                datos["Market Channel Reference #"] = match.group(1)

        if "Destination Purchase Order #" in line:
            match = re.search(r"Destination Purchase Order #\s+([^\s]+)", line)
            if match:
                datos["Destination Purchase Order #"] = match.group(1)

        if "Market:" in line:
            datos["Market"] = line.split("Market:")[-1].strip()

        if "PO Channel:" in line and "Market" in line:
            try:
                po_part = line.split("PO Channel:")[1]
                po_channel = po_part.split("Market")[0].strip()
                datos["PO Channel"] = po_channel
            except IndexError:
                datos["PO Channel"] = ""

        if "Ship Cancel Date" in line:
            datos["Ship Cancel Date"] = line.split("Ship Cancel Date")[-1].strip()

        if "In DC Date" in line:
            datos["In DC Date"] = line.split("In DC Date")[-1].strip()

    # --- Ship To extraction entre 'Ship To' y 'Agent Name'
    try:
        pattern = r"Ship To\s*(.*?)\s*Agent Name"
        match = re.search(pattern, joined_text, re.DOTALL)
        if match:
            raw_block = match.group(1)

            # Eliminar datos no deseados
            raw_block = re.sub(r'Factory.*?\n', '', raw_block, flags=re.DOTALL)
            raw_block = re.sub(r'COFACO INDUSTRIES S A C.*?\n', '', raw_block, flags=re.DOTALL)
            raw_block = re.sub(r'Jr San Andres.*?Molitalia', '', raw_block)
            raw_block = re.sub(r'Lima Lima \d{5}', '', raw_block)
            raw_block = re.sub(r'\bPE\b', '', raw_block)

            cleaned = "\n".join([line.strip() for line in raw_block.splitlines() if line.strip()])
            datos["Ship To"] = cleaned.replace('\n', ' ').strip()
    except Exception as e:
        print(f"❌ Error extrayendo Ship To: {e}")

    return datos

def extraer_datos_pdf(pdf_path):
    productos = []
    productos_set = set()
    campos_generales = extraer_campos_especificos(pdf_path)
    tallas_especiales = {"M/T", "L/T", "XL/T", "XXL/T"}

    with pdfplumber.open(pdf_path) as pdf:
        # Empieza desde la página 1 (índice 1), ya que la 0 es para campos generales
        for page_num in range(1, len(pdf.pages)):
            page = pdf.pages[page_num]
            texto = page.extract_text() or ""
            lines = texto.splitlines()

            for i in range(len(lines)):
                line = lines[i].strip()

                if all(keyword in line for keyword in ["Style", "Size", "Qty", "Unit", "Cost"]):
                    continue

                dept_match = re.match(r'^(\d{4})\s+(\d{6})', line)
                if dept_match:
                    parts = line.split()
                    dept = dept_match.group(1)
                    style_no = dept_match.group(2)
                    size = ""
                    unit_cost = 0.0
                    total_cost = 0.0
                    qty = 0

                    prepack = ""
                    prepack_type = ""
                    sku = ""
                    universal_cc = ""

                    style_desc_parts = []
                    for part in parts[6:]:
                        if re.match(r'^\d+$', part):
                            break
                        style_desc_parts.append(part)
                    style_desc = " ".join(style_desc_parts)

                    for j in range(i, i + 6):
                        if j >= len(lines): break
                        candidate = lines[j].strip()
                        size_match = re.search(
                            r'\b(XS|S|M|L|XL|XXL|XXXL|M/T|L/T|XL/T|XXL/T)\b\s+([\d.]+)\s+([\d,]+\.\d{2})',
                            candidate
                        )
                        if size_match:
                            size = size_match.group(1)
                            unit_cost = float(size_match.group(2))
                            total_cost = float(size_match.group(3).replace(",", ""))
                            qty = int(round(total_cost / unit_cost)) if unit_cost > 0 else 0
                            break

                    prepack_match = re.search(r'\b(Bulk|Pack|PrePack)\b', line)
                    if prepack_match:
                        prepack = prepack_match.group(0)
                        prepack_type = prepack

                    sku_match = re.search(r'\b\d{13}\b', line)
                    if sku_match:
                        sku = sku_match.group(0)

                    if i + 1 < len(lines):
                        universal_cc = lines[i + 1].strip()

                    if not style_desc or not size or qty == 0:
                        continue

                    clave = f"{dept}|{style_no}|{sku}|{size}|{qty}"
                    if clave in productos_set:
                        continue
                    productos_set.add(clave)

                    producto = {
                        **campos_generales,
                        "Dept": dept,
                        "Style No": style_no,
                        "PrePack": prepack,
                        "PrePack Type": prepack_type,
                        "Sku": sku,
                        "Style Description": style_desc,
                        "Qty Ordered (each)": qty,
                        "Universal CC# Color Desc": universal_cc,
                        "Size Desc": size,
                        "Unit Cost": unit_cost,
                        "Total Cost": total_cost
                    }

                    if size in tallas_especiales:
                        productos.append(producto.copy())
                    else:
                        productos.append(producto)
    return pd.DataFrame(productos)

def seleccionar_pdfs():
    archivos_pdf = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    if not archivos_pdf:
        return

    # Obtener la carpeta donde se encuentran los PDFs
    carpeta_pdf = os.path.dirname(archivos_pdf[0]) if archivos_pdf else os.path.expanduser("~")
    output_path = os.path.join(carpeta_pdf, "datos_finales.xlsx")

    all_data = []

    for pdf_path in archivos_pdf:
        try:
            df = extraer_datos_pdf(pdf_path)
            if df.empty:
                print(f"⚠️ No se encontraron datos en: {pdf_path}")
                continue
            all_data.append(df)
        except Exception as e:
            print(f"❌ Error procesando {pdf_path}: {e}")

    if not all_data:
        messagebox.showwarning("Advertencia", "No se encontraron datos válidos en los archivos seleccionados.")
        return

    # --- Antes de concatenar los DataFrames, agrega la columna 'PDF Path' ---
    for i, df in enumerate(all_data):
        df['PDF Path'] = archivos_pdf[i]

    resultado = pd.concat(all_data, ignore_index=True)

    # ---- Preparación de la hoja "Comercial" ----
    columnas_comercial_pivot = [
        "Market Channel Reference #", "Destination Purchase Order #",
        "Style No", "Style Description", "Universal CC# Color Desc",
        "Market", "PO Channel", "Ship Cancel Date", "In DC Date",
        "Ship To", "Dept", "PrePack", "PrePack Type", "Sku",
        "Unit Cost", "Total Cost"
    ]
    resultado_pivot = resultado.pivot_table(
        index=columnas_comercial_pivot,
        columns="Size Desc",
        values="Qty Ordered (each)",
        aggfunc="sum"
    ).reset_index()

    resultado_pivot.columns = [
        col if col in columnas_comercial_pivot else col for col in resultado_pivot.columns
    ]

    tallas = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "M/T", "L/T", "XL/T", "XXL/T"]
    columnas_ordenadas = columnas_comercial_pivot + tallas

    for col in columnas_ordenadas:
        if col not in resultado_pivot.columns:
            resultado_pivot[col] = ''

    resultado_comercial = resultado_pivot[columnas_ordenadas]
    # ---- Preparación de la hoja "Sistemas" ----
    columnas_base = [
        "Market Channel Reference #", "Destination Purchase Order #", "Style No",
        "Style Description", "Universal CC# Color Desc", "Market", "PO Channel",
        "Ship Cancel Date", "In DC Date", "Ship To", "Dept", "PrePack",
        "PrePack Type", "Sku", "Unit Cost"
    ]

   # Agrupar por Style No y Destination Purchase Order # y crear una serie con todos los datos para cada combinacion
    tallas = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "M/T", "L/T", "XL/T", "XXL/T"]
    resultado_sistemas = resultado.groupby(
        ["Style No", "Destination Purchase Order #", "Universal CC# Color Desc"]
    ).apply(lambda x: pd.Series({
        **{col: x[col].iloc[0] for col in columnas_base if col != "Style No" and col != "Destination Purchase Order #" and col != "Universal CC# Color Desc"},
        **{talla: x.loc[x['Size Desc'] == talla, 'Qty Ordered (each)'].sum() if talla in x['Size Desc'].values else 0 for talla in tallas},
        'Total Cost': x['Total Cost'].sum(),
        'PDF Path': x['PDF Path'].iloc[0]  # <-- Agrega esta línea
    })).reset_index()

    # Establezco el orden de la Data
    tallas = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "M/T", "L/T", "XL/T", "XXL/T"]
    columnas_sistemas_ordenadas = [
        "Market Channel Reference #",   # A
        "Destination Purchase Order #", # B
        "Style No",                     # C
        "Style Description",            # D
        "Universal CC# Color Desc",     # E
        "Market",                       # F
        "PO Channel",                   # G
        *tallas,                        # H-N
        "Total",                        # Nueva columna después de tallas
        "Ship Cancel Date",             # O
        "In DC Date",                   # P
        "Ship To",                      # Q
        "Dept",                         # R
        "PrePack",                      # S
        "PrePack Type",                 # T
        "Sku",                          # U
        "Unit Cost",                    # V
        "Total Cost"                    # W
    ]

    # Aseguramos que la data este ok y la ordenamos
    for col in columnas_sistemas_ordenadas:
        if col not in resultado_sistemas.columns:
            resultado_sistemas[col] = 0

    # Calcular la suma de las tallas para cada fila y agregar la columna 'Total'
    resultado_sistemas['Total'] = resultado_sistemas[tallas].sum(axis=1)

    # Agrega la columna 'URL PDF' al final
    resultado_sistemas['URL PDF'] = resultado_sistemas['PDF Path']
    columnas_sistemas_ordenadas.append('URL PDF')

    resultado_sistemas = resultado_sistemas[columnas_sistemas_ordenadas]

    # --- SEPARAR FILAS PARA TALLAS NORMALES Y ESPECIALES ---
    tallas_normales = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]
    tallas_especiales = ["M/T", "L/T", "XL/T", "XXL/T"]

    filas_finales = []

    for _, row in resultado_sistemas.iterrows():
        # Fila para tallas normales (todas las especiales en 0)
        if row[tallas_normales].sum() > 0:
            fila_normales = row.copy()
            for talla in tallas_especiales:
                fila_normales[talla] = 0
            fila_normales['Total'] = fila_normales[tallas_normales].sum()
            filas_finales.append(fila_normales)

        # Fila para todas las tallas especiales agrupadas (todas las normales en 0)
        if row[tallas_especiales].sum() > 0:
            fila_especiales = row.copy()
            for talla in tallas_normales:
                fila_especiales[talla] = 0
            fila_especiales['Total'] = fila_especiales[tallas_especiales].sum()
            filas_finales.append(fila_especiales)

    resultado_sistemas = pd.DataFrame(filas_finales)

    # ---- Escritura en Excel ----
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            resultado_comercial.to_excel(writer, sheet_name="Comercial", index=False, startrow=0, startcol=0)

            # Get the xlsxwriter objects from the dataframe writer object.
            workbook = writer.book
            worksheet = writer.sheets['Comercial']

            # Add a format. Light blue fill color.
            fmt = workbook.add_format({'bg_color': '#ADD8E6'})

            # Iterate over each row and apply conditional formatting.
            last_style_no = None
            for row_num, style_no in enumerate(resultado_comercial['Style No']):
                if row_num > 0 and style_no != last_style_no:
                    worksheet.set_row(row_num, None, fmt)
                last_style_no = style_no

            resultado_sistemas.to_excel(writer, sheet_name="Sistemas", index=False)
            workbook_sistemas = writer.book
            worksheet_sistemas = writer.sheets['Sistemas']

            # Add a format. Light blue fill color.
            fmt_sistemas = workbook_sistemas.add_format({'bg_color': '#ADD8E6'})

            # Iterate over each row and apply conditional formatting.
            last_style_no_sistemas = None
            last_destination_po_sistemas = None
            last_universal_cc_sistemas = None

            for row_num, (style_no_sistemas, destination_po_sistemas, universal_cc_sistemas) in enumerate(zip(resultado_sistemas['Style No'], resultado_sistemas['Destination Purchase Order #'], resultado_sistemas['Universal CC# Color Desc'])):
                if row_num > 0 and (style_no_sistemas != last_style_no_sistemas or destination_po_sistemas != last_destination_po_sistemas or universal_cc_sistemas != last_universal_cc_sistemas):
                    worksheet_sistemas.set_row(row_num, None, fmt_sistemas)

                last_style_no_sistemas = style_no_sistemas
                last_destination_po_sistemas = destination_po_sistemas
                last_universal_cc_sistemas = universal_cc_sistemas

            # Ajusta el ancho de columnas
            for idx, col in enumerate(resultado_sistemas.columns):
                max_len = max(
                    resultado_sistemas[col].astype(str).map(len).max(),
                    len(str(col))
                ) + 2
                worksheet_sistemas.set_column(idx, idx, max_len)

            # Escribe hipervínculo en la columna 'URL PDF'
            url_col_idx = resultado_sistemas.columns.get_loc('URL PDF')
            for row in range(1, len(resultado_sistemas) + 1):
                pdf_path = resultado_sistemas.iloc[row - 1]['URL PDF']
                worksheet_sistemas.write_url(row, url_col_idx, f'file:///{pdf_path}', string='Abrir PDF')

    except Exception as e:
        print(f"Error al escribir en el archivo Excel: {e}")
        messagebox.showerror("Error", f"Ocurrió un error al escribir en el archivo Excel:\n{e}")
        return  # Salir de la función si hay un error en la escritura

    messagebox.showinfo("Éxito", f"Datos exportados a:\n{output_path}")
    try:
        if os.name == 'nt':
            os.startfile(output_path)
        elif os.name == 'posix':
            subprocess.call(['open', output_path])
    except:
        pass

# Interfaz gráfica
root = tk.Tk()
root.title("Extractor PDF")
root.geometry("450x200")

label = tk.Label(root, text="Selecciona PDFs con tabla en la página 2", pady=20)
label.pack()

boton = tk.Button(root, text="Seleccionar PDFs", command=seleccionar_pdfs, height=2, width=25)
boton.pack()

root.mainloop()