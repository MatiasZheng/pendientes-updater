import pandas as pd
import os
import glob

# Step 1: Search for the Excel file starting with "TempFiles_"
temp_files = glob.glob("TempFiles_*.xlsx")

if not temp_files:
    raise FileNotFoundError("No TempFiles_ Excel files found.")

# Step 2: Read data from the first found file
temp_file_path = temp_files[0]
df = pd.read_excel(temp_file_path)

# Step 3: Create a new Excel file and a new sheet
output_file_path = "Pendientes.xlsx"
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Relatorio Suportes', index=False)

    # Step 4: Adding calculated columns with Excel formulas
    workbook  = writer.book
    worksheet = writer.sheets['Relatorio Suportes']
    
    # Adding calculated formulas
    # Assuming the data starts at row 2 for Pandas (0-indexed + header)
    row_count = df.shape[0] + 1  # +1 for header row
    columns = {
        'Plazo': f'=IF(A2<>"", ...)',  # Add your specific formula here
        'Creado': f'=IF(A2<>"", ...)',
        'Demora': f'=IF(A2<>"", ...)',
        'Antigüedad': f'=IF(A2<>"", ...)',
        'Finalizado': f'=IF(A2<>"", ...)',
        'Asignado': f'=IF(A2<>"", ...)',
        'Agendado': f'=IF(A2<>"", ...)',
    }
    
    for i, (col_name, formula) in enumerate(columns.items(), start=len(df.columns)):
        worksheet.write(0, i, col_name)  # Write header for calculated column
        worksheet.write_formula(1, i, formula)  # Write formula for calculated column
    
    # Adjust column width for better readability
    for i in range(len(columns)):
        worksheet.set_column(i, i, 20)

print("Script executed successfully, Pendientes.xlsx created.")
