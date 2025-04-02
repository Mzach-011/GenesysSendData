import pyodbc
import pandas as pd
from datetime import datetime, timedelta
import getpass
import os
import msoffcrypto
import io

# Set save directory
save_directory = r"H:\ACT_SKRIPTY_USEFULL\Zdraveni_Novych_Klientu_Ahoj\AutomatedGenerationNewClients"
os.makedirs(save_directory, exist_ok=True)  # Ensure the directory exists

# Get user input
username = input("Enter username: ")
password = getpass.getpass("Enter password: ")

# Database connection details
server = "msdwh-dwh.mpu.cz"
database = "DWH_L1_OnlnCore"
procedure = "[dbo].[usp_GetNewAccounts]"

# Calculate date parameter
cur_date = datetime.now().date()
used_date = cur_date - timedelta(days=2)
used_date_str = used_date.strftime('%Y-%m-%d')

print(f"Executing stored procedure: EXEC {procedure} '{used_date_str}'")

try:
    # Connect to SQL Server
    conn = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}")
    cursor = conn.cursor()

    print("Connection established successfully")

    # Execute the stored procedure
    cursor.execute(f"SET NOCOUNT ON; EXEC {procedure} '{used_date_str}'")

    # Fetch data and store in DataFrame
    columns = [column[0] for column in cursor.description]  # Get column names
    rows = cursor.fetchall()

    if rows:
        df = pd.DataFrame.from_records(rows, columns=columns)
        print(f"Rows returned: {len(df)}")

        # Save to Excel (unencrypted first)
        output_file = os.path.join(save_directory, f"NewClients_{used_date_str}.xlsx")
        df.to_excel(output_file, index=False, engine="openpyxl")

        # Encrypt the Excel file
        encrypted_file = os.path.join(save_directory, f"NewClients_{used_date_str}_protected.xlsx")

        encrypted = io.BytesIO()
        
        # Encrypt the file with a password using msoffcrypto
        with open(output_file, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.encrypt("TB_SA_NewClients_1", encrypted)  # Set password

            # Save the encrypted file to a new location
            with open(encrypted_file, "wb") as ef:
                ef.write(encrypted.getvalue())  # Save the encrypted content to the file

        print(f"🔒 Encrypted Excel file saved at: {encrypted_file}")

    else:
        print("Stored Procedure executed successfully, but no results were returned.")

except Exception as e:
    print(f"Error occurred: {str(e)}")

finally:
    if cursor:
        cursor.close()
    if conn:
        conn.close()
