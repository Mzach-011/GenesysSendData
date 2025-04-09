import pyodbc
import pandas as pd
from datetime import datetime, timedelta
import getpass
import os
import msoffcrypto
import io
import win32com.client as win32



message = """
************************************************************************************************
************************************************************************************************
************************************************************************************************
This report generation was made possible through the divine intervention of our Lord and Savior,
the DWH ArchiteKt, ArchiPanda 🐼🙏. 

Without his infinite wisdom, unmatched skills in ETL, and his deep love for data, 
none of this would be possible. ArchiPanda’s work is the true backbone of the DWH universe,  
and may his pandas always be plentiful and his queries forever optimized.  
🐼🐼🐼 Bless the pandas, bless the data, and bless ArchiPanda! 🐼🐼🐼. 

⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣀⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣀⣀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⣠⣶⣿⣿⣿⣿⣷⡀⣀⣠⡤⠤⠤⠤⠤⠤⣄⣀⡀⣴⣿⣿⣿⣿⣷⣄⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⢰⣿⣿⣿⣿⣿⣿⠟⠋⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠉⠛⢿⣿⣿⣿⣿⣿⣇⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⢸⣿⣿⣿⣿⠋⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠙⢿⣿⣿⣿⡟⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠹⢿⡿⠁⠀⠀⠀⣠⣤⣤⣄⠀⠀⠀⠀⢠⣤⣤⣄⡀⠀⠀⠀⢻⡿⠟⠁⠀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢰⠃⠀⠀⢀⣾⣿⣿⣿⡟⣀⣀⣀⣀⢸⣿⣿⣿⣷⡄⠀⠀⠀⣧⡀⣀⡀⠀⠀⠀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⠀⠀⢀⣀⡀⡟⠀⠀⠀⢸⣿⣿⣿⡏⠘⢿⣿⣿⣿⠏⠙⣿⣿⣿⡇⢀⣴⣾⡿⢿⡿⢿⣶⣦⡀⠀⠀⠀⠀
⠀⠀⠀⠀⠀⣴⣾⣿⣿⣿⣿⣶⣄⠀⠀⠻⣿⠿⠃⠠⣀⣨⣏⣀⡀⠀⠻⠿⡿⠁⢸⣿⣹⡷⠿⠿⢿⣍⣿⡇⠀⠀⠀⠀
⣀⣀⣀⣀⣰⣿⣿⣿⣿⣿⣿⣿⣿⣧⣀⣀⣀⣀⣀⣀⣀⣀⣀⣀⣀⣀⣀⣀⣀⣀⣘⣿⣿⣄⣀⣀⣀⣿⣿⣇⣀⣀⣀⣤
⠀⠀⠀⠀⠘⠿⢿⣿⣿⣿⣿⣿⠿⠏⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
⠀ ⠀⠀⠀⠀⠀⠈⠉⠉⠉⠁
***********************************************************************************************
***********************************************************************************************
***********************************************************************************************
"""

print(message)



print ("Vyber kterou blbost generuju:")
print ("Zmáčkni 1 pro vítání občánků")
print ("Zmáčkni 2 pro ty co po cestě ztratili občanku nebo jí nezvládli vyfoti z obou stran")
user_choice = input("Máčk: ")





if user_choice == '1':
    procedure = "[dbo].[usp_GetNewAccounts]"
    subject = "Sestava nových klientů za předchozí dny"
    body = "Hello, \n\nAttached is the report for the new clients.\n\nBest regards. \n MZ"
    save_directory = r"H:\ACT_SKRIPTY_USEFULL\Zdraveni_Novych_Klientu_Ahoj\AutomatedGenerationNewClients"  
    excel_password = "TB_SA_NewClients_1"  # Password for NewClients
elif user_choice == '2':
    procedure = "[dbo].[usp_GetApplicationsWithoutID]"  
    subject = "Sestava chybějících ID"
    body = "Hello, \n\nAttached is the report for the missing or incomplete  IDs.\n\nBest regards. \n MZ"
    save_directory = r"H:\ACT_SKRIPTY_USEFULL\Obcanky_MM\AutomatedGenerationMissingIds"  
    excel_password = "TB_SA_IDs_1"  # Password for MissingIds
else:
    print("It's impossible to underestimate you")
    exit()


# Get user input for DB login
username = input("UID (Owner): ")
password = getpass.getpass("Heslo: ")

# Database connection details
server = "msdwh-dwh.mpu.cz"
database = "DWH_L1_OnlnCore"

# Calculate date parameter
cur_date = datetime.now().date()
used_date = cur_date - timedelta(days=1)
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
        output_file = os.path.join(save_directory, f"Report_{used_date_str}.xlsx")
        df.to_excel(output_file, index=False, engine="openpyxl")

        # Encrypt the Excel file
        encrypted_file = os.path.join(save_directory, f"Report_{used_date_str}_protected.xlsx")

        encrypted = io.BytesIO()

        # Encrypt the file with a password using msoffcrypto
        with open(output_file, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.encrypt(excel_password, encrypted)  # Use the appropriate password for the selected report

            # Save the encrypted file to a new location
            with open(encrypted_file, "wb") as ef:
                ef.write(encrypted.getvalue())  # Save the encrypted content to the file

        print(f"🔒 Encrypted Excel file saved at: {encrypted_file}")

        # Select the most recent Excel file in the save directory
        files = [f for f in os.listdir(save_directory) if f.endswith('.xlsx') and '_protected' in f]
        latest_file = max(files, key=lambda x: os.path.getmtime(os.path.join(save_directory, x)))

        # Construct the full path to the latest encrypted file
        latest_file_path = os.path.join(save_directory, latest_file)

        # Draft the email with Outlook
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0: olMailItem (new mail item)

        mail.Subject = subject
        mail.Body = body
        mail.To = "mzach@mediso.cz"  # Change as needed
        mail.Attachments.Add(latest_file_path)

        mail.Send()
        print(f"Email drafted successfully with attachment. Draft saved in Outlook.")

    else:
        print("Stored Procedure executed successfully, but no results were returned.")

except Exception as e:
    print(f"Error occurred: {str(e)}")

finally:
    if cursor:
        cursor.close()
    if conn:
        conn.close()
