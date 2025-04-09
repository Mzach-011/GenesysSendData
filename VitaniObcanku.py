import pyodbc
import pandas as pd
from datetime import datetime, timedelta
import getpass
import os
import msoffcrypto
import io
import win32com.client as win32

#***********************************************************
#Autor: mzach
#Popis: Konzolová aplikace, která slouží k odeslání reportů mailem - aktuálně obsahuje dva různé reporty (chybějící občanky a vítání občánků)
#
#Jak použít: Aktuálně skript neběží nikde samostatně jednou denně se zapne dávkovej soubor (BAT) a odešle se pomocí konzole
#
#Prerekvizita: Tvořeno na moje sturktury na mojem PC nutné upravit pro jiné + mít loglou mailovku.
#
#Changelog:
#  09.04.2025	mzach - Založena část generující json
#************************************************************/

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

#Jednoduché nastavení pro jednotlivé reporty (do této části by bylo samozřejmě vhodné přehodit i recipienty apod., jelikož ale chodí na stejné neřeším)
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


#Nastavení uživatele a hesla pro DB (předvyplněn je tam owner)
default_username = "SVC_DWH_PROD_OWNER"    
username = input(f"UID (Owner) [{default_username}]: ") or default_username
password = getpass.getpass("Heslo: ")

#Nastavení connectiony a DB aktuálně může běžet nad jednou DB samozřejmě možné přihodit do parametrů výše
server = "msdwh-dwh.mpu.cz"
database = "DWH_L1_OnlnCore"

#Bordel z D-1 asi pro srandu kralíkům spíš 
cur_date = datetime.now().date()
used_date = cur_date - timedelta(days=1)
used_date_str = used_date.strftime('%Y-%m-%d')

print(f"Executing stored procedure: EXEC {procedure} '{used_date_str}'")

try:
    #Connectiona do SQl
    conn = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}")
    cursor = conn.cursor()

    print("Connection established successfully")

    # Pust procku - musí bejt SET NOCOUN ON !!!!! 
    cursor.execute(f"SET NOCOUNT ON; EXEC {procedure} '{used_date_str}'")

    # Nahraj srajdy 
    columns = [column[0] for column in cursor.description]  # Get column names
    rows = cursor.fetchall()

    if rows:
        df = pd.DataFrame.from_records(rows, columns=columns)
        print(f"Rows returned: {len(df)}")

        # Dynamically set the report name based on procedure
        if procedure == "[dbo].[usp_GetNewAccounts]":
            report_name = "NewClients"
        elif procedure == "[dbo].[usp_GetApplicationsWithoutID]":
            report_name = "ApplicationsWithoutID"
        else:
            report_name = "UnknownReport"  # Default name if neither matches

        # Set the file path with the dynamic name
        output_file = os.path.join(save_directory, f"{report_name}_{used_date_str}.xlsx")
        encrypted_file = os.path.join(save_directory, f"{report_name}_{used_date_str}_protected.xlsx")

        # Save the unprotected Excel file
        df.to_excel(output_file, index=False, engine="openpyxl")

        # tady už tvořím ten Excel pro zaheslovaný hodnoty 
        encrypted = io.BytesIO()

        #Zahesluj
        with open(output_file, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.encrypt(excel_password, encrypted)  # heslo z parametrů nastavenejch

            # lalala
            with open(encrypted_file, "wb") as ef:
                ef.write(encrypted.getvalue())  # Napal data z prvního do tohodle
        #Existuje nějakej zakryptovanej soubor
        print(f"🔒 Encrypted Excel file saved at: {encrypted_file}")

        # najdi nejnovější excel v tý složce, kterej je protected (podle data uložení tam)
        files = [f for f in os.listdir(save_directory) if f.endswith('.xlsx') and '_protected' in f]
        latest_file = max(files, key=lambda x: os.path.getmtime(os.path.join(save_directory, x)))

        # Vytvoření cesty jen k tomu excelu
        latest_file_path = os.path.join(save_directory, latest_file)

        # Draft excel appky 
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0: olMailItem => musí bejt CreateItem0 to je novej mail

        mail.Subject = subject
        mail.Body = body

       # mail.To = "mzach@mediso.cz"
        mail.To = "JNejepsa@mediso.cz"  #Aktuálně nastaveno takhle ale jak jsem psal mohlo by bejt v parametrech nahoře 
        mail.CC = "mzach@mediso.cz;LSmolak@mediso.cz"  # CC recipients
        mail.Attachments.Add(latest_file_path)

        mail.Send()
        print(f"Odeslal jsem E-mail.Panda by byla hrdá.")

    else:
        print("Procka nevrátila žádná data..")

except Exception as e:
    print(f"Error occurred: {str(e)}")

finally:
    if cursor:
        cursor.close()
    if conn:
        conn.close()
