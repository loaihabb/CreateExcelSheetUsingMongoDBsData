import pymongo
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

# MongoDB'ye bağlan
client = pymongo.MongoClient("< YOUR_MONGO_DATABASE >")
db = client["< CLIENT >"]
collection = db["< COLLECTION >"]


def generate_excel(month):
    # Retrieve data from MongoDB
    appointments = list(collection.find({"datetwo": {"$regex": "^2024-05"}}))
    
    # Convert data to DataFrame
    df = pd.DataFrame(appointments)
    
    # Add a new column: "Rent_Deposit"
    df["Rent_Deposit"] = df["rent"] - df["deposit"]

    # change the name of columns
    df = df.rename(columns={"dateone": "Data Entry Date"})
    df = df.rename(columns={"datetwo": "Rental Date"})
    df = df.rename(columns={"time": "Check-in Time"})
    df = df.rename(columns={"timetwo": "Check-out Time"})
    df = df.rename(columns={"number": "Phone Number "})

    # Delete column "__v"
    df = df.drop("__v", axis=1)


    # Calculate the sum of Rent_Deposit column and write it into Excel
    total_rent_deposit = df["Rent_Deposit"].sum()
    df_total = pd.DataFrame({"Total_Rent_Deposit": [total_rent_deposit]})
    df = pd.concat([df, df_total], axis=1)

    # Generate Excel file name with month and year
    excel_file = f"appointments_{month}.xlsx"
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
    df.to_excel(writer, index=False)

    # Get Excel file object
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Expand all columns
    for i, col in enumerate(df.columns):
        column_len = max(df[col].astype(str).str.len().max(), len(col)) + 2  # ekstra boşluk
        worksheet.set_column(i, i, column_len)

    # Save Excel file
    writer._save()
    print(f"{excel_file} başarıyla oluşturuldu.")

    #* print(appointments)
# Create Excel files for each month
current_month = datetime.now().strftime("%B_%Y")
generate_excel(current_month)