from pymongo import MongoClient
import openpyxl
import json
from bson import ObjectId
from datetime import datetime

# MongoDB connection details
mongo_url = "mongodb+srv://Manas:ManasZongo2023@zongovitasa.xajrbmc.mongodb.net/ZVMASTER?retryWrites=true&w=majority"
database_name = "ZV_DEMO"

# Load the Excel workbook
excel_path = "DemoData.xlsx" 
workbook = openpyxl.load_workbook(excel_path)
sheet = workbook.active

try:
    # Connect to MongoDB
    client = MongoClient(mongo_url)
    db = client[database_name]
    print("Connected to MongoDB!")

    for row in sheet.iter_rows(min_row=2, values_only=True):
        operation = row[3] 
        if operation == "Insert":
            collection_name = row[2]  
            body = json.loads(row[4])

            if '_id' in body and isinstance(body['_id'], dict) and '$oid' in body['_id']:
                body['_id'] = body['_id']['$oid']

            if '_id' in body:
                del body['_id']

            current_date = datetime.now()

            for date_field in ['maint_sch_date', 'ins_sch_date']:
                if date_field in body:
                    body[date_field] = current_date

            collection = db[collection_name]
            collection.insert_one(body)
            print(body)
            print(f"Inserted into {collection_name}")

finally:
    if 'client' in locals() and client is not None:
        client.close()
        print("MongoDB connection closed.")
