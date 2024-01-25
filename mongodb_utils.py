# mongodb_utils.py
from pymongo import MongoClient

def connect_to_mongodb():
    # Replace 'your_connection_string' with the actual MongoDB connection string
    connection_string = 'mongodb+srv://<username>:<password>@cluster0.ahnhw.mongodb.net/'
    client = MongoClient(connection_string)
    db = client['simple_excel_db']  # Replace 'simple_excel_db' with the desired database name
    collection = db['simple_excel_collection']  # Replace 'simple_excel_collection' with the desired collection name
    return collection

def save_data_to_mongodb(collection, data):
    collection.update_one({}, {"$set": data}, upsert=True)

def load_data_from_mongodb(collection):
    return collection.find_one({}, {"_id": 0})

def cloud_save_to_mongodb(data, formulas, calculated_values):
    try:
        collection = connect_to_mongodb()

        # Save data to MongoDB
        data_to_save = {
            "data": data,
            "formulas": formulas,
            "calculated_values": calculated_values
        }
        save_data_to_mongodb(collection, data_to_save)

        return True, "Data saved to MongoDB."
    except ConnectionError as ce:
        return False, f"Unable to connect to MongoDB: {ce}"
    except RuntimeError as re:
        return False, f"Error saving data to MongoDB: {re}"

def cloud_load_from_mongodb():
    try:
        collection = connect_to_mongodb()

        # Load data from MongoDB
        loaded_data = load_data_from_mongodb(collection)

        if loaded_data:
            return True, loaded_data
        else:
            return False, "No data found in MongoDB."
    except ConnectionError as ce:
        return False, f"Unable to connect to MongoDB: {ce}"
    except RuntimeError as re:
        return False, f"Error loading data from MongoDB: {re}"
