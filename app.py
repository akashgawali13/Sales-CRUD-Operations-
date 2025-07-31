import streamlit as st
import pandas as pd
import os

# Path to Excel file
DATA_FILE = r"C:\Users\akash\Downloads\new_data.xlsx"

# Load data from Excel
def load_data():
    if os.path.exists(DATA_FILE):
        try:
            return pd.read_excel(DATA_FILE, engine='openpyxl')
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            return pd.DataFrame(columns=['OrderID', 'Product', 'Category', 'Quantity', 'Price', 'Date'])
    else:
        return pd.DataFrame(columns=['OrderID', 'Product', 'Category', 'Quantity', 'Price', 'Date'])

# Save data to Excel
def save_data(df):
    try:
        df.to_excel(DATA_FILE, index=False, engine='openpyxl')
        st.success("Data saved successfully.")
    except Exception as e:
        st.error(f"Error saving file: {e}")

# Create Record GUI
def create_record():
    st.subheader("Create New Record")
    df = load_data()

    order_id = st.number_input("OrderID", min_value=1, step=1)
    if order_id in df['OrderID'].values:
        st.warning("OrderID already exists!")
        return

    product = st.text_input("Product")
    category = st.text_input("Category")
    quantity = st.number_input("Quantity", min_value=0, step=1)
    price = st.number_input("Price", min_value=0.0, step=0.01)
    date = st.date_input("Date")

    if st.button("Add Record"):
        new_record = {
            "OrderID": order_id,
            "Product": product,
            "Category": category,
            "Quantity": quantity,
            "Price": price,
            "Date": date
        }

        df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
        save_data(df)

# Read/View Records
def read_records():
    st.subheader("Read Records")
    df = load_data()

    view_option = st.radio("Choose View Mode", ("View All", "Filter by OrderID"))

    if view_option == "View All":
        st.dataframe(df)
    else:
        oid = st.number_input("Enter OrderID", min_value=1, step=1)
        record = df[df['OrderID'] == oid]
        if record.empty:
            st.warning("No record found.")
        else:
            st.dataframe(record)

# Update Record GUI
def update_record():
    st.subheader("Update Record")
    df = load_data()

    oid = st.number_input("Enter OrderID to update", min_value=1, step=1)
    if oid not in df['OrderID'].values:
        st.warning("OrderID not found.")
        return

    index = df[df['OrderID'] == oid].index[0]

    st.write("Leave field blank to keep current value.")

    product = st.text_input("Product", value=df.at[index, 'Product'])
    category = st.text_input("Category", value=df.at[index, 'Category'])
    quantity = st.text_input("Quantity", value=str(df.at[index, 'Quantity']))
    price = st.text_input("Price", value=str(df.at[index, 'Price']))
    date = st.text_input("Date", value=str(df.at[index, 'Date']))

    if st.button("Update Record"):
        try:
            df.at[index, 'Product'] = product or df.at[index, 'Product']
            df.at[index, 'Category'] = category or df.at[index, 'Category']
            df.at[index, 'Quantity'] = int(quantity) if quantity else df.at[index, 'Quantity']
            df.at[index, 'Price'] = float(price) if price else df.at[index, 'Price']
            df.at[index, 'Date'] = pd.to_datetime(date) if date else df.at[index, 'Date']
            save_data(df)
        except Exception as e:
            st.error(f"Error updating record: {e}")

# Delete Record GUI
def delete_record():
    st.subheader("Delete Record")
    df = load_data()

    oid = st.number_input("Enter OrderID to delete", min_value=1, step=1)
    if st.button("Delete"):
        if oid in df['OrderID'].values:
            df = df[df['OrderID'] != oid]
            save_data(df)
            st.success("Record deleted.")
        else:
            st.warning("OrderID not found.")

# Main GUI
def main():
    st.title("📊 Sales Records CRUD App")
    menu = ["Create", "Read", "Update", "Delete"]
    choice = st.sidebar.radio("Select Operation", menu)

    if choice == "Create":
        create_record()
    elif choice == "Read":
        read_records()
    elif choice == "Update":
        update_record()
    elif choice == "Delete":
        delete_record()

if __name__ == "__main__":
    main()
