
if __name__ == '__main__':
    import tkinter as tk
    from tkinter import filedialog
    import pandas as pd


    def load_file_1():
        global original_file_path
        original_file_path = filedialog.askopenfilename(title="Select the original file")


    def load_file_2():
        global adjustment_file_path
        adjustment_file_path = filedialog.askopenfilename(title="Select the adjustment file")


    def process_files():
        if not original_file_path or not adjustment_file_path:
            result_label.config(text="Please select both files before processing.")
            return

        def process_original_document(file_path):
            invoice_data = pd.read_excel(file_path, skiprows=22, usecols=list(range(1, 55)), engine='openpyxl')

            # Combining the necessary columns for each field, with correct handling of merged Excel cells
            def combine_columns_by_index(data_frame, start_idx, end_idx):
                # Combining data across specified column indices, dropping NaNs and joining non-empty strings
                return data_frame.iloc[:, start_idx:end_idx + 1].apply(
                    lambda row: ' '.join(row.dropna().astype(str).values), axis=1
                )

            # Iterate over each row to find 'Итого' and break the loop once found
            print("row.values")
            for index, row in invoice_data.iterrows():
                print(row.values)
                if 'Итого:' in row.values:
                    invoice_data = invoice_data[:index]  # Keep rows only before 'Итого'
                    break

            # Specify the exact columns for each data field based on your Excel layout
            invoice_data['Item Number'] = range(1, len(invoice_data) + 1)  # Sequential numbering for items
            invoice_data['Vendor Code'] = combine_columns_by_index(invoice_data, 2, 7)  # Columns D to I for "Артикул"
            invoice_data['Description'] = combine_columns_by_index(invoice_data, 8,
                                                                   38)  # Columns J to AN for description
            invoice_data['Quantity'] = combine_columns_by_index(invoice_data, 39, 44)  # Columns AO to AT for quantity
            invoice_data['Unit'] = combine_columns_by_index(invoice_data, 45, 47)  # Columns AU to AW for unit
            invoice_data['Price'] = combine_columns_by_index(invoice_data, 48, 52)  # Columns AX to BB for price
            invoice_data['Total'] = combine_columns_by_index(invoice_data, 53, 53)  # Columns BC to BI for total

            # Select only the relevant columns in the desired order for the final DataFrame
            data = invoice_data[
                ['Item Number', 'Vendor Code', 'Description', 'Quantity', 'Unit', 'Price', 'Total']]
            return data

        def process_processed_document(file_path):
            # This function processes the already processed (final output) document
            data = pd.read_excel(file_path)
            # Logic for subtracting rows or different column handling
            # For example, handling different columns:
            print(data)

            data.rename(columns={'Товары (работы, услуги)': 'Description'}, inplace=True)
            data.rename(columns={'Артикул': 'Vendor Code' }, inplace=True)
            data.rename(columns={'№': 'Item Number'}, inplace=True)
            data.rename(columns={'Кол-во': 'Quantity'}, inplace=True)
            data.rename(columns={'Ед.': 'Unit'}, inplace=True)
            data.rename(columns={'Цена': 'Price'}, inplace=True)
            print(data)
            return data

        def determine_and_process(file_path):
            # Determine if the file is an original or a processed final output based on its content or filename
            if "_processed" in file_path:
                print("Processing processed document.")
                return process_processed_document(file_path)
            else:
                print("Processing original document.")
                return process_original_document(file_path)

        # Example usage:
        # file_path_original = 'path_to_your_original_file.xlsx'
        # file_path_final = 'path_to_your_final_output.xlsx'
        final_data = determine_and_process(original_file_path)

        # Read the Excel file, skipping the first 22 rows that are headers and other non-data content


        # Write the DataFrame to a CSV file
        pd.set_option('display.max_columns', None)
        print("final_data")
        print(final_data)
        final_data.to_csv('output_invoice_data_final.csv', index=False)

        #--------------------------------------------------------------------------------------------------------------

        # Load the main DataFrame
        main_df = pd.read_csv('output_invoice_data_final.csv')

        # Load the adjustment document
        adjustment_df_first = pd.read_excel(adjustment_file_path, usecols="A:AI", engine='xlrd')

        # Apply a function to each row that checks for the keyword and prints the row
        def check_keyword_and_print(row):
            row_str = row.astype(str)
            print("row_str")
            print(row_str)  # Print the row that contains the keyword
            contains_keyword = row_str.str.contains('Универсальный передаточный документ', case=False).any()  # case insensitive
            if contains_keyword:
                print("row_str yes")
                print(row_str)  # Print the row that contains the keyword
            print("contains_keyword: ")
            print(contains_keyword)
            return contains_keyword


        # Identify rows containing the keyword that causes breaks
        skip_rows = adjustment_df_first.apply(check_keyword_and_print, axis=1, result_type="reduce")
        print("skip_rows row")
        print(skip_rows)
        skip_row_indices = skip_rows[skip_rows].index.tolist()
        skip_row_indices_correct = []
        for e in skip_row_indices:
            skip_row_indices_correct.append(e + 1)

        print("skip_row_indices_correct")
        print(skip_row_indices_correct)
        print(skip_row_indices_correct.__sizeof__())

        expanded_skip_rows = []
        for x in skip_row_indices_correct:
            expanded_skip_rows.extend(range(x, x + 5))  # x + 5 because range is exclusive of the stop value

        expanded_skip_rows.extend(range(0, 19))
        print("expanded_skip_rows")
        print(expanded_skip_rows)
        adjustment_df_second = pd.read_excel(adjustment_file_path, skiprows=expanded_skip_rows, usecols="D:AI", engine='xlrd')

        # Continue with data processing
        print("adjustment_df.head()")
        print(adjustment_df_second.head())

        # Iterate over each row to find 'Итого' and break the loop once found
        for index, row in adjustment_df_second.iterrows():
            print(row.values)
            if 'Всего к оплате' in row.values:
                adjustment_df_second = adjustment_df_second[:index]  # Keep rows only before 'Итого'
                break




        # Define a function to combine specified columns into a single string (for Vendor Code and Description)
        def combine_columns(df, start_col, end_col):
            return df.iloc[:, start_col:end_col + 1].apply(
                lambda row: ' '.join(row.dropna().astype(str).values), axis=1
            )


        # If headers are incorrectly set as data (like 'tpnd-50-to'), you might need to manually set headers
        adjustment_df_second.columns = ['Vendor Code', 'Other Column 1', 'Other Column 2', 'Other Column 3', 'Other Column 4', 'Other Column 5', 'Other Column 6', 'Other Column 7', 'Other Column 8', 'Other Column 9', 'Other Column 10', 'Other Column 11', 'Other Column 12', 'Other Column 13', 'Other Column 14', 'Other Column 15', 'Other Column 16', 'Other Column 17', 'Other Column 18', 'Other Column 19', 'Other Column 20', 'Other Column 21', 'Other Column 22', 'Other Column 23', 'Other Column 24', 'Other Column 25', 'Other Column 26', 'Other Column 27', 'Other Column 28', 'Other Column 29', 'Other Column 30',
                        'Quantity']  # Replace with actual column headers or rename as needed

        print("adjustment_df_second")
        print(adjustment_df_second)
        adjustment_df_second = adjustment_df_second[['Vendor Code', 'Other Column 7', 'Other Column 25', 'Other Column 28']]
        adjustment_df_second.rename(columns={'Other Column 7': 'Description'}, inplace=True)
        adjustment_df_second.rename(columns={'Other Column 25': 'Unit'}, inplace=True)
        adjustment_df_second.rename(columns={'Other Column 28': 'Quantity'}, inplace=True)
        print("adjustment_df_second")
        print(adjustment_df_second)
        adjustment_df_second.to_csv('adjustment_df.csv', index=False)

        merged_df = pd.merge(main_df, adjustment_df_second[['Vendor Code', 'Unit']], on=['Vendor Code', 'Unit'], how='left',
                             indicator=True)

        # Set 'Quantity' to 0 where a match was found
        merged_df['Quantity'] = merged_df.apply(lambda row: 0 if row['_merge'] == 'both' else row['Quantity'], axis=1)

        # Drop the '_merge' column used for identifying matched rows
        merged_df.drop(columns=['_merge'], inplace=True)

        # Save the adjusted DataFrame to a new CSV file
        merged_df.to_csv(original_file_path.replace(".xlsx", "", 1) + '_with_zeros.csv', index=False)
        merged_df.to_excel(original_file_path.replace(".xlsx", "", 1) + '_with_zeros.xlsx', index=False)

        # Create a separate DataFrame where rows with Quantity 0 are removed
        non_zero_df = merged_df[merged_df['Quantity'] != 0]

        # Save this new DataFrame to a separate Excel file
        non_zero_df = non_zero_df[['Item Number', 'Vendor Code', 'Description', 'Quantity', 'Unit', 'Price']]
        non_zero_df.rename(columns={'Description': 'Товары (работы, услуги)'}, inplace=True)
        non_zero_df.rename(columns={'Vendor Code': 'Артикул'}, inplace=True)
        non_zero_df.rename(columns={'Item Number': '№'}, inplace=True)
        non_zero_df.rename(columns={'Quantity': 'Кол-во'}, inplace=True)
        non_zero_df.rename(columns={'Unit': 'Ед.'}, inplace=True)
        non_zero_df.rename(columns={'Price': 'Цена'}, inplace=True)
        # adjustment_df.rename(columns={'Total': 'Сумма'}, inplace=True)

        non_zero_df.to_excel(original_file_path.replace(".xlsx", "", 1) + '_processed.xlsx', index=False)
        # result_label.config(text="Files processed and saved successfully.")
        root.title("Files processed and saved successfully.")
        print("done.")


    # Setup the basic UI
    root = tk.Tk()
    root.title("Data Adjustment Tool")

    load_button_1 = tk.Button(root, text="Загрузить Счет", command=load_file_1)
    load_button_1.pack(pady=10)

    load_button_2 = tk.Button(root, text="Загрузить Передаточный документ", command=load_file_2)
    load_button_2.pack(pady=10)

    process_button = tk.Button(root, text="Начать процесс", command=process_files)
    process_button.pack(pady=20)

    result_label = tk.Label(root, text="")
    result_label.pack(pady=20)

    root.mainloop()

