import pandas as pd
import os

#path to excel file with impact categories
excel_file = r"C:\Users\...\EF3.1_EN15804_impact_categories_initial.xlsx" #please change path

#Makes a dictionary of dataframes (one df per excel sheet)
excel_to_read = pd.read_excel(excel_file, sheet_name=None, header=2, usecols=['Name', 'Category','Amount', 'Unit','Uncertainty']) #header=3 starts reading the data from line 4 on and takes the column name from line 3
print(excel_to_read)

#makes one dataframe out of excel_to_read (which was a dictionary containing multiple dataframes) containing all biosphere flows of all impact categories in one dataframe
combined_df = pd.concat(excel_to_read.values(), ignore_index=True)

# Function to add another column "One_name" which contains the content of column "Name" and "Category"
#def concatenate_columns(row):
 #   return f"{row['Name']} {row['Category']}"

# Calls the function concatenate_columns
#combined_df['One_Name'] = combined_df.apply(concatenate_columns, axis=1)
#combined_df['Duplicates_Count'] = combined_df.groupby('One_Name')['One_Name'].transform('count')
combined_df['Count'] = 1
combined_df = combined_df.groupby(['Name', 'Category'])['Count'].count().reset_index()

# Path to new excel document
path = r"C:\Users\schindler\Desktop\Projects\Skripts\OBDinBW2"
combined_df.to_excel(excel_writer= os.path.join(path, "biosphere_flows_combined_one_list.xlsx"))

print(combined_df)

# List to save rows to delete
#rows_to_delete = []

# Mark line as "to be deleted" when there is a duplicate in the column "one_Name"
#for index, row in combined_df.iterrows():
 #   if combined_df['One_Name'].duplicated(keep=False)[index]:
  #      rows_to_delete.append(index)

# Delete marked line in dataframe and save it as new dataframe unique_df
#unique_df = combined_df.drop(rows_to_delete)

#print(unique_df)

# Create new excel document
#unique_df.to_excel(excel_writer= os.path.join(path, "unique_biosphere_flows_one_list.xlsx"))

# Define folder where new excel files are saved
folder = r"C:\Users\schindler\Desktop\Projects\Skripts\OBDinBW2\Marked"

# Iteration over every sheet in the excel_to_read
for sheet_name, df_excel in excel_to_read.items():
    # Merge der DataFrames unter Verwendung eines linken Joins
    merged_df_duplicates = pd.merge(df_excel, combined_df[['Name', 'Category', 'Count']],
                         how='inner', on=['Name', 'Category'])

    # Aktualisiere das ursprüngliche Arbeitsblatt im Dictionary mit der neuen Spalte "Duplicates_count"
    excel_to_read[sheet_name] = merged_df_duplicates

    # Saves the marked excel sheet in a new excel document in the previously defined folder
    with pd.ExcelWriter(os.path.join(folder, f"{sheet_name}_marked.xlsx")) as writer:
        merged_df_duplicates.to_excel(writer, index=False)

# Iterates over every sheet in the excel_to_read
#for sheet_name, df_excel in excel_to_read.items():
    # Compares the columns "Name" & "Category" with "unique_df"
 #   merged_df = pd.merge(df_excel[['Name', 'Category']], unique_df[['Name', 'Category']], how='inner',
  #                       indicator=True, on=['Name', 'Category'])

    # Marks the line in the excel_to_read, which is also in the "unique_df"
   # df_original['Marked'] = merged_df['_merge'] == 'both'

    # Saves the marked excel sheet in a new excel document in the previously defined folder
 #   with pd.ExcelWriter(os.path.join(folder, f"{sheet_name}_marked.xlsx")) as writer:
  #      df_excel.to_excel(writer, index=False)

# Prints the marked dataframes
for sheet_name, df in excel_to_read.items():
    print(f"Sheet '{sheet_name}':\n{df}\n")
