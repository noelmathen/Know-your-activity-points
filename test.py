import pandas as pd
import numpy as np
# # Define data
# data1 = {
#     'Name': ['Alice', 'Bob', 'Charlie'],
#     'Age': [30, 35, 40],
#     'Department': ['HR', 'Finance', 'IT'],
#     'Salary': [70000, 50000, 60000]
# }
# df1 = pd.DataFrame(data1)

# data2 = {
#     'Name': ['Alice', 'Bob', 'Charlie', 'David'],
#     'Age': [30, 35, 40, 25],
#     'Department': ['HR', 'Finance', 'IT', 'Marketing'],
#     'Salary': [70000, 50000, 60000, 55000]
# }
# df2 = pd.DataFrame(data2)

# data3 = {
#     'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Emma'],
#     'Age': [30, 35, 40, 25, 28],
#     'Department': ['HR', 'Finance', 'IT', 'Marketing', 'Sales'],
# }
# df3 = pd.DataFrame(data3)

# list_of_df = [df1, df2, df3]


# print("\nDataFrame 1:")
# print(df1)
# print("\nDataFrame 2:")
# print(df2)
# print("\nDataFrame 3:")
# print(df3)


# excel_file = "test_excel.xlsx"

# # Create a Pandas Excel writer object
# with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
#     # Write each DataFrame to a different range of rows
#     start_row = 0
#     for df in list_of_df:
#         df.to_excel(writer, startrow=start_row, index=True)
#         start_row += len(df) + 2  # Add 2 extra rows for space

# print(f"All DataFrames written to '{excel_file}' in a single sheet with space between them successfully!")

# df.index += 1
# df.to_excel('test_excel.xlsx', index=True, sheet_name='sheet1', startcol=1)




# import pandas as pd

# # Create a string variable
# my_string = "Hello, world!"

# # Create a DataFrame with the string as a single column
# df = pd.DataFrame({'Text': [my_string]})

# # Specify the Excel file name
# excel_file = 'output.xlsx'

# # Write the DataFrame to an Excel file
# df.to_excel(excel_file, index=False, header=False, startrow=12)

# print("String variable written to Excel file successfully!")



import pandas as pd
df = pd.DataFrame({'A': [1, 2, np.nan, np.nan, np.nan], 'B': [3, 4, np.nan, 5, np.nan]})
print(df)
df = df.dropna(how='all')
print(df)




# column_names = df.iloc[0]  # Get the first row as column names
# df = df[1:].reset_index(drop=True)  # Exclude the first row from the DataFrame
# df.columns = column_names 
# # Display the DataFrame
# print(df)
# print(df[float('50000')].sum())
# print(df.loc[1, 'Age'])
# # import pandas as pd

# # Create a list of strings
# data = ['apple', 'banana', 'orange', 'kiwi']

# # Convert the list into a DataFrame
# df = pd.DataFrame(data, columns=['Fruits'])

# # Display the DataFrame
# print(df)
