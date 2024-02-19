import pandas as pd

# Define data
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [30, 35, 40],
    'Department': ['HR', 'Finance', 'IT'],
    'Salary': [50000, 60000, 70000]
}

# Create DataFrame
df = pd.DataFrame(data)

column_names = df.iloc[0]  # Get the first row as column names
df = df[1:].reset_index(drop=True)  # Exclude the first row from the DataFrame
df.columns = column_names 
# Display the DataFrame
print(df)
print(df[float('50000')].sum())
# print(df.loc[1, 'Age'])
# # import pandas as pd

# # Create a list of strings
# data = ['apple', 'banana', 'orange', 'kiwi']

# # Convert the list into a DataFrame
# df = pd.DataFrame(data, columns=['Fruits'])

# # Display the DataFrame
# print(df)
