
column_names = df.iloc[0]  # Get the first row as column names
df = df[1:].reset_index(drop=True)  # Exclude the first row from the DataFrame
df.columns = column_names 