import pandas as pd

# Read the Excel file with the correct header row
excel_file = r'C:\Users\pserr\Desktop\Dissertation\PsetDefinition.xlsx'  # Replace with the correct path to your Excel file
df = pd.read_excel(excel_file, header=0)

# Print the columns of the DataFrame for debugging
print("Columns in the DataFrame:", df.columns)

# Rename the columns to match the expected names
df.columns = ['PropertySet', 'Instance/Type', 'ElementList', 'PropertyName', 'DataType', 'RevitParameterName']

# Strip any leading/trailing spaces from column names
df.columns = df.columns.str.strip()

# Function to generate the PropertySet definition
def generate_property_set(df):
    property_set_definitions = []
    current_property_set = None

    for index, row in df.iterrows():
        if row['PropertySet'] != current_property_set:
            if current_property_set is not None:
                property_set_definitions.append('\n')
            current_property_set = row['PropertySet']
            property_set_definitions.append(
                "PropertySet:\t" + row['PropertySet'] + "\t" + row['Instance/Type'] + "\t" + row['ElementList'] + "\n"
            )
        if pd.isna(row['RevitParameterName']):
            property_set_definitions.append(
                f"\t{row['PropertyName']}\t{row['DataType']}\n"
            )
        else:
            property_set_definitions.append(
                f"\t{row['PropertyName']}\t{row['DataType']}\t{row['RevitParameterName']}\n"
            )

    return ''.join(property_set_definitions)

# Generate the property set definitions
property_set_definitions = generate_property_set(df)

# Save to a text file
with open('user_defined_property_sets.txt', 'w') as f:
    f.write(property_set_definitions)

print("Property set definitions successfully generated and saved to 'user_defined_property_sets.txt'")
