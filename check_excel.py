import pandas as pd

file_path = "input/studentsdata.xlsx"

df = pd.read_excel(file_path)

print("Excel loaded successfully.")
print()
print("Columns:")
print(df.columns.tolist())
print()
print("Number of students:", len(df))
print()
print("Student data:")
print(df.to_string(index=False))