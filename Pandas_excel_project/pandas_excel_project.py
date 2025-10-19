import pandas as pd

df = pd.read_excel('employee_data.xlsx')
# print("✅ Original Data:")
# print(df.head())

# Add a new column for Total Salary (Salary + Bonus)
df['Total_Salary'] = df["Salary"] + df["Bonus"]
# df.to_excel("processed_employee_data.xlsx", index=False)

# Find Average Salary per Department
avg_salary = df.groupby("Department")["Total_Salary"].mean().reset_index()
# avg_salary.to_excel("avg_salary_by_department.xlsx", index=False)

# Filter Employees with Experience > 3 Years
experienced = df[df["Experience"] > 3]
# experienced.to_excel("experienced_employees.xlsx", index=False)

summary_data = {
    "Total Employees": [len(df)],
    "Average Salary": [df["Salary"].mean()],
    "Highest Salary": [df["Salary"].max()],
    "Total Bonus Paid": [df["Bonus"].sum()],
    "Average Experience (Years)": [df["Experience"].mean()]
}
summary_df = pd.DataFrame(summary_data)

#### Write all DataFrames into ONE Excel file (multiple sheets) ####
with pd.ExcelWriter('emloyee_analysis.xlsx', engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Total_salary", index=False)
    avg_salary.to_excel(writer, sheet_name="Average_Salary", index=False)
    experienced.to_excel(writer, sheet_name="Experienced_Staff", index=False)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    
print("✅ Data written to 'employee_analysis.xlsx' with multiple sheets!")

