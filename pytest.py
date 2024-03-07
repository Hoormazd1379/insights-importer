import pandas as pd

file = './Audience.csv'
with open(file, 'r', encoding='UTF-16') as f:
    lines = f.readlines()

pre_lines = []
new_lines = []

for line in lines:
    newline = ""
    qopen = False
    for char in line:
        if char == '"':
            qopen = not qopen
        elif qopen and not char == ',':
            newline += char
        elif not qopen:
            newline += char
    newline = newline.replace('"', '')
    pre_lines.append(newline)

i = 0
while i < len(pre_lines):
    line = pre_lines[i].strip()  # remove leading and trailing whitespaces
    if line:  # if the line is not empty
        if ',' not in line and not line.isdigit():  # if the line does not have a comma
            if i + 1 < len(pre_lines):  # if there is a next line
                new_lines.append("*" + line + pre_lines[i + 1])  # append the current line and the next line
                i += 1  # skip the next line
        else:
            new_lines.append(line)
            new_lines.append('\n')  # add a newline character
    i += 1

with open(file, 'w', encoding='UTF-16') as f:
    f.writelines(new_lines)




def read_multiple_tables(file):
    with open(file, 'r', encoding='UTF-16') as f:
        lines = f.readlines()

    tables = []
    table = []
    for line in lines[1:]:  # skip the first line ("sep=,")
        if '*' in line:  # this is a header line
            if table:  # if there is already a table, add it to the list of tables
                tables.append(pd.DataFrame(table[1:], columns=table[0]))
            table = [line.strip().split(',')]  # start a new table
        else:
            table.append(line.strip().split(','))  # add a row to the current table
    if table:  # add the last table to the list of tables
        tables.append(pd.DataFrame(table[1:], columns=table[0]))

    return tables

tables = read_multiple_tables(file)

for table in tables:
    print(table)

totalFollowersIG = 0
totalFollowersFB = 0

topCityIG = "N/A"
topCityFB = "N/A"

topCountryIG = "N/A"
topCountryFB = "N/A"

topAgeIG = "N/A"
topAgeFB = "N/A"

for table in tables:
    if "*Facebook followersFB_PAGEFOLLOWUNIQUE_USERS" in table.columns.values[0]:
        totalFollowersFB = table["*Facebook followersFB_PAGEFOLLOWUNIQUE_USERS"][0]
        print(totalFollowersFB)
    
    if "*Instagram followersIG_ACCOUNTFOLLOWUNIQUE_USERS" in table.columns.values[0]:
        totalFollowersIG = table["*Instagram followersIG_ACCOUNTFOLLOWUNIQUE_USERS"][0]
        print(totalFollowersIG)
    
    if "*Facebook followers by gender and age" in table.columns.values[0]:
        if not table.empty:
            table['Women'] = table['Women'].str.rstrip('%').astype('float')
            table['Men'] = table['Men'].str.rstrip('%').astype('float')

            max_women = table['Women'].idxmax()
            max_men = table['Men'].idxmax()

            if table.loc[max_women, 'Women'] > table.loc[max_men, 'Men']:
                topAgeFB = f"Women, {table.loc[max_women, '*Facebook followers by gender and ageAge']}, with {table.loc[max_women, 'Women']}%"
            else:
                topAgeFB = f"Men, {table.loc[max_men, '*Facebook followers by gender and ageAge']}, with {table.loc[max_men, 'Men']}%"
        print(topAgeFB)
    
    if "*Instagram followers by gender and age" in table.columns.values[0]:
        if not table.empty:
            table['Women'] = table['Women'].str.rstrip('%').astype('float')
            table['Men'] = table['Men'].str.rstrip('%').astype('float')

            max_women = table['Women'].idxmax()
            max_men = table['Men'].idxmax()

            if table.loc[max_women, 'Women'] > table.loc[max_men, 'Men']:
                topAgeIG = f"Women, {table.loc[max_women, '*Instagram followers by gender and ageAge']}, with {table.loc[max_women, 'Women']}%"
            else:
                topAgeIG = f"Men, {table.loc[max_men, '*Instagram followers by gender and ageAge']}, with {table.loc[max_men, 'Men']}%"
        print(topAgeIG)

    if "*Facebook followers by top countries" in table.columns.values[0]:
        if not table.empty:
            table['Value'] = table['Value'].str.rstrip('%').astype('float')
            max_country = table['Value'].idxmax()
            topCountryFB = f"{table.loc[max_country, '*Facebook followers by top countriesTop countries']}, with {table.loc[max_country, 'Value']}%"
        print(topCountryFB)

    if "*Instagram followers by top countries" in table.columns.values[0]:
        if not table.empty:
            table['Value'] = table['Value'].str.rstrip('%').astype('float')
            max_country = table['Value'].idxmax()
            topCountryIG = f"{table.loc[max_country, '*Instagram followers by top countriesTop countries']}, with {table.loc[max_country, 'Value']}%"
        print(topCountryIG)

