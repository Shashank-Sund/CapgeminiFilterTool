import pandas as pd
import numpy as np

#Change Path of input file here
df = pd.read_excel('testdata.xlsx')

#adding additonal columns
df['PM PM ID'] = np.nan
df['PM PM Name'] = np.nan
df['PM PM Designation'] = np.nan
df['PM PM PM ID'] = np.nan
df['PM PM PM Name'] = np.nan
df['PM PM PM Designation'] = np.nan
df['PM PM PM PM ID'] = np.nan
df['PM PM PM PM Name'] = np.nan
df['PM PM PM PM Designation'] = np.nan
df['PM PM PM PM PM ID'] = np.nan
df['PM PM PM PM PM Name'] = np.nan
df['PM PM PM PM PM Designation'] = np.nan
df['PM PM PM PM PM PM Name'] = np.nan
df['Ultimate Senior Director'] = np.nan

added_cols_dict = {'People Manager ID':['PM PM ID', 'PM PM Name', 'PM PM Designation'],
'PM PM ID':['PM PM PM ID', 'PM PM PM Name', 'PM PM PM Designation'],
'PM PM PM ID':['PM PM PM PM ID', 'PM PM PM PM Name', 'PM PM PM PM Designation'],
'PM PM PM PM ID':['PM PM PM PM PM ID', 'PM PM PM PM PM Name', 'PM PM PM PM PM Designation'],
'PM PM PM PM PM ID':['PM PM PM PM PM PM Name']}

class ultimate_sd(object):

    #used to determine which row or column in dict is being looked up
    class counter(object):
        def __init__(self, count):
            self.count = count

    def __init__(self, df, added_cols_dict):
        self.df = df
        self.added_cols_dict = added_cols_dict

    def execute(self):

        def lookup(item, key, value, counts, df):
            # lookup using the whatver record the count is on
            try:
                record = df.iloc[counts.count]
                counts.count += 1
                lookup_id = record[key]
                found = df[df.iloc[:, 0] == lookup_id].iloc[0]

                if 'ID' in value:
                    return found['People Manager ID']
                elif 'Name' in value:
                    return found['People Manager Name']
                elif 'Designation' in value:
                    return found['People Manager Designation']
                else:
                    return 4
            except:
                pass

        def get_sd(row):
            # gets senior directors for each record
            if row['Designation'] == 'Senior Director':
                return str(row['First Name']) + ' ' + str(row['Middle Name']) + ' ' + str(row['Last Name'])
            elif row['People Manager Designation'] == 'Senior Director':
                return row['People Manager Name']
            elif row['PM PM Designation'] == 'Senior Director':
                return row['PM PM Name']
            elif row['PM PM PM Designation'] == 'Senior Director':
                return row['PM PM PM Name']
            elif row['PM PM PM PM Designation'] == 'Senior Director':
                return row['PM PM PM PM Name']
            elif row['PM PM PM PM PM Designation'] == 'Senior Director':
                return row['PM PM PM PM PM Name']
            else:
                return " "

        col_counters = self.counter(0)
        for key, value in self.added_cols_dict.items():
            for item in value:
                counts = self.counter(0)
                value2 = value[col_counters.count]
                self.df[value2] = self.df[value2].apply(lookup, args=(key, value2, counts, self.df))
                #reset counter  back to 0 every three iterations
                col_counters.count += 1
                if col_counters.count == 3:
                    col_counters.count = 0
        self.df['Ultimate Senior Director'] = self.df.apply(get_sd, axis=1)

#create object instance passing in file and lookup ids with columns to be filled
output = ultimate_sd(df, added_cols_dict)
output.execute()

#output to excel
output.df.to_excel('output.xlsx')
#print(output.df)
