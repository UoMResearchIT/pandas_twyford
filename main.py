import pandas as pd
import glob
import numpy as np

excel_files = glob.glob('*.xlsx')
out_file = 'output.xlsx'

for file in excel_files:
    if file != out_file:
        df = pd.read_excel(file, header=1, usecols='B:AC')
        search_terms = ['geothermal', 'tidal energy', 'renewable energy', 'responsible energy', 'new energy',
                        'energy-efficient', 'clean energy', 'solar power', 'solar energy', 'photovoltaic', 'wind power',
                        'turbine-generated power', 'tidal energy', 'hydropower', 'geothermal power']

        df.columns = df.columns.str.strip().str.lower().str.replace('\n', ' ').str.replace(' ', '_').str.replace('(', '').str.replace(')', '')

        for term in search_terms:
            # Create an initial temp column (suffixed with '_temp') for current search term and insert search result 
            temp_col_name = term + '_temp'
            df[temp_col_name] = df.deal_text.str.contains(term, case=False)
            # Create a temp column (named after search term) and put in search term found.
            df[term] = np.where(df[temp_col_name] == True, term, '')
            # Drop the initial temp column
            df.drop(columns=[temp_col_name], inplace=True)
        # Concatonate all of the discovered (non-empty) search terms (columns with term as header)
        df['matches'] = df[search_terms].apply(lambda row: ';'.join(filter(None, row.values.astype(str))), axis=1)
        # Drop second lot of temp columns
        df.drop(columns=search_terms, inplace=True)

df.to_excel(out_file)
