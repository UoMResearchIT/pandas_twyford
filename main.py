import pandas as pd
import glob

excel_files = glob.glob('*.xlsx')
for file in excel_files:
    df = pd.read_excel(file, header=1, usecols='B:AC')
    search_terms = ['geothermal', 'tidal energy', 'renewable energy', 'responsible energy', 'new energy',
                    'energy-efficient', 'clean energy', 'solar power', 'solar energy', 'photovoltaic', 'wind power',
                    'turbine-generated power', 'tidal energy', 'hydropower', 'geothermal power']

    col_names = list(df.columns)
    alliance_participants = col_names[1]

    for term in search_terms:
        results = df.loc[df['Deal Text'].str.contains(term, case=False)]
        if not results.empty:
            print(term)
            print(results[alliance_participants].to_string())

