import pandas as pd
import glob

excel_files = glob.glob('*.xlsx')
out_file = 'output.xlsx'
df_out = pd.DataFrame()
for file in excel_files:
    if file != out_file:
        df = pd.read_excel(file, header=1, usecols='B:AC')
        search_terms = ['geothermal', 'tidal energy', 'renewable energy', 'responsible energy', 'new energy',
                        'energy-efficient', 'clean energy', 'solar power', 'solar energy', 'photovoltaic', 'wind power',
                        'turbine-generated power', 'tidal energy', 'hydropower', 'geothermal power']

        df.columns = df.columns.str.strip().str.lower().str.replace('\n', ' ').str.replace(' ', '_').str.replace('(', '').str.replace(')', '')

        for term in search_terms:
            results = df.loc[df.deal_text.str.contains(term, case=False)]
            if not results.empty:
                if df_out.empty:
                    df_out = results
                else:
                    df_out = pd.concat([df_out, results])

df_out.to_excel(out_file)
