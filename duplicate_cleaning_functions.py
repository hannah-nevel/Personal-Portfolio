import numpy as np
import pandas as pd
from datetime import datetime
from datetime import date


def main():

    file = pd.read_csv(r"C:\Users\Hannah Nevel\Documents\Duplicate Proj - All Related Files\Duplicate Check by Email - 10-18-2023\inv accts leads with phone.csv", encoding='ISO-8859-1', na_values='')
    df = pd.DataFrame(file)

    #create list of duplicates from original data set filtering by confirmed email address
    duplicates = df[df.duplicated(subset=['Confirmed Email Addresses'], keep=False) == True]
    
    #replace empty confirmed emails with NAN and drop from dataset
    duplicates['Confirmed Email Addresses'].replace('', np.nan, inplace=True)
    duplicates.dropna(subset=['Confirmed Email Addresses'], inplace=True)
   
    #remove rows where same account ID occurs multiple times, keep first occurence
    true_dups = duplicates.groupby('Account ID', as_index=False).first()

    #count occurences of confirmed email addresses and keep rows where count is greater than 1
    true_dups['Email Count'] = true_dups.groupby('Confirmed Email Addresses')['Confirmed Email Addresses'].transform('size')
    true_dups = true_dups[true_dups['Email Count'] > 1]

    #create dataframe of invested disposition
    invested_dups = true_dups[true_dups['Investment Disposition'] == 'Invested']
    #create dataframe of untouched disposition
    untouched_dups = true_dups[true_dups['Investment Disposition'] == 'Untouched']

    todays_date = str(date.today())
    filename = 'Duplicate List' + ' - test 2' + todays_date + '.xlsx'

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        true_dups.to_excel(writer, index = False, sheet_name ='All Duplicates')
        invested_dups.to_excel(writer, index = False, sheet_name ='Invested')
        untouched_dups.to_excel(writer, index = False, sheet_name ='Untouched')

    


main()

