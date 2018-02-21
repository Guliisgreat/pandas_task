import pandas as pd

def main():
   # read excel
    excel_files = 'Missing.xlsx'
    address = pd.read_excel(excel_files)
   # clean data
    address.fillna('none', inplace=True)
    address.loc[address.apply(lambda x: x['Town'] <= 99999999, axis=1), 'Town'] = 'Fan Cheng is fat'
   # insert two column (contain city? and pass?)
    (num_row, num_column) = address.shape
    address.insert(loc=num_column, column= 'contain?', value=int)
    address.insert(loc=num_column+1, column = 'pass or fail?', value=str)
   # address contain city or not?
    address.loc[address.apply(lambda x: x['Town'] in x['FullAddress'], axis=1), 'contain?'] = 1
    address.loc[address.apply(lambda x: x['Town'] not in x['FullAddress'], axis=1), 'contain?'] = 0
   # pass or not
    # count frequncy to filter the object with only one address
    address['freq'] = address.groupby('FenergoID')['FenergoID'].transform('count')
    address.loc[address['freq'] ==1 & address.apply(lambda x: x['Town'] in x['FullAddress'], axis=1), 'pass or fail?'] = 'pass'
    address.loc[address['freq'] == 1 & address.apply(lambda x: x['Town'] not in x['FullAddress'], axis=1), 'pass or fail?'] = 'fail'
    # multiple address
    address.loc[(address['freq'] > 1) & (address['AddressType'] == 'Registered') & (address.apply(lambda x: x['Town'] in x['FullAddress'], axis=1)), 'pass or fail?'] = 'pass'
    address.loc[(address['freq'] > 1) & (address['AddressType'] == 'Registered') & (address.apply(lambda x: x['Town'] not in x['FullAddress'], axis=1)), 'pass or fail?'] = 'fail'
    address.loc[(address['pass or fail?'] != 'pass') & (address['pass or fail?'] != 'fail'),'pass or fail?'] = 'nothing'
   # delete frequency column
    del address['freq']

    print(address)
    address.to_excel('output.xlsx')

if __name__ == '__main__':
    main()