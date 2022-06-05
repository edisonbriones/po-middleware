from pandas import read_csv, read_excel, DataFrame, concat, Series
from tkinter import Tk, filedialog
from glob import glob

def ask_filepath(title='Open File', folder=False):
    '''
    asks for filepath from user
    '''
    if folder:
        return filedialog.askdirectory(title=title)
    else:
        return filedialog.askopenfilename(title=title, filetypes=[('Excel files', '*.xls*')])

def consolidate_po():
    '''
    returns consolidated PO
    '''
    
    # function to split values
    def desc_split(text, delimiter, positions):
        data = text.split(delimiter)
        return [data[position].strip() for position in positions]
    
    # input from user
    fnames = ask_filepath(title='Select folder directory: ', folder=True)
    fnames = glob(fnames + r'\*.csv')
    
    # get csv with specific columns and combine
    dfs = [read_csv(fname, header=None, usecols=[0,1,10,14,16,17,18,19,20,25], parse_dates=[0,2,3]) for fname in fnames]
    df = concat(dfs, ignore_index=True)
    
    # remove trailing -00 in PO number
    df[1] = df[1].apply(lambda x: x.split('-')[0])
    
    # get store code and store name
    df[[15,16]] = df.loc[:,16].apply(lambda x: x.split(':')[-3].split(' ')[:2]).apply(Series)
    
    # remove leading zeros in store code
    df[15] = df[15].str.lstrip('0')
    
    # get material description, sku code, quantity, unit price
    df[[21, 22, 23, 24]] = df.loc[:, 19].apply(lambda x: desc_split(x, ':', [2,4,6,8])).apply(Series)
    
    # remove letters in packing quantity
    df[23] = df[23].str.extract('(\d+)')
    
    # remove column
    df.drop(columns=[19], inplace=True)
    
    # sort columns and rename
    cols = df.columns.to_list()
    cols = sorted(cols)
    df = df[cols]
    df.rename(columns={0:'PO Date', 1:'PO Number', 10:'Del Date', 14:'Cancel Date', 
                       15:'Store Code', 16:'Store Name', 17:'Line number', 18:'Barcode', 
                       21:'Material Description', 22:'SKU Code', 23:'Packing',24:'Unit Price',
                       20:'Qty', 25:'Amount'}, inplace=True)
    return df

def manage_po(data):
    '''
    returns managed file
    '''
    
    item_master_file = ask_filepath(title='Open SKU Master file')
    store_master_file = ask_filepath(title='Open Store Master file')
    status_master_file = ask_filepath(title='Open SAP Status file')
    
    # get item master file and map item code
    sku_master = read_excel(item_master_file, sheet_name=0, usecols=['Barcode','URC Code'], dtype={'URC Code':'str'})
    skus = sku_master.set_index(['Barcode'])['URC Code'].dropna().to_dict()
    data[['Mapped Item Code']] = data['Barcode'].map(skus)
    
    # get store master file, map customer code and delivery date
    store_mapping = read_excel(store_master_file, sheet_name=1, usecols=['Store Code','URC Customer Code', 'Del Sched'], 
                               parse_dates=['Del Sched'], dtype={'Store Code':'str','URC Customer Code': 'str'})
    store_code = store_mapping.set_index('Store Code')['URC Customer Code'].dropna().to_dict()
    store_del = store_mapping.set_index('Store Code')['Del Sched'].dropna().to_dict()
    data['Mapped Customer Code'] = data['Store Code'].map(store_code)
    data['Mapped delivery Date'] = data['Store Code'].map(store_del)
    
    # get status master file and map status
    sapstatus = read_excel(status_master_file, header=1, usecols=['URC Code', 'SAP STATUS'], dtype='str')
    status = sapstatus.set_index(['URC Code'])['SAP STATUS'].dropna().to_dict()
    data['Status'] = data['Mapped Item Code'].map(status)
    
    return data

def create_po_header(df):
    '''
    returns ORDERHDR file
    '''
    # selected columns exclude inactive items
    df = df.loc[(df['Status']!='Material Excluded'),
                ('PO Number', 'Mapped delivery Date', 'PO Date', 
                 'Mapped Customer Code')].drop_duplicates().reset_index(drop=True)
    
    # assign columns
    df = df.assign(Order_Type='Z8DO', 
                   Order_Reason=None, 
                   Sales_Org='BCFG', 
                   Dist_Channel=12, 
                   Division=97,
                   Sales_Group=None, 
                   Sales_Office=None, 
                   Del_Date=df['Mapped delivery Date'].dt.strftime('%Y%m%d'), 
                   PO_Type='EMAL',
                   Soldto=df['Mapped Customer Code'],
                   Shipto=df['Mapped Customer Code'], 
                   Credit_Ctrl_Area=None, 
                   Mster_contract_number=None, 
                   RefDocNo=None, 
                   zounr=None, 
                   REMARKS=None)
    
    # change format
    df['PO Date'] = df['PO Date'].dt.strftime('%Y%m%d')
    
    # replace column names
    df.columns = df.columns.str.replace('_', ' ')
    df.rename(columns={'RefDocNo':'Ref.Doc.No'}, inplace=True)
    
    # sort columns
    cols = ['PO Number', 'Order Type','Order Reason','Sales Org','Dist Channel','Division','Sales Group','Sales Office',
            'Del Date','PO Type','PO Date','Soldto','Shipto','Credit Ctrl Area','Mster contract number','Ref.Doc.No',
            'zounr','REMARKS']
    
    return df[cols]

def create_po_details(df):
    '''
    returns ORDERDTL file
    '''
    # select columns
    df = df.loc[(df['Status']!='Material Excluded'),
                ('PO Number', 'Mapped Item Code', 'Qty')].drop_duplicates().reset_index(drop=True)
    
    # add column headers
    df = df.assign(Unit='CS', Remarks=None, MatGrp5=None, CondGrp1=None, 
                   IO_Number=None, Condition_Type=None, Amount=None, Currency=None)
    
    # rename columns
    df.columns = df.columns.str.replace('_', ' ')
    df.rename(columns={'Mapped Item Code': 'Material Number'}, inplace=True)
        
    return df

def main():
    combined_po = consolidate_po()
    managed_po = manage_po(combined_po)
    orderhdr = create_po_header(managed_po)
    orderdtl = create_po_details(managed_po)
    
    folder = ask_filepath(title='Select output folder', folder=True)
    
    combined_po.to_csv(folder + 'Compiled csv file.csv', index=False)
    managed_po.to_csv(folder + 'Managed PO file.csv', index=False)
    orderhdr.to_csv(folder + 'ORDERHDR.csv', index=False)
    orderdtl.to_csv(folder + 'ORDERDTL.csv', index=False)
 
if __name__ == '__main__':
    root = Tk()
    root.withdraw()
    main()
