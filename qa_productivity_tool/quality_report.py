import os
import pandas as pd

def pf_lookup():
    qreport = Quality_Report()
    pc_dict = qreport.all_pcs_dict()
    print(qreport.return_product_family('40'))
protocol = pf_lookup
class Quality_Report(object):
    '''A class for other reports to inherit from'''
    def __init__(self,report_name,ref_path = r'C:\\Users\steven.hansen01\data_automation\\venv\\Cross_Reference_data_v0_1.xlsx',
    to_path = './Test_Reports' ):
        #self.og_data_path = og_data_path
        self.ref_path = ref_path
        self.report_name = report_name
        self.to_path = to_path
        self.all_pcs_dict()
        self.status_of_report = 0
    def meta_data(self):
        return self.r_filename,self.path,self.date
    def mostrecentreport(self,og_data_path = r"C:\Users\steven.hansen01\Downloads"):
        print(f'Searching for most recent {self.report_name} file ')
        self.og_data_path = og_data_path
        print(self.og_data_path)
        entries = {}
        with os.scandir(self.og_data_path) as it:
            for entry in it:
                if entry.name.startswith(self.report_name):
                #print(entry.name)
                    entries.update({entry.name:[entry.stat().st_ctime,entry.path]})
        for key in entries:
            if entries[key] == max(entries.values()):
                timepulled = pd.Timestamp(entries[key][0], unit = 's')
                day,month,year = timepulled.day,timepulled.month,timepulled.year
                datepulled = str(month)+"/"+str(day)+"/"+str(year)
                self.r_filename = key
                self.path = entries[key][1]
                self.date=datepulled
                return self.r_filename,self.path,self.date
    def all_pcs_dict(self):
        '''returns a dictionary with keys = Product Code and 
        values: 0 = description, 1 = UOM, 2 = Product Family and 
        3 = Value Stream, 4 = Lid Type'''
        self.all_pcs_df = pd.read_excel(open(self.ref_path,'rb'),sheet_name='All PCs')
        self.pcs_dict = {}
        for i, pc in enumerate(self.all_pcs_df['Product']):
            pc = str(pc)
            descrip = self.all_pcs_df.loc[i,'Description']
            uom = self.all_pcs_df.loc[i,'Per Case or 1 if UOM is ea']
            pf = self.all_pcs_df.loc[i, 'Product Family']
            vs = self.all_pcs_df.loc[i,'Value Stream']
            lid_type = self.all_pcs_df.loc[i,'Lid Type']
            self.pcs_dict[pc] = [descrip,uom,pf,vs,lid_type]
        return self.pcs_dict
    def return_product_family(self,pc):
        pc = pc.strip('-')
        try:
            return self.pcs_dict[pc][2]
        except KeyError:
            return 'pf not found'
    def return_value_stream(self,pc):
        pc = pc.strip('-')
        try:
            return self.pcs_dict[pc][3]
        except KeyError:
            return 'VS not found'
    def return_UOM(self,pc):
        pc = pc.strip('-')
        try:
            return self.pcs_dict[pc][1]
        except KeyError:
            return 'UOM not found'
    def return_lid_type(self,pc):
        pc = pc.strip('-')
        try:
            return self.pcs_dict[pc][4]
        except KeyError:
            return 'lid type not found'
    def tmpecc_remover(self,pc):
        try:
            pc = pc.strip('-')
            return pc.split('~')[0]
        except:
            return None
    def check_report_status(self):
        if self.status_of_report == 0:
            self.run_report()
    def index_compressor(self, df,name_of_index):
        new_df = pd.DataFrame()
        i = 0
        for name_of_index, frame in df.groupby(by = name_of_index):
            for col in frame.columns:

                new_df.loc[i,col] = str(frame[col].dropna().unique()).strip('[]').strip("'")
            i+=1
        return new_df
    def date_converter(self,df, col_name_list,date_format = '%d-%b-%Y'):
        
        
        for col in col_name_list:
            df[col] = df[col].fillna(0.)
            df[col]=pd.to_datetime(df[col],errors = 'coerce')
            df[col] = df[col].dt.strftime(date_format)
            #df[col] = list(map(lambda x: x.strftime(date_format),df[col] ))
        return df

if __name__ =='__main__':
    protocol()