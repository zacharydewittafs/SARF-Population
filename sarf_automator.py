from pdf_filler import PdfFileFiller
import pandas as pd
import os
import shutil
import sys

def main():
    print('Starting process...')
    #Local directory the SARF template is located
    sarf_template_path = 'SARF_Template/'

    #Local directory the completed P&P file(s) is located 
    user_data_path = 'P&P_Files/'
    #Sheet name of the completed user data in the Excel file  
    user_data_sheetname = 'CRM Users'
    #Row number of the user data column headers 
    user_data_header_row_num = 2

    #Prefix for the newly created PDF files
    new_file_prefix = 'CRM_SARF_'

    if not os.path.exists(user_data_path):
        os.mkdir(user_data_path)
    
    if not os.path.exists(sarf_template_path):
        os.mkdir(sarf_template_path)

    automator = SarfAutomator(sarf_template_path)
    print('Loading user data files...')
    automator.load_data(user_data_path, user_data_sheetname, user_data_header_row_num)
    print('Generating SARFs...')
    automator.run(new_file_prefix)
    print('Process complete.')

class SarfAutomator():
    """
    Takes data from completed P&P Excel files and tranforms them into completed SARF PDF files

    Attributes
    ----------
    user_data: list
        List of dictionaries containing the user data from each input data source
    pdf_filler: PdfFileFiller
        Object used to read, populate, and create PDF files
    """
    
    def __init__(self, sarf_template_path):
        """
        Parameters
        ----------
        sarf_template_path: str
            Path of file directory that contains the SARF Template PDF file
        """

        self.user_data = []

        #Check if SARF path exists and contains only one template file
        if os.path.exists(sarf_template_path):
            if len(os.listdir(sarf_template_path)) < 1:
                raise ValueError(f"No SARF template found in {sarf_template_path}")
            elif len(os.listdir(sarf_template_path)) > 1:
                raise ValueError(f"There should only be one SARF template in {sarf_template_path}") 
            elif not os.listdir(sarf_template_path)[0].endswith('.pdf'):
                raise ValueError(f"SARF Template must be a .pdf")
        else:
            raise ValueError(f"SARF Template directory '{sarf_template_path}' cannot be found.")

        sarf_file_name = os.listdir(sarf_template_path)[0]
        self.pdf_filler = PdfFileFiller(sarf_template_path + sarf_file_name)

    def load_data(self, user_data_path, data_sheetname, header_row_num=0):
        """
        Reads each Excel file in the User Data file path and maps it to the required SARF fields

        Parameters
        ----------
        user_data_path: str
            Path of file directory that contains the completed user data Excel files
        data_sheetname: str
            Name of the Excel sheet that contains the user data in each of the Excel files
        header_row_num: int, optional
            Row number (0-based) that contains the column headers of the user data (defaults to 0)
        """
        
        #List of user data records from each Excel file
        user_data_filenames = []

        #SARF PDF field mappings
        #Mappings with list values will be concatenated together with each value separated by ', '
        #Mappings with dictionary values, will map the dict key column by default, but if that value is blank it will map the dict value column instead
        field_mappings = {
            '1 Name':['Last Name', 'First Name', 'Middle Name'],
            '1 Preferred Email':'Email Address\n(state.gov preferred)',
            '1 Job Title':'Job Title',
            '1 Employment Type':'Employment Type',
            '1 Office  Post':'Office',
            '1 Notes':'Bureau',
            '1 Timezone':'Time Zone',
            '1 Existing Okta Account':'Do you have an existing Okta account?  ',
            '1 DOS email':{"DoS Email Address \n(only if @state.gov wasn't already listed in column E)":'Email Address\n(state.gov preferred)'},
            '1 no email':"Do you have a DoS Email Address? \n(only if @state.gov wasn't already listed in column E)",
            '1 mobile device':'Do you have access to a mobile phone in your workplace?',
            '1 mobile app':'Do you have the ability to download the Okta Verify moble app to a work or personal phone, and use it at your workplace?',
            '1 CRM User Type':'User Type',
            '1 Exec Contacts':'Executive Contacts ',
            '1 Printing':'Event Printing'
        }

        #SARF PDF Override Value mappings for checkbox and radio button values that require specific values to select the correct option 
        #and includes a default string value in cases of a blank value in the column
        value_override_mappings = {
            '1 Existing Okta Account':{'':'Yes','Yes':'/0', 'No':'/1', "I Don't Know":'/2'},
            '1 no email':{'':'/No', 'Yes':'/No', 'No':'Yes', '/No':'No'},
            '1 mobile device':{'': 'Yes', 'Yes':'/Yes', 'No':'/No'},
            '1 mobile app':{'': 'Yes', 'Yes':'/0', 'No':'/No'},
            '1 CRM User Type':{'':'CRM User', 'CRM User':'/1#20CRM#20User', 'CRM Mission/Office Admin':'/1#20Admin', 'CRM Contacts Only User':'/1#20Contacts'},
            '1 Exec Contacts':{'':'No'},
            '1 Printing':{'':'No'}
        }

        #SARF PDF field mappings for default string values 
        default_string_mappings = {
            '1 Request Type':'New User',
            '1 Application Access':'CRM Only'
        }
        
        #Check if user data path exists
        if not os.path.exists(user_data_path):
            raise ValueError(f"User Data directory '{user_data_path}' cannot be found.")
        
        #Get all Excel files (.xls or .xlsx) files in the directory
        for file_name in os.listdir(user_data_path):
            if file_name.split('.')[1] in ['xls','xlsx']:
                user_data_filenames.append(user_data_path + file_name)
        
        #Flag if no valid Excel files were found in the directory
        if not user_data_filenames:
            raise ValueError (f"No valid user data files in '{user_data_path}' found. Please make sure files are in .xls or .xlsx format and try again.")

        #Store the user data from each valid Excel file
        for data_filename in user_data_filenames:
            #Read in the Excel data into a DataFrame
            user_data = pd.read_excel(data_filename, sheet_name=None, encoding='utf-8-sig', skiprows=header_row_num)
            
            #Flag if the user data sheet cannot be found in the Excel file
            if data_sheetname not in user_data.keys():
                raise ValueError(f"Sheet name: '{data_sheetname}' could not be found in file: '{data_sheetname}' - please confirm sheet name and try again.")

            #Get the data in the Excel sheet
            user_data = user_data[data_sheetname]
            #Exclude the Example row
            user_data = user_data[user_data['User Number'] != 'Example']
            #Remove any rows that do not have the First and Last Name columns completed
            user_data.dropna(subset=['First Name', 'Last Name'], inplace=True)
            #Convert all values to strings
            user_data = user_data.fillna('').applymap(lambda x: str(x))

            #Send warning message if any of the values in the specified columns have a blank value
            flag_blank_fields = ['User Type']
            for flag_field in flag_blank_fields:
                flagged_data = user_data[user_data[flag_field] == '']
                if flagged_data.shape[0] > 0:
                    print(f"***WARNING: Blank value(s) found in {flag_field} field in {data_filename} - Value(s) will be defaulted.***")

            #Create new DataFrame with SARF PDF field mappings and formatted values
            formatted_user_data = pd.DataFrame(columns=field_mappings.keys())
            #Map field mappings
            for col in formatted_user_data.columns:
                if isinstance(field_mappings[col], list):
                    #Map the value of each column separated by ', '
                    formatted_user_data[col] = user_data[field_mappings[col]].apply(lambda row: ', '.join(row).strip(' ,'), axis=1)
                elif isinstance(field_mappings[col], dict):
                    #Map the key value by default, but use the key value if the default column is blank
                    default_field = list(field_mappings[col].keys())[0]
                    backup_field = field_mappings[col][default_field]
                    formatted_user_data[col] = user_data[[default_field, backup_field]].apply(lambda row: row[default_field] if row[default_field] != '' else row[backup_field], axis=1)
                else:
                    formatted_user_data[col] = user_data[field_mappings[col]]
            formatted_user_data['id'] = user_data['User Number']
            
            #Update input values with PDF equivalent values to appear correctly in generated PDF file
            if value_override_mappings:
                for field in value_override_mappings:
                    for value in value_override_mappings[field]:
                        formatted_user_data[field] = formatted_user_data[field].apply(lambda x: value_override_mappings[field][value] if x.lower()==value.lower() else x)
            
            #Map static string value to the specified field of all records in the data set
            if default_string_mappings:
                for field in default_string_mappings:
                    formatted_user_data[field] = default_string_mappings[field]

            #Convert data to dictionary and add to list of all user data to map to PDF file
            self.user_data.append(formatted_user_data.to_dict(orient='records'))
            print('Done')

    def run(self, new_pdf_filename_prefix=None):
        """
        Execute the process of taking all user data information and generate a separate, completed SARF PDF file for each Excel data source

        Parameters
        ----------
        new_pdf_filename_prefix: str, optional
            The prefix of the new SARF PDF file name (defaults to None)
        """
        
        #Go through each group of user data and generate a completed SARF PDF file for each dataset
        for cur_user_data in self.user_data:
            #Set the file name prefix to an empty string if one was not specified 
            if not new_pdf_filename_prefix:
                new_pdf_filename_prefix = ''
            #Create the new file name from the passed prefix value and the Bureau value of the first user data record in the dataset
            new_sarf_filename = new_pdf_filename_prefix + cur_user_data[0]['1 Notes'] + '.pdf'


            print(f'Creating {new_sarf_filename}...')
            #Create a temporary directory for the current dataset
            new_sarf_temp_directory = new_sarf_filename.replace('.pdf', '_temp')
            if not os.path.exists(new_sarf_temp_directory):
                os.mkdir(new_sarf_temp_directory)
            
            #For each user data record in the current dataset, create a new SARF PDF file with the populated field values into the temp directory
            for user_record in cur_user_data:
                #The page number the SARF user data section resides (0 based)
                sarf_user_info_page_num = 2
                self.pdf_filler.update_pdf_form_values(new_sarf_temp_directory + '/' + new_sarf_filename , user_record, sarf_user_info_page_num)

            #Merge all individual user record PDF files generated into a single PDF file
            files = []
            for file in os.listdir(new_sarf_temp_directory):
                if file.endswith('.pdf'):
                    files.append(new_sarf_temp_directory + '/' + file)
            #Sort files by the unique id of the data (e.g. User Number)
            files = sorted(files, key=lambda x: int(x.split('.')[0].split('_')[-1]))
            #Merge the PDFs
            self.pdf_filler.merge_pdfs(files, new_sarf_filename, header_pages=sarf_user_info_page_num)
            #Delete the temporary directory and all files within it
            shutil.rmtree(new_sarf_temp_directory)


if __name__ == "__main__":
    main()
    sys.exit()
        
