from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject, createStringObject

class PdfFileWriter2(PdfFileWriter):
    #Inherits/extends the PyPDF2 built-in PdfFileWriter class (https://pythonhosted.org/PyPDF2/PdfFileWriter.html)
    def __init__(self):
        super().__init__()
    
    def update_checkbox_radio_field_values(self, page, fields):
        """
        Updates check box and radio fields on a pdf form

        Parameters
        ----------
        page: PageObject
            The page of the pdf
        fields: dict
            Dictionary of the field names and values to be updated on the pdf page
        """

        #Go through each annotation on the pdf page (i.e. form fields)
        for i in range(0, len(page['/Annots'])):
            #Get the current annotation
            annotation = page['/Annots'][i].getObject()
            if '/Parent' in annotation:
                #If the field is a radio button, then get the parent annotation
                annotation = annotation['/Parent']
            
            #Update each specified field in the PDF page
            for fieldname in fields:
                #Check if the field name of the current annotation is equal to our field name
                if annotation.get('/T') == fieldname:
                    #Update the form value and the appearance stream values
                    if '/Kids' in annotation: #if radio button
                        if fields[fieldname] == 'Yes':
                            annotation.update({
                                NameObject('/V'):NameObject('/0'),
                                NameObject('/AS'):NameObject('/0')
                            })
                        elif fields[fieldname] == 'No':
                            annotation.update({
                                NameObject('/V'):NameObject('/1'), 
                                NameObject('/AS'):NameObject('/1')
                            })
                        elif fields[fieldname] == '/0':
                            annotation.update({
                                NameObject('/V'):NameObject('/0'),
                                NameObject('/AS'):NameObject('/0')
                            })

                        elif fields[fieldname] == '/1':
                            annotation.update({
                                NameObject('/V'):NameObject('/1'),
                                NameObject('/AS'):NameObject('/1')
                            })
                            
                        elif fields[fieldname] == '/Yes':
                            annotation.update({
                                NameObject('/V'):NameObject('/Yes'),
                                NameObject('/AS'):NameObject('/Yes')
                            })
                            
                        elif fields[fieldname] == '/No':
                            annotation.update({
                                NameObject('/V'):NameObject('/No'),
                                NameObject('/AS'):NameObject('/No')
                            })
                        else:
                            annotation.update({
                                NameObject('/V'):NameObject(fields[fieldname]),
                                NameObject('/AS'):NameObject(fields[fieldname])
                        })
                    else:
                        if fields[fieldname] == 'Yes':
                            #Check the checkbox field
                            annotation.update({
                                NameObject('/V'):NameObject('/Yes'), #value
                                NameObject('/AS'):NameObject('/Yes') #appearance stream
                            })
                        else:
                            #Uncheck the checkbox field
                            annotation.update({
                                NameObject('/V'):NameObject('/Off'),
                                NameObject('/AS'):NameObject('/Off')
                            })

    def convert_dropdown_to_text(self, page, fields):
        """
        Converts fields identified as dropdown fields to a text field to display the value passed
        Note: This will remove the dropdown options from these field in the new pdf generated

        Parameters
        ----------
        page: PageObject
            The page of the pdf
        fields: dict
            Dictionary of the field names and values to be updated on the pdf page
        """

        for i in range(0, len(page['/Annots'])):
            #Get the current annotation
            annotation = page['/Annots'][i].getObject()
            for fieldname in fields:
                #Check if the field name of the current annotation is equal to our field name
                if annotation.get('/T') == fieldname:
                    #Change the field type of the current form field from a choice field to a text field
                    annotation.update({
                        NameObject('/FT'): NameObject('/Tx')
                    })


    def set_need_appearances(self):
        """
        Set NeedAppearances flag on interactive form in order to see 
        field values appear in form fields
        """

        catalog = self._root_object
        #Get the AcroForm tree
        if '/AcroForm' not in catalog:
            self._root_object.update({
                NameObject('/AcroForm'):IndirectObject(len(self._objects), 0, self)
            })
        
        need_appearances = NameObject('/NeedAppearances')
        self._root_object['/AcroForm'][need_appearances] = BooleanObject(True)

class PdfFileFiller():
    """
    This class completes the form fields of a pdf and the completed fields into a new pdf file

    Attributes
    ----------
    pdf_template_file_path: str
        The file path/name of the PDF template used to generate new, populated PDF files
    pdf_template: PdfFileReader
        A PdfFileReader object of the PDF file template 
    """

    def __init__(self, pdf_template_file_path):
        """
        Parameters
        ----------
        pdf_template_file_path: str
            The file path/name of the PDF template used to generate new, populated PDF files
        """

        self.pdf_template_file_path = pdf_template_file_path
        self.pdf_template = PdfFileReader(pdf_template_file_path)

        if "/AcroForm" in self.pdf_template.trailer["/Root"]:
            self.pdf_template.trailer["/Root"]["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)})
    
    def update_pdf_form_values(self, new_pdf_file_name, data, pageNum=0):
        """
        Creates a new PDF file from the PDF template with form values populated based on the field name and values passed

        Parameters
        ----------
        new_pdf_file_name: str
            The new PDF file name
        data: dict
            Dictionary of the name of the PDF form fields and their associated values (requires having a unique 'id' field)
        pageNum: int, optional
            Page number of the PDF Template to populate data into (default is 0 - e.g. the first page)
        """

        #Create a new PdfFileReader instance (this is required due PyPDF2 retaining any modification to a PageObject for the PdfFileReader)
        pdf_template = PdfFileReader(self.pdf_template_file_path)
        if "/AcroForm" in pdf_template.trailer["/Root"]:
            pdf_template.trailer["/Root"]["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)})


        new_pdf = PdfFileWriter2()
        #Set NeedAppearances on new PdfFileWriter so it's applied to the new PDF
        trailer = pdf_template.trailer["/Root"]["/AcroForm"]
        new_pdf._root_object.update({
            NameObject('/AcroForm'): trailer})
        new_pdf.set_need_appearances()
        if "/AcroForm" in new_pdf._root_object:
            new_pdf._root_object["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)})

        #Get page from PDF template
        page = pdf_template.getPage(pageNum)
        #Add the page to the new PDF
        new_pdf.addPage(page)

        pdf_text_fields = []
        pdf_checkbox_radio_fields = []
        pdf_dropdown_fields = []

        #Rename each form field name in the page to a unique value to prevent the data in the first page from being 
        #written to all pages if the PDF file were to be merged with another PDF file from the same template
        for j in range(0, len(page['/Annots'])):
            writer_annot = page['/Annots'][j].getObject()
            if '/Parent' in writer_annot: #if radio button field
                writer_annot = writer_annot['/Parent']
                if writer_annot.get('/T').endswith('###'+data['id']):
                    continue
            #Update the form field name
            writer_annot.update({
                NameObject("/T"): createStringObject(writer_annot.get('/T')+'###'+data['id'])
            })

            #Determine the field types of each form field in the PDF page
            if writer_annot.get('/FT') == '/Tx':
                #Is a text field
                pdf_text_fields.append(writer_annot.get('/T'))
            elif writer_annot.get('/FT') == '/Btn':
                #Is a button field (i.e. checkbox or radio field)
                pdf_checkbox_radio_fields.append(writer_annot.get('/T'))
            elif writer_annot.get('/FT') == '/Ch' and '/Opt' in writer_annot:
                #Is a choice field (i.e. dropdown field)
                pdf_dropdown_fields.append(writer_annot.get('/T'))

        #Filter the data into the appropriate field type
        unique_id = '###' + data['id']
        text_field_data = {k + unique_id:v for k,v in data.items() if k + unique_id in pdf_text_fields or k in pdf_dropdown_fields}
        checkbox_radio_field_data = {k + unique_id:v for k,v in data.items() if k + unique_id in pdf_checkbox_radio_fields}
        dropdown_field_data = {k + unique_id:v for k,v in data.items() if k + unique_id in pdf_dropdown_fields}

        #Convert dropdown field to text field
        new_pdf.convert_dropdown_to_text(page, dropdown_field_data)

        #Update the fields in the PDF page for each field type
        new_pdf.updatePageFormFieldValues(page, text_field_data)
        new_pdf.update_checkbox_radio_field_values(page, checkbox_radio_field_data)
        new_pdf.updatePageFormFieldValues(page, dropdown_field_data)

        new_pdf.set_need_appearances()
        if "/AcroForm" in new_pdf._root_object:
            new_pdf._root_object["/AcroForm"].update(
                {NameObject("/NeedAppearances"): BooleanObject(True)})

        #Create a new PDF file with the unique id of the data tagged at the end with the completed form fields
        new_pdf_file = new_pdf_file_name.replace('.pdf', '_' + str(data['id']) + '.pdf')
        with open(new_pdf_file, 'wb') as out:
            new_pdf.write(out)
    
    def merge_pdfs(self, input_filenames, output_filename, header_pages=None):
        """
        Merges multiple PDF files into a new, single PDF file

        Parameters
        ----------
        input_filenames: list
            List of PDF file names to be combined into a single PDF file
        output_filename: str
            Name of the new merged PDF file 
        header_pages: int, optional
            Number of pages from the PDF template to include in the beginning of the merged PDF file 
            (ex. header_pages = 2 will include the first and second page of the PDF template as the first and second pages of the merged PDF file)
        """

        writer = PdfFileWriter2()
        writer.set_need_appearances()

        #Add the header pages to the new PDF file
        for i in range(0, header_pages):
            writer.addPage(self.pdf_template.getPage(i))
        
        #Add each page from each PDF file to the end of the new PDF file
        for input_file in input_filenames:
            file = PdfFileReader(input_file)
            for pageNum in range(0, file.getNumPages()):
                page = file.getPage(pageNum - 1)
                writer.addPage(page)
        
        #Create and write to new PDF file
        new_file = open(output_filename, 'wb')
        writer.write(new_file)
        new_file.close()