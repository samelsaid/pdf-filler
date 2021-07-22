#! python3

# Author: Elsaid Salem
# Credits: Gabor Szabo, GeeksForGeeks, Andrew Krcatovich
# Notes: use .xls format for the excel, make sure to add a no_images dir to the output location in -o param

#tahsine: x 23, y 445


# Known Issues:
'''
- two files are created, need to delete image-less one, for now sort by .pdf extension
- add more args for options, but keep them optional  ?
    - find a way to reduce the number of needed paramters to execute
- need better error / exception handling
'''

from reportlab.pdfgen import canvas
import argparse
import xlrd
import pdfrw
from PyPDF2 import PdfFileWriter, PdfFileReader
import io

# need to add in argparse and assign variables to parameters

parser = argparse.ArgumentParser()
parser.add_argument("--excel", dest="excel_path", help="Path to your excel file, required")
parser.add_argument("--sheet", dest="sheet_index", help="Index of the sheet to load data from, 0 being the first sheet. Default value is 0", default=0)
parser.add_argument("--pdf", dest="pdf_path", help="Path to your pdf template file, required")
parser.add_argument("--fill-images", dest="fill_images", help="whether or not to fill images, default is yes", default="y")
parser.add_argument("--images", dest="image_path", help="Path to the directory of your images, don't include terminal / , required if filling images . default value is ./excel_to_pdf", default="./excel_to_pdf")
parser.add_argument("--img-field", dest="name_field", help="Which field in your excel corresponds to the image names, required if filling images")
parser.add_argument("--img-extension", dest="img_ext", help="What extension does your image use (eg png, jpg), required if filling images, don't include the '.' - default value is .jpg", default=".jpg")
parser.add_argument("--output", "-o", dest="pdf_output", help="fixed name of your output file (the part that doesn't change, don't include terminal /). Default value is ./excel_to_pdf", default="./excel_to_pdf/")
parser.add_argument("--x-axis", "-x", dest="x_loc", help="image location on x (0,0 is bottom left of page). Default value is 23", type=int, default=23)
parser.add_argument("--y-axis", "-y", dest="y_loc", help="image location on x (0,0 is bottom left of page). Default value is 445", type=int, default=445)
parser.add_argument("--width", "-w", dest="img_width", help="Image size based on width, keeps aspect ration", type=int, default=100)



# make sure your excel columns and pdf fields have the same names
# loads the excel data to fill the pdf
def load_excel(path, sheet):
    # get excel path
    excel_path = path
    # load workbook
    workbook = xlrd.open_workbook(excel_path)
    # load sheet from workbook
    data_table = workbook.sheet_by_index(sheet)

    # collect some ranges for the loop
    tot_rows = data_table.nrows
    tot_cols = data_table.ncols

    # initialze a dictionary for data storage
    data_dict = {}

    # collect the data itself by rows
    for i in range(2,tot_rows):
        for ii in range(tot_cols):
            col_name = data_table.cell_value(1, ii)
            data_dict[col_name] = data_table.cell_value(i, ii)
        #print(data_dict)
        yield data_dict

# adds the image to the file 
# Note: fill the pdf first, then add the image
def add_image(image_name, image_root, pdf_in_path, pdf_in_name, x_loc, y_loc, img_width):

    in_pdf_file = pdf_in_path + pdf_in_name
    out_pdf_file = f"{pdf_in_path}../{pdf_in_name}.pdf"
    img_file = f'{image_root}/{int(image_name):06}.jpg'
 
    packet = io.BytesIO()
    can = canvas.Canvas(packet)
    #can.drawString(10, 100, "Hello world")
    x_start = x_loc 
    y_start = y_loc
    can.drawImage(img_file, x_start, y_start, img_width, preserveAspectRatio=True, mask='auto')
    can.showPage()

    can.save()
 
    #move to the beginning of the StringIO buffer
    packet.seek(0)
 
    new_pdf = PdfFileReader(packet)
 
    # read the existing PDF
    existing_pdf = PdfFileReader(open(in_pdf_file, "rb"))
    output = PdfFileWriter()
 
 # for multiple pages to add to
    # for i in range(len(existing_pdf.pages)):
    #     page = existing_pdf.getPage(i)
    #     page.mergePage(new_pdf.getPage(i))
    #     output.addPage(page)

# for adding image to just the first page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
 
    outputStream = open(out_pdf_file, "wb")
    output.write(outputStream)
    outputStream.close()
 

# use this by calling fill_pdf in a for each of load_excel
def fill_pdf(input_pdf_path, data_dict, pdf_output, fill_images):

    # some needed hearders to parse the pdf
    ANNOT_KEY = '/Annots'
    ANNOT_FIELD_KEY = '/T'
    ANNOT_VAL_KEY = '/V'
    ANNOT_RECT_KEY = '/Rect'
    SUBTYPE_KEY = '/Subtype'
    WIDGET_SUBTYPE_KEY = '/Widget'

    
    if_images_path = "no_images/" if fill_images else ""
    if_image = ".pdf" if not fill_images else ""
# name of the final pdf starts here

    # the file path for filling
    output_pdf_path = f"{pdf_output}{if_images_path}{data_dict['Department']}_{data_dict['Employee Name']}{if_image}"
    # breaking up the path for the image filling, not used in this function
    output_pdf_root = f"{pdf_output}{if_images_path}"
    output_pdf_name = f"{data_dict['Department']}_{data_dict['Employee Name']}{if_image}" 

    # reading pdf fields and filling them in
    template_pdf = pdfrw.PdfReader(input_pdf_path)
    for page in template_pdf.pages:
        annotations = page[ANNOT_KEY]
        for annotation in annotations:
            if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                if annotation[ANNOT_FIELD_KEY]:
                    key = annotation[ANNOT_FIELD_KEY][1:-1]
                    if key in data_dict.keys():
                        if type(data_dict[key]) == bool:
                            if data_dict[key] == True:
                                annotation.update(pdfrw.PdfDict(
                                    AS=pdfrw.PdfName('Yes')))
                        else:
                            annotation.update(
                                pdfrw.PdfDict(V=f'{data_dict[key]}')
                            )
                            annotation.update(pdfrw.PdfDict(AP=''))
    # refreshes the page to actually show the changes
    template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true'))) 
    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)
    # need the file name to add the image to the correct file
    return output_pdf_root, output_pdf_name


def main():
    # collect params
    try:
        excel_path = parser.parse_args().excel_path
        sheet_index = int(parser.parse_args().sheet_index)
        pdf_path  = parser.parse_args().pdf_path
        image_path = parser.parse_args().image_path
        name_field  = parser.parse_args().name_field
        img_ext  = parser.parse_args().img_ext
        pdf_output  = parser.parse_args().pdf_output
        x_loc  = parser.parse_args().x_loc
        yloc  = parser.parse_args().y_loc
        img_width = parser.parse_args().img_width
        fill_images = True if parser.parse_args().fill_images.lower().startswith("y") else False
       
    except:
        print("Make sure you provide all the necessary parameters!\nPlease run with -h first to understand the options...")
        exit(1)

    # collect data row by row 
    for emp_data in load_excel(excel_path,sheet_index):
        # fill template pdf with data
        out_file = fill_pdf(pdf_path, emp_data, pdf_output, fill_images)
        print(out_file)
        # then add the image
        if fill_images:
            add_image(emp_data[name_field],image_path,out_file[0], out_file[1],x_loc,yloc, img_width)
            

if __name__ == "__main__":
    main()
    print("\n\nDone!")