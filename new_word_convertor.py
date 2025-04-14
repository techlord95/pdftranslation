import aspose.pdf as apdf
from io import FileIO
from os import path
path_infile = ("C:/Users/Srijan/Downloads/Doc1.pdf")
path_outfile = ("C:/Users/Srijan/Downloads/Doc1_converted.docx")

document = apdf.Document(path_infile)
save_options = apdf.DocSaveOptions()
save_options.format = apdf.DocSaveOptions.DocFormat.DOC_X
document.save(path_outfile, save_options)

# print(infile + " converted into " + outfile)