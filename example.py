from pdf2xlsx import Pdf2XlsxConverter

# set user's proxy if it is necessary
my_proxy = None

# initialize class object
user_converter = Pdf2XlsxConverter()

# conversion of the single pdf file
user_converter.SetFiles(pdf_files_paths = "C:\\USER_PDF_FILES\\example_1.pdf", output_path = "C:\\USER_XLSX_FILES\\")
user_converter.WebConvert(user_proxy = my_proxy)

# conversion of the multiple pdf files
user_converter.SetFiles(pdf_files_paths = ["C:\\USER_PDF_FILES\\example_1.pdf", "C:\\USER_PDF_FILES\\example_2.pdf"])
user_converter.WebConvert(user_proxy = my_proxy)