import requests
import copy
import os
import time

"""
Pdf2XlsxConverter class, which can be used for a conversion of pdf files into xlsx (MS Excel 2010) files
via internet (conversion will be through www.pdftoexcel.com). This method saves the structure of your pdf files.
"""
class Pdf2XlsxConverter:

    # initialize class object
    def __init__(self):

        self.__excel_paths = [] # absolute paths of the converted xlsx files
        self.__output_path = "" # output file directory
        self.__pdf_paths = [] # absolute paths of the pdf files

    # set input data
    def SetFiles(self, pdf_files_paths, output_path = None):

        self.__excel_paths.clear()

        if output_path == None:
            self.__output_path = os.getcwd() + "\\" # default directory is an executable script's directory
        else:
            self.__output_path = output_path

        self.__pdf_paths.clear()

        if type(pdf_files_paths).__name__ == "str": # single pdf file
            self.__pdf_paths = copy.deepcopy([pdf_files_paths])
        elif type(pdf_files_paths).__name__ == "list": # multiple pdf files
            self.__pdf_paths = copy.deepcopy(pdf_files_paths)

    # convert pdf files into xlsx files using web, namely www.pdftoexcel.com
    # arguments:
    # user_proxy - user's proxy
    # max_waiting_time - maximum time (in sec) to wait for the end of the conversion process
    # checking_time - interval (in sec) to delay the conversion
    def WebConvert(self, user_proxy = None, max_waiting_time = 60, checking_time = 3):

        if max_waiting_time <= 0:
            max_waiting_time = 60
        if checking_time <= 0:
            checking_time = 3

        for pdf_file_path in self.__pdf_paths:

            print("Conversion of the file {}\n".format(pdf_file_path))
            user_file = {"Filedata": open(pdf_file_path, "rb")}

            # uploading POST request
            uploading_resp = requests.post(url = "https://www.pdftoexcel.com/upload.instant.php", proxies = user_proxy, files = user_file).json()

            if uploading_resp["status"] == "1": # if pdf file was uploaded successfully...

                res_file_id = uploading_resp["jobId"] # get unique file ID
                start_time = time.time() # set current time

                # wait for the end of the conversion...
                while True:

                    # set the delay
                    time.sleep(checking_time)

                    # checking GET request
                    waiting_resp = requests.get("https://www.pdftoexcel.com/getIsConverted.php?jobId="+ res_file_id, proxies = user_proxy).json()

                    # stopping criterion for the delay
                    if (time.time() - start_time) > max_waiting_time or waiting_resp["status"] == "converted":
                        break

                # downloading GET request
                downloading_resp = requests.get("https://www.pdftoexcel.com/fetch.php?id=" + res_file_id, proxies = user_proxy)

                # create output file's path
                self.__excel_paths.append(self.__output_path + os.path.splitext(os.path.basename(pdf_file_path))[0] + ".xlsx")

                # write data to the file
                with open(self.__excel_paths[-1], "wb") as f:
                    for chunk in downloading_resp.iter_content(100000):
                        f.write(chunk)

                print("Success!\n")

            else: # if pdf file was uploaded unsuccessfully...

                print("Failure!\n")
                continue

            # close input pdf file
            input_file = open(pdf_file_path, "r")
            input_file.close()

        print("Conversion has been completed!\n")