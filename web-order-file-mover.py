import os
import datetime
import PyPDF2
import re
import time
import configparser
import shutil
import pyperclip


class Globals:
    def __init__(self):
        self.process_date_dt = datetime.date.today()
        self.process_date_str = datetime.datetime.strftime(datetime.datetime.today(), "%Y%m%d")
        self.report_files = set()
        self.art_files = set()
        self.dated_files = set()
        self.processing_directory = os.path.join("\\\\JTSRV3", "Print Facility", "Job Ticket Feed docs",
                                                 "WebToPrint")
        self.count_art_files = 1

        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

    def get_report_files(self):
        """Sets self.report_files to a set of packing slips, work orders, daily reports for processing date"""
        srch = re.compile("[\w]*[\d]{8}.pdf")
        self.report_files = set((list(filter(srch.match, self.dated_files))))

    def process_date_files(self):
        """Sets self.dated_files to a list of all files in
        processing directory with file modified property the same as processing date"""

        file_list = [f for f in os.listdir(self.processing_directory)]

        self.dated_files = set([f for f in file_list
                                if datetime.datetime.strftime(datetime.datetime.fromtimestamp(
                                 os.path.getmtime(os.path.join(self.processing_directory, f))), "%Y%m%d") ==
                                self.process_date_str])

    def get_art_files(self):
        """Sets self.art_files to a set of artwork files for processing date"""
        self.art_files = self.dated_files
        self.art_files.difference_update(self.report_files)

    def move_farm_bureau_art(self):
        # savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
        #                         "In Progress", "01-Web Order Art", "FB Monthly Web Order")

        savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
                                "In Progress", "01-Web Order Art", "FB Monthly Web Order TEST")

        file_search = re.compile("FB[\s\S]*.pdf")
        file_match = (list(filter(file_search.match, self.art_files)))
        if file_match:
            try:
                job_number = self.config['farmbureau'][self.process_date_str[:6]]
                job_directory = [f for f in os.listdir(savepath) if f[:5] == job_number]

                if not job_directory:
                    print("Save directory {}* does not exist".format(os.path.join(savepath, job_number)))
                    return

                for n, f in enumerate(file_match):
                    # print(os.path.join(savepath, job_directory[0]), f)
                    self.move_file_and_split(f, os.path.join(savepath, job_directory[0]))

            except KeyError:
                print("No job number assigned to date search.  Edit config.ini file.")

    def move_willis_art(self):
        # savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
        #                         "In Progress", "01-Web Order Art", "Willis Auto Web Orders")

        savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
                                "In Progress", "01-Web Order Art", "Willis Auto Web Orders TEST")

        file_search = re.compile("WAG[\s\S]*.pdf")
        file_match = (list(filter(file_search.match, self.art_files)))

        if file_match:
            try:
                job_number = self.config['willis'][self.process_date_str[:6]]
                job_directory = [f for f in os.listdir(savepath) if f[:5] == job_number]

                if not job_directory:
                    print("Save directory {}* does not exist".format(os.path.join(savepath, job_number)))
                    return

                for n, f in enumerate(file_match):
                #     # print(os.path.join(savepath, job_directory[0]), f)
                    self.move_file_and_split(f, os.path.join(savepath, job_directory[0]))

            except KeyError:
                print("No job number assigned to date search.  Edit config.ini file.")

    def move_medica_art(self):
        # savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
        #                         "In Progress", "01-Web Order Art", "Medica Monthly Web Orders")
        savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
                                "In Progress", "01-Web Order Art", "Medica Monthly Web Orders TEST")

        file_search = re.compile("MMH[\s\S]*.pdf")
        file_match = (list(filter(file_search.match, self.art_files)))

        if file_match:
            try:
                job_number = self.config['medica'][self.process_date_str[:6]]
                job_directory = [f for f in os.listdir(savepath) if f[:5] == job_number]

                if not job_directory:
                    print("Save directory {}* does not exist".format(os.path.join(savepath, job_number)))
                    return

                for n, f in enumerate(file_match):
                    #     # print(os.path.join(savepath, job_directory[0]), f)
                    self.move_file_and_split(f, os.path.join(savepath, job_directory[0]))

            except KeyError:
                print("No job number assigned to date search.  Edit config.ini file.")

    def move_file_and_split(self, file, path):
        newdir = file.split('_')[0]
        if not os.path.isdir(os.path.join(path, newdir)):
            os.makedirs(os.path.join(path, newdir))

        print("Copying file {0} of {1}: {2}".format(self.count_art_files, len(self.art_files), file))
        shutil.copy(os.path.join(self.processing_directory, file),
                    os.path.join(path, newdir, "w-{0}".format(file)))

        wpdf = PyPDF2.PdfFileReader(os.path.join(path, newdir, "w-{0}".format(file)))
        spdf = PyPDF2.PdfFileWriter()
        data_sheet = PyPDF2.PdfFileWriter()

        for page in range(wpdf.getNumPages()):
            if page == 0:
                data_sheet.addPage(wpdf.getPage(page))
            else:
                spdf.addPage(wpdf.getPage(page))

        with open(os.path.join(path, newdir, file), 'wb') as s:
            spdf.write(s)

        with open(os.path.join(path, newdir, "{0}-data sheet.pdf".format(file[:-4])), 'wb') as s:
            data_sheet.write(s)

        os.remove(os.path.join(path, newdir, "w-{0}".format(file)))
        self.count_art_files += 1


def main():
    g = Globals()
    g.process_date_str = input("Processing date (YYYYMMDD): ")
    g.process_date_files()
    g.get_report_files()
    g.get_art_files()

    g.move_farm_bureau_art()
    g.move_willis_art()
    g.move_medica_art()

    print("File move complete!")
    time.sleep(2)


if __name__ == '__main__':
    main()
