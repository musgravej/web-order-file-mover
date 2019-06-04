import sys
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
        # self.processing_directory = os.path.join("\\\\JTSRV3", "Print Facility", "Job Ticket Feed docs",
        #                                          "WebToPrint")

        self.processing_directory = os.path.join(os.curdir, "webtoprint")

        self.count_art_files = 1

        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

    def get_art_files(self):
        """Sets self.art_files to a set of artwork files for processing date"""
        print("Making list of art files")
        self.art_files = self.dated_files
        self.art_files.difference_update(self.report_files)

    def get_report_files(self):
        """Sets self.report_files to a set of packing slips, work orders, daily reports for processing date"""
        print("Making list of report files")
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

    def get_report_counts(self):
        wo_srch = re.compile("[\d]*_WO_[\d]*.pdf")
        kit_srch = re.compile("[\d]*_WO_split_[\d]*.pdf")

        work_orders = set((list(filter(wo_srch.match, self.report_files))))
        kit_orders = set((list(filter(kit_srch.match, self.report_files))))

        count_string = ""

        print("Getting order counts")

        for order in work_orders:
            if order[:5] == '20403':
                cnt = self.count_pdf_pages(os.path.join(self.processing_directory, order))
                count_string += "Wellmark: {}\r\n".format(cnt)

            if order[:5] == '19404':
                cnt = self.count_pdf_pages(os.path.join(self.processing_directory, order))
                count_string += "Farm Bureau: {}\r\n".format(cnt)

            if order[:5] == '23396':
                cnt = self.count_pdf_pages(os.path.join(self.processing_directory, order))
                count_string += "Medica: {}\r\n".format(cnt)

            if order[:5] == '18241':
                cnt = self.lite_portal_counts(os.path.join(self.processing_directory, order))
                count_string += cnt

        for order in kit_orders:
            if order[:5] == '20403':
                cnt = self.count_pdf_pages(os.path.join(self.processing_directory, order))
                count_string += "Wellmark Kits: {}\r\n".format(cnt)

            if order[:5] == '19404':
                cnt = self.count_pdf_pages(os.path.join(self.processing_directory, order))
                count_string += "Farm Bureau Kits: {}\r\n".format(cnt)

            if order[:5] == '23396':
                cnt = self.count_pdf_pages(os.path.join(self.processing_directory, order))
                count_string += "Medica Kits: {}\r\n".format(cnt)

        pyperclip.copy(count_string)

    def count_pdf_pages(self, pdf_path):
        """Counts the number of pages in pdf_path (full pdf path and file name)"""
        pdf = PyPDF2.PdfFileReader(pdf_path)
        return pdf.getNumPages()

    def lite_portal_counts(self, pdf_path):
        """Reads pdf, gets portal counts, returns string of portal counts"""
        pdf = PyPDF2.PdfFileReader(pdf_path)
        portal_counts = dict()
        count_string = ""
        for page in range(pdf.getNumPages()):
            portal = pdf.getPage(page).extractText().split('\n')[0]
            portal_counts[portal] = portal_counts.get(portal, 0) + 1

        for key, value in portal_counts.items():
            count_string += "{0}: {1}\r\n".format(key, value)

        return count_string

    def move_farm_bureau_art(self):
        savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
                                "In Progress", "01-Web Order Art", "FB Monthly Web Order")

        # savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
        #                         "In Progress", "01-Web Order Art", "FB Monthly Web Order TEST")

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
                    self.move_file_and_split(f, os.path.join(savepath, job_directory[0]))

            except KeyError:
                print("No job number assigned to date search.  Edit config.ini file.")

    def move_willis_art(self):
        savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
                                "In Progress", "01-Web Order Art", "Willis Auto Web Orders")

        # savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
        #                         "In Progress", "01-Web Order Art", "Willis Auto Web Orders TEST")

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
                    self.move_file_and_split(f, os.path.join(savepath, job_directory[0]))

            except KeyError:
                print("No job number assigned to date search.  Edit config.ini file.")

    def move_medica_art(self):
        savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
                                "In Progress", "01-Web Order Art", "Medica Monthly Web Orders")
        # savepath = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
        #                         "In Progress", "01-Web Order Art", "Medica Monthly Web Orders TEST")

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
    if not os.path.isfile('config.ini'):
        print("Error: No config.ini file")
        time.sleep(2)
        sys.exit()

    g = Globals()
    g.process_date_str = input("Processing date (YYYYMMDD): ")
    # g.process_date_str = "20190524"
    g.process_date_files()
    g.get_report_files()
    g.get_art_files()

    g.move_farm_bureau_art()
    g.move_willis_art()
    g.move_medica_art()
    g.get_report_counts()

    print("File move complete!  Report counts copied to clipboard.")
    time.sleep(2)


if __name__ == '__main__':
    main()
