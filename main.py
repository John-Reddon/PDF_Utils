# pyuic5 pdf_utils.ui -o pdf_utils.py
# pyinstaller -F --specpath \dist --clean -n PDF_Utils --noconsole --icon icone.ico main.py
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from pdf_utils import Ui_MainWindow

import sys
import glob
import PyPDF2
import os
import threading
import logging
import subprocess

class PDF_utils(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        # Build the UI
        super(PDF_utils, self).__init__(parent)
        self.setupUi(self)

        # Folders :
        self.user = os.path.join(os.path.expanduser('~'))
        self.root = os.getcwd()
        self.output_folder = self.root + os.sep + "PDFs"
        try:
            os.mkdir(self.output_folder)
        except FileExistsError:
            pass

        # Logger
        logging.basicConfig(filename=self.output_folder + os.sep + 'app.log', filemode='a',
                            format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)
        logging.info('New app launch!')

        # Ui init
        self.merge_rv_frame.hide()
        self.merge_docs_frame.hide()
        self.link_buttons()

        # variables
        self.path_of_pdfs = None
        self.even_pages_path = None
        self.odd_pages_path = None
        self.multi_merge = False
        self.odd_even_merge = False
        self.document_path = None
        self.name_is_ok = False
        self.files_are_given = False



        # Some more  UI options
        self.show()

    @staticmethod
    def startfile(filename):
        try:
            os.startfile(filename)
        except:
            subprocess.Popen(['xdg-open', filename])


    def link_buttons(self):
        self.path_to_pdfs_button.clicked.connect(lambda: self.get_folder_of_pdf())
        self.multi_merge_button.toggled.connect(lambda: self.enter_multi_merge_mode())
        self.document_name.textChanged.connect(lambda : self.name_is_input())

        self.even_pages_button.clicked.connect(lambda: self.get_even())
        self.odd_pages_button.clicked.connect(lambda : self.get_odd())
        self.odd_even_merge_button.toggled.connect(lambda : self.enter_rv_merge_mode())

        self.MergeButton.clicked.connect(lambda: self.merging())

    def enable_button(self):
        if self.name_is_ok and self.files_are_given:
            self.MergeButton.setEnabled(True)
        else:
            self.MergeButton.setEnabled(False)

    def name_is_input(self):
        self.name_is_ok = False
        if self.document_name.text() != "" and "." not in self.document_name.text():
            self.document_path = self.output_folder + os.sep + self.document_name.text() + ".pdf"
            if os.path.exists(self.document_path):
                self.name_is_ok = False
            else:
                self.name_is_ok = True
        self.enable_button()

    # Merging multiple PDFS
    def enter_multi_merge_mode(self):
        self.re_init_pdfs()
        self.MergeButton.setEnabled(False)
        self.multi_merge = True
        self.odd_even_merge = False
        self.document_name.setText("")
        self.document_path = None
        self.document_name.setEnabled(True)
        self.files_are_given = False
        self.name_is_ok = False

    def re_init_pdfs(self):
        self.pdf_list.clear()
        self.path_of_pdfs = None
        self.document_name.setText("")

    def get_folder_of_pdf(self):
        self.re_init_pdfs()
        self.files_are_given = False
        folder_of_pdfs = str(QFileDialog.getExistingDirectory(self, "Dossier contenant les PDFs à fusionner",
                                                                   directory=self.user))
        if folder_of_pdfs:
            self.user = folder_of_pdfs
            self.path_of_pdfs = glob.glob(folder_of_pdfs + os.sep + "*.pdf")
            if self.path_of_pdfs != []:
                self.path_of_pdfs.sort()
                for idx, path in enumerate(self.path_of_pdfs):
                    self.pdf_list.insertItem(idx, os.path.basename(path))
                    logging.info(f"Adding in multi merge mode the file {path}")
                self.files_are_given = True

        self.enable_button()

    def multi_merge_function(self, readers, total_pages):
        try:
            writer = PyPDF2.PdfWriter()
            processed_pages = 0
            for reader in readers:
                for page in range(len(reader.pages)):
                    processed_pages += 1
                    writer.add_page(reader.pages[page])
                    self.progressBar.setValue(int(processed_pages*100/total_pages))

            result = open(self.document_path, "wb")
            writer.write(result)
            result.close()
        except Exception as e:
            logging.error(f"The was an error trying to multi merge: {e}")


    def combine_rv_function(self, odd, even, total_pages):
        bigger_doc = max(len(odd.pages), len(even.pages))
        processed_pages = 0
        try:
            writer = PyPDF2.PdfWriter()
            for page in range(bigger_doc):
                try:
                    writer.add_page(odd.pages[page])
                    processed_pages += 1
                    self.progressBar.setValue(int(processed_pages*100/total_pages))
                except Exception as e:
                    logging.error(f"The was an error trying to combine RV merge: {e}")
                    pass
                try:
                    writer.add_page(even.pages[page])
                    processed_pages += 1
                    self.progressBar.setValue(int(processed_pages*100/total_pages))
                except Exception as e:
                    logging.error(f"The was an error trying to combine RV merge: {e}")
                    pass

            result = open(self.document_path, "wb")
            writer.write(result)
            result.close()
        except Exception as e:
            logging.error(f"The was an error trying to combine RV merge: {e}")

    # R/V mode part
    def re_init_rv(self):
        self.even_pages_path = None
        self.even_pages_text.setText("")
        self.odd_pages_path = None
        self.odd_page_text.setText("")
        self.document_name.setText("")

    def enter_rv_merge_mode(self):
        self.re_init_rv()
        self.MergeButton.setEnabled(False)
        self.multi_merge = False
        self.odd_even_merge = True
        self.document_name.setText("")
        self.document_path = None
        self.document_name.setEnabled(True)
        self.files_are_given = False
        self.name_is_ok = False

    def get_even(self):
        self.even_pages_path = str(QFileDialog.getOpenFileName(self,
                                                               "Document contenant les pages paires (2, 4, ...)",
                                                               directory=self.user)[0])
        if self.even_pages_path:
            self.user = os.path.dirname(self.even_pages_path)
            self.even_pages_text.setText(os.path.basename(self.even_pages_path))
            logging.info(f"Added even pages file, {self.even_pages_path}")
            if self.odd_pages_path:
                self.files_are_given = True
        self.enable_button()

    def get_odd(self):
        self.odd_pages_path = str(QFileDialog.getOpenFileName(self,
                                                              "Document contenant les pages impaires (1, 3, ...)",
                                                              directory=self.user)[0])
        if self.odd_pages_path:
            self.user = os.path.dirname(self.odd_pages_path)
            self.odd_page_text.setText(os.path.basename(self.odd_pages_path))
            logging.info(f"Added even pages file, {self.odd_pages_path}")
            if self.even_pages_path:
                self.files_are_given = True
        self.enable_button()

    # Merging
    def merging(self):
        self.MergeButton.setEnabled(False)
        try:
            if self.multi_merge:
                logging.info(f"Asked to merge files into this pdf: {self.document_path}.")
                PDFs = []
                pages = 0
                for path in self.path_of_pdfs:
                    reader = PyPDF2.PdfReader(path)
                    PDFs.append(reader)
                    pages += len(reader.pages)

                thr = threading.Thread(target=self.multi_merge_function, args=(PDFs, pages,))
                thr.start()
                thr.join()

                self.startfile(self.output_folder)
                self.progressBar.setValue(0)
                self.re_init_pdfs()

            if self.odd_even_merge:
                logging.info(f"Asked to combine RV into this pdf: {self.document_path}.")

                odd_reader = PyPDF2.PdfReader(self.odd_pages_path)
                even_reader = PyPDF2.PdfReader(self.even_pages_path)
                pages = len(odd_reader.pages) + len(even_reader.pages)

                thr = threading.Thread(target=self.combine_rv_function, args=(odd_reader, even_reader, pages,))
                thr.start()
                thr.join()

                self.startfile(self.output_folder)
                self.progressBar.setValue(0)
                self.re_init_rv()
        except Exception as e:
            logging.error(f"The was an error trying to merge: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ui = PDF_utils()
    sys.exit(app.exec_())
