import os
import glob
import tkinter as tk #required download package
from tkinter import filedialog
from PyPDF2 import PdfMerger #required download package

#create an instance of the pick path dialog where PDFs are
in_application_window = tk.Tk()
pdf_in_dir = filedialog.askdirectory(parent=in_application_window,
                                 initialdir=os.getcwd(),
                                 title="Please select folder that contains the PDFs to merge:")

os.chdir(pdf_in_dir) #swap this with arg from modal
pdfs_to_merge = glob.glob("*.pdf") #get the PDF from the dir
in_application_window.destroy() #end the dialog

#create an instance of the Mergers
merger = PdfMerger()

#merge the files
for pdf in pdfs_to_merge:
    merger.append(pdf)

# Create an instance of the pick path dialog where the merged PDF needs to go
out_application_window = tk.Tk()
pdf_out_dir = filedialog.askdirectory(parent=out_application_window,
                                 initialdir=os.getcwd(),
                                 title="Please select folder where you want to put the merged file:")
#write out the merged file
merger.write(os.path.join(pdf_out_dir, "Merged.pdf"))
merger.close()

#end the dialog
out_application_window.destroy() # End the dialog

