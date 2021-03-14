import sys, os
import win32com.client


wdFormatPDF = 17


def convert(_in, _out):
	in_file = os.path.abspath(_in)
	out_file = os.path.abspath(_out)
	word = win32com.client.Dispatch('Word.Application')
	doc = word.Documents.Open(in_file)
	doc.SaveAs(out_file, FileFormat = wdFormatPDF)
	doc.Close()
	word.Quit()

destination = sys.argv[1]
for file in os.listdir(destination):
	convert(destination +'\\' + file, destination + '\\' + file+'.pdf')

