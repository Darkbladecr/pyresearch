from optparse import OptionParser
from Bio import Entrez
from openpyxl import load_workbook
from pubmedHelpers import getPubmedIds, saveWorksheet
import numpy
from tqdm import tqdm
with open('config.txt', 'r') as f:
    Entrez.email = f.readline()

parser = OptionParser()
parser.add_option("-s", "--search", dest="searchTerm", help="search term")
parser.add_option("-i", "--input", dest="input", help="Name of input file")
parser.add_option("-d", "--data", dest="data", help="Name of numpy file")
parser.add_option("-t", "--title", dest="title", help="Name of worksheet")
(options, args) = parser.parse_args()
if len(options.searchTerm) == 0:
	print("Please input a search term with -s or --search")
	exit(1)
print("Search Term: %s") % options.searchTerm

idlist = getPubmedIds(options.searchTerm)
total = len(idlist)
print("Number of articles: %s") % total

data = numpy.load('%s.npy' % options.input)
matched = list()
for record in tqdm(data):
	if record['PMID'] in idlist:
		matched.append(record)

print('Saving to excel')
wb = load_workbook('%s.xlsx' % options.input)
orderedEntries = ['Full Author Names', 'Authors', 'Publication Date', 'Title', 'Publication Type', 'Journal Title', 'Source', 'Language', 'scopusID', 'PMID', 'Citations', 'Citations in Past Year', 'Citations Rate', 'Country of Origin']
saveWorksheet(wb, options.title, matched, options.searchTerm, orderedEntries)

wb.save('%s.xlsx' % options.input)
numpy.save('subsections/%s.npy' % options.title, matched)
