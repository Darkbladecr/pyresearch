from optparse import OptionParser
from Bio import Entrez
from openpyxl import load_workbook
from pubmedHelpers import getPubmedIds, saveWorksheet
import numpy
from operator import itemgetter
from tqdm import tqdm
with open('config.txt', 'r') as f:
    Entrez.email = f.readline()

parser = OptionParser()
parser.add_option("-s", "--search", dest="searchTerm", help="search term")
parser.add_option("-i", "--input", dest="input", help="Name of input file")
parser.add_option("-d", "--data", dest="data", help="Name of numpy file")
parser.add_option("-t", "--title", dest="title", help="Name of worksheet")
parser.add_option("-x", "--exclude", dest="exclude", action="store_true", default=False, help="exclude citations")
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
matched = sorted(matched, key=itemgetter('Citations Rate', 'Citations'), reverse=True)

print('Saving to excel')
wb = load_workbook('%s.xlsx' % options.input)
saveWorksheet(wb, options.title, matched, options.searchTerm, orderedEntries=True, autoFilter=True)

wb.save('%s.xlsx' % options.input)
if options.exclude:
    directory = 'subsections-self_exclude'
else:
    directory = 'subsections'
numpy.save('%s/%s.npy' % (directory, options.title), matched)
