from optparse import OptionParser
from Bio import Entrez
Entrez.email = "smig88@gmail.com"
from openpyxl import Workbook
from pubmedHelpers import outputYearlyData, outputPubDataScopus, outputPubDataPubmed
import numpy
import glob
from tqdm import tqdm

parser = OptionParser()
parser.add_option("-i", "--input", dest="input", help="Name of input file")
parser.add_option("-a", "--all", dest="all", action="store_true", default=False, help="input directory")
(options, args) = parser.parse_args()

def outputOverviewData(file):
	publications = numpy.load('%s.npy' % file)

	yearlyData = outputYearlyData(publications)
	pubTypesData = outputPubDataScopus(publications)
	pubSubsetsData = outputPubDataPubmed(publications)
	wb = Workbook()
	ws = wb.active
	ws.title = 'Overview'
	ws.append(['Search Term:', file])
	ws.append(['Total', sum(yearlyData.values())])
	for pubType, v in pubTypesData.items():
		ws.append([pubType, v])
	ws.append(['Publication Subsets'])
	for subset, v in pubSubsetsData.items():
		ws.append([subset, v])
	ws.append(['Year', file])
	oldestYear = min(yearlyData.keys())
	for year in range(oldestYear, 2017):
		if year not in yearlyData.keys():
			yearlyData[year] = 0
		ws.append([year, yearlyData[year]])

	wb.save('%s.xlsx' % file)

if options.all:
	files = glob.glob("subsections/*.npy")
	files = [f[:-4] for f in files]
	print(files)
	for f in tqdm(files):
		outputOverviewData(f)
else:
	outputOverviewData(options.input)


