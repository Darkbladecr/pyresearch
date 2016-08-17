from optparse import OptionParser
from Bio import Entrez
from openpyxl import Workbook
from pubmedHelpers import cyberVgamma, outputYearlyData, outputPubDataScopus, outputPubDataPubmed, saveWorksheet
import numpy
import glob
from tqdm import tqdm
with open('config.txt', 'r') as f:
    Entrez.email = f.readline()

parser = OptionParser()
parser.add_option("-i", "--input", dest="input", help="Name of input file")
parser.add_option("-a", "--all", dest="all", action="store_true", default=False, help="input directory")
(options, args) = parser.parse_args()

searchTerms = {
    "Brain Metastases": "Stereotactic AND (radiosurgery OR radiotherapy) AND (brain OR cranial OR cranium) AND (metastasis OR metastases)",
    "Spinal Metastases": "Stereotactic AND (radiosurgery OR radiotherapy) AND (spine OR brainstem OR spinal cord) AND (metastasis OR metastases)",
    "Meningioma": "Stereotactic AND (radiosurgery OR radiotherapy) AND Meningioma",
    "Glioblastoma": "Stereotactic AND (radiosurgery OR radiotherapy) AND Glioblastoma",
    "AVM": "Stereotactic AND (radiosurgery OR radiotherapy) AND (Ateriovenous malformation OR AVM)",
    "Acoustic Neuroma": "Stereotactic AND (radiosurgery OR radiotherapy) AND (Acoustic neuroma OR vestibular schwanoma)",
    "Trigeminal Neuralgia": "Stereotactic AND (radiosurgery OR radiotherapy) AND (Trigeminal Neuralgia OR Tic Douloureux)",
    "Psychiatry": "Stereotactic AND (radiosurgery OR radiotherapy) AND (Epilepsy OR depression OR anxiety disorder OR obsessive compulsive OR obsessive compulsion)"
}


def outputOverviewData(file):
	publications = numpy.load('%s.npy' % file)
	searchTerm = searchTerms[file[12:]]

	subsections = cyberVgamma(publications, searchTerm)
	gammaSearch = "%s AND (gamma knife OR gammaknife)" % searchTerm
	cyberSearch = "%s AND (Cyberknife OR cyber knife OR LINAC)" % searchTerm
	cyberPubs = subsections['cyber']
	gammaPubs = subsections['gamma']
	numpy.save('cyber_data.npy', cyberPubs)
	numpy.save('gamma_data.npy', gammaPubs)

	yearlyData = outputYearlyData(publications)
	yearlyDataGamma = outputYearlyData(gammaPubs)
	yearlyDataCyber = outputYearlyData(cyberPubs)

	pubTypesData = outputPubDataScopus(publications)
	pubTypesDataGamma = outputPubDataScopus(gammaPubs)
	pubTypesDataCyber = outputPubDataScopus(cyberPubs)
	pubSubsetsData = outputPubDataPubmed(publications)
	pubSubsetsDataGamma = outputPubDataPubmed(gammaPubs)
	pubSubsetsDataCyber = outputPubDataPubmed(cyberPubs)

	wb = Workbook()
	ws = wb.active
	ws.title = 'Overview'
	ws.append(['Search Term:', searchTerm, gammaSearch, cyberSearch])
	ws.append(['Total', sum(yearlyData.values()), sum(yearlyDataGamma.values()), sum(yearlyDataCyber.values())])
	for pubType, v in pubTypesData.items():
		ws.append([pubType, v, pubTypesDataGamma.get(pubType, 0), pubTypesDataCyber.get(pubType, 0)])
	ws.append(['Publication Subsets'])
	for subset, v in pubSubsetsData.items():
		ws.append([subset, v, pubSubsetsDataGamma.get(subset, 0), pubSubsetsDataCyber.get(subset, 0)])
	ws.append(['Year', searchTerm, gammaSearch, cyberSearch])
	oldestYear = min(yearlyData.keys())
	for year in range(oldestYear, 2017):
		if year not in yearlyData.keys():
			yearlyData[year] = 0
		if year not in yearlyDataGamma.keys():
			yearlyDataGamma[year] = 0
		if year not in yearlyDataCyber.keys():
			yearlyDataCyber[year] = 0
		ws.append([year, yearlyData[year], yearlyDataGamma[year], yearlyDataCyber[year]])
	saveWorksheet(wb, file[12:], publications, searchTerm, orderedEntries=True)
	saveWorksheet(wb, 'Gamma Knife', gammaPubs, gammaSearch, orderedEntries=True)
	saveWorksheet(wb, 'Cyberknife', cyberPubs, cyberSearch, orderedEntries=True)
	wb.save('%s.xlsx' % file)

if options.all:
	files = glob.glob("subsections/*.npy")
	files = [file[:-4] for file in files]
	print(files)
	for f in tqdm(files):
		outputOverviewData(f)
else:
	outputOverviewData(options.input)
