from optparse import OptionParser
from Bio import Entrez
from openpyxl import Workbook
from authors import outputAuthors
from journals import outputJournals
from pubmedHelpers import cyberVgamma, outputYearlyData, outputPubDataScopus, outputPubDataPubmed, saveWorksheet
import numpy
import glob
from tqdm import tqdm
with open('config.txt', 'r') as f:
    Entrez.email = f.readline()

parser = OptionParser()
parser.add_option("-i", "--input", dest="input", help="Name of input file")
parser.add_option("-a", "--all", dest="all", action="store_true", default=False, help="input directory")
parser.add_option("-x", "--exclude", dest="exclude", action="store_true", default=False, help="exclude citations")
(options, args) = parser.parse_args()

searchTerms = {
    "Brain Metastases": "Stereotactic AND (radiosurgery OR radiotherapy) AND (brain OR cranial OR cranium) AND (metastasis OR metastases)",
    "Spinal Metastases": "Stereotactic AND (radiosurgery OR radiotherapy) AND (spine OR brainstem OR spinal cord) AND (metastasis OR metastases)",
    "Meningioma": "Stereotactic AND (radiosurgery OR radiotherapy) AND Meningioma",
    "Glioblastoma": "Stereotactic AND (radiosurgery OR radiotherapy) AND (Glioblastoma OR GBM OR high grade glioma OR astrocytoma oligodendroglioma)",
    "AVM": "Stereotactic AND (radiosurgery OR radiotherapy) AND (Ateriovenous malformation OR AVM)",
    "Acoustic Neuroma": "Stereotactic AND (radiosurgery OR radiotherapy) AND (Acoustic neuroma OR vestibular schwanoma)",
    "Trigeminal Neuralgia": "Stereotactic AND (radiosurgery OR radiotherapy) AND (Trigeminal Neuralgia OR Tic Douloureux)",
    "Psychiatry": "Stereotactic AND (radiosurgery OR radiotherapy) AND (depression OR anxiety disorder OR obsessive compulsive OR obsessive compulsion)"
}


def outputOverviewData(file):
	publications = numpy.load('%s.npy' % file)
	if options.exclude:
		searchTerm = searchTerms[file[25:]]
	else:
		searchTerm = searchTerms[file[12:]]

	yearlyData = outputYearlyData(publications)
	pubTypesData = outputPubDataPubmed(publications)
	pubSubsetsData = outputPubDataScopus(publications)

	print("Grabing Subsections")
	subsections = cyberVgamma(publications, searchTerm)

	wb = Workbook()
	ws = wb.active
	ws.title = 'Overview'
	ws.append([
	    'Search Term:',
	    searchTerm,
	    subsections['gamma']['query'],
	    subsections['cyber']['query'],
	    subsections['linac']['query'],
	    subsections['novalis']['query'],
	    subsections['tomo']['query']
	])
	ws.append([
	    'Total',
	    sum(yearlyData.values()),
	    sum(subsections['gamma']['yearlyData'].values()),
	    sum(subsections['gamma']['yearlyData'].values()),
	    sum(subsections['linac']['yearlyData'].values()),
	    sum(subsections['novalis']['yearlyData'].values()),
	    sum(subsections['tomo']['yearlyData'].values())
	])
	for pubType, v in pubTypesData.items():
	    ws.append([
	        pubType,
	        v,
	        subsections['gamma']['pubTypesData'].get(pubType, 0),
	        subsections['cyber']['pubTypesData'].get(pubType, 0),
	        subsections['linac']['pubTypesData'].get(pubType, 0),
	        subsections['novalis']['pubTypesData'].get(pubType, 0),
	        subsections['tomo']['pubTypesData'].get(pubType, 0)
	    ])
	ws.append(['Publication Subsets'])
	for subset, v in pubSubsetsData.items():
	    ws.append([
	        subset,
	        v,
	        subsections['gamma']['pubSubsetsData'].get(subset, 0),
	        subsections['cyber']['pubSubsetsData'].get(subset, 0),
	        subsections['linac']['pubSubsetsData'].get(subset, 0),
	        subsections['novalis']['pubSubsetsData'].get(subset, 0),
	        subsections['tomo']['pubSubsetsData'].get(subset, 0)
	    ])
	ws.append([
	    'Year',
	    searchTerm,
	    subsections['gamma']['query'],
	    subsections['cyber']['query'],
	    subsections['linac']['query'],
	    subsections['novalis']['query'],
	    subsections['tomo']['query']
	])
	oldestYear = min(yearlyData.keys())
	for year in range(oldestYear, 2017):
	    if year not in yearlyData.keys():
	        yearlyData[year] = 0
	    if year not in subsections['gamma']['yearlyData'].keys():
	        subsections['gamma']['yearlyData'][year] = 0
	    if year not in subsections['cyber']['yearlyData'].keys():
	        subsections['cyber']['yearlyData'][year] = 0
	    if year not in subsections['linac']['yearlyData'].keys():
	        subsections['linac']['yearlyData'][year] = 0
	    if year not in subsections['novalis']['yearlyData'].keys():
	        subsections['novalis']['yearlyData'][year] = 0
	    if year not in subsections['tomo']['yearlyData'].keys():
	        subsections['tomo']['yearlyData'][year] = 0
	    ws.append([
	        year,
	        yearlyData[year],
	        subsections['gamma']['yearlyData'][year],
	        subsections['cyber']['yearlyData'][year],
	        subsections['linac']['yearlyData'][year],
	        subsections['novalis']['yearlyData'][year],
	        subsections['tomo']['yearlyData'][year]
	    ])

	saveWorksheet(wb, 'Pubmed Stats', publications, searchTerm, orderedEntries=True, autoFilter=True)

	authorData = outputAuthors(publications)
	saveWorksheet(wb, 'Authors', authorData, searchTerm=searchTerm, autoFilter=True)

	journalData = outputJournals(publications)
	saveWorksheet(wb, 'Journals', journalData)

	saveWorksheet(wb, 'Gamma Knife', subsections['gamma']['records'], subsections['gamma']['query'], orderedEntries=True, autoFilter=True)
	saveWorksheet(wb, 'CyberKnife', subsections['cyber']['records'], subsections['cyber']['query'], orderedEntries=True, autoFilter=True)
	saveWorksheet(wb, 'linac', subsections['linac']['records'], subsections['linac']['query'], orderedEntries=True, autoFilter=True)
	saveWorksheet(wb, 'Novalis', subsections['novalis']['records'], subsections['novalis']['query'], orderedEntries=True, autoFilter=True)
	saveWorksheet(wb, 'TomoTherapy', subsections['tomo']['records'], subsections['tomo']['query'], orderedEntries=True, autoFilter=True)
	wb.save('%s.xlsx' % file)

if options.exclude:
    directory = 'subsections-self_exclude'
else:
    directory = 'subsections'
if options.all:
	files = glob.glob("%s/*.npy" % directory)
	files = [file[:-4] for file in files]
	print(files)
	for f in tqdm(files):
		outputOverviewData(f)
else:
	outputOverviewData(options.input)
