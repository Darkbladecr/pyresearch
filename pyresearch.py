from optparse import OptionParser
from Bio import Entrez, Medline
with open('config.txt', 'r') as f:
    Entrez.email = f.readline()
from pubmedHelpers import getPubmedIds, parsePubmed, cyberVgamma, outputYearlyData, outputPubDataScopus, outputPubDataPubmed, saveWorksheet
from authors import outputAuthors
from journals import outputJournals
from scopusAPI import searchScopus, parseScopus, citeMetadata2
import numpy
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from datetime import datetime
date = datetime.now().strftime('%Y-%m-%d')
from tqdm import tqdm

parser = OptionParser()
parser.add_option("-s", "--search", dest="searchTerm", help="search term")
parser.add_option("-i", "--input", dest="input", help="input file")
(options, args) = parser.parse_args()
if len(options.searchTerm) == 0:
	print("Please input a search term with -s or --search")
	exit(1)
print("Search Term: %s") % options.searchTerm

if options.input:
	scopusPublications = numpy.load(options.input)
else:
	idlist = getPubmedIds(options.searchTerm, 0)
	print("Number of articles: %s") % len(idlist)
	handle = Entrez.efetch(db="pubmed", id=idlist, rettype="medline", retmode="text")
	records = Medline.parse(handle)

	count = 0
	pmids = list()
	publications = list()
	print("Pubmed gather started")
	parsePubmed(records, len(idlist), publications, pmids)
	# numpy.save('pubmed_data.npy', publications)

	scopusPublications = list()
	count = 0
	print("Scopus gather started")
	with tqdm(total=len(publications)) as pbar:
		while count < len(publications):
			query = ""
			for id in pmids[count:count+200]:
				query += " OR PMID(%s)" % id
			query = query[4:]
			content = searchScopus(query, 0)
			if int(content['opensearch:itemsPerPage']) == 1:
				parseScopus(content['entry'], scopusPublications)
			else:
				for r in content['entry']:
					parseScopus(r, scopusPublications)
			count += 200
			# print("Articles gathered: %d") % len(scopusPublications)
			pbar.update(200)
		# numpy.save('scopus_data.npy', scopusPublications)

	print('Merging data from Pubmed to Scopus records')
	for i, r in tqdm(enumerate(scopusPublications)):
		for pub in publications:
			if pub['PMID'] == r['PMID']:
				temp = r
				temp['Full Author Names'] = pub['Full Author Names']
				temp['Authors'] = pub['Authors']
				temp['Source'] = pub['Source']
				temp['Language'] = pub['Language']
				temp['Publication Date'] = pub['Publication Date']
				temp['Publication Subset'] = pub['Publication Type']
				scopusPublications[i] = temp
	# numpy.save('merged_data.npy', scopusPublications)

scopusids = [v['scopusID'] for v in scopusPublications]
scopusCites = dict()
count = 0
print('Gathering Articles citation data')
with tqdm(total=len(scopusPublications)) as pbar:
	while count < len(scopusPublications):
		query = ",".join(scopusids[count:count+25])
		content = citeMetadata2(query)
		scopusCites.update(content)
		count += 25
		pbar.update(25)
		# print("Article citation metadata gathered: %d") % len(scopusCites.keys())

for i, r in tqdm(enumerate(scopusPublications)):
	scopusPublications[i]['Authors'] = scopusCites[r['scopusID']]['scopusAuthors']
	try:
		scopusPublications[i]['Citations'] = scopusCites[r['scopusID']]['Citations']
		scopusPublications[i]['Citations in Past Year'] = scopusCites[r['scopusID']]['Citations in Past Year']
		scopusPublications[i]['Citations Rate'] = scopusCites[r['scopusID']]['Citations Rate']
	except TypeError:
		print("Cite metadata error on article %d") % i

numpy.save('data-%s.npy' % date, scopusPublications)

print("Grabing Gamma knife vs Cyberknife data")
subsections = cyberVgamma(scopusPublications, options.searchTerm)
gammaSearch = "%s AND (gamma knife OR gammaknife)" % options.searchTerm
cyberSearch = "%s AND (Cyberknife OR cyber knife OR LINAC)" % options.searchTerm
cyberPubs = subsections['cyber']
gammaPubs = subsections['gamma']
numpy.save('cyber_data.npy', cyberPubs)
numpy.save('gamma_data.npy', gammaPubs)

print("Gathering Yearly Data")
yearlyData = outputYearlyData(scopusPublications)
yearlyDataGamma = outputYearlyData(gammaPubs)
yearlyDataCyber = outputYearlyData(cyberPubs)

print("Gathering Publication Types Data")
pubTypesData = outputPubDataScopus(scopusPublications)
pubTypesDataGamma = outputPubDataScopus(gammaPubs)
pubTypesDataCyber = outputPubDataScopus(cyberPubs)
pubSubsetsData = outputPubDataPubmed(scopusPublications)
pubSubsetsDataGamma = outputPubDataPubmed(gammaPubs)
pubSubsetsDataCyber = outputPubDataPubmed(cyberPubs)

wb = Workbook()
ws = wb.active
ws.title = 'Overview'
ws.append(['Search Term:', options.searchTerm, gammaSearch, cyberSearch])
ws.append(['Total', sum(yearlyData.values()), sum(yearlyDataGamma.values()), sum(yearlyDataCyber.values())])
rowNum = 3
for pubType, v in pubTypesData.items():
	ws.append([pubType, v, pubTypesDataGamma.get(pubType, 0), pubTypesDataCyber.get(pubType, 0)])
	rowNum += 1
ws.append(['Publication Subsets'])
rowNum += 1
for subset, v in pubSubsetsData.items():
	ws.append([subset, v, pubSubsetsDataGamma.get(subset, 0), pubSubsetsDataCyber.get(subset, 0)])
	rowNum += 1
ws.append(['Year', options.searchTerm, gammaSearch, cyberSearch])
rowNum += 1
chartStart = rowNum
oldestYear = min(yearlyData.keys())
for year in range(oldestYear, 2016):
	if year not in yearlyData.keys():
		yearlyData[year] = 0
	if year not in yearlyDataGamma.keys():
		yearlyDataGamma[year] = 0
	if year not in yearlyDataCyber.keys():
		yearlyDataCyber[year] = 0
	ws.append([year, yearlyData[year], yearlyDataGamma[year], yearlyDataCyber[year]])
	rowNum += 1

c1 = LineChart()
dates = Reference(ws, min_row=chartStart, min_col=1, max_col=1, max_row=rowNum)
s1 = Reference(ws, min_row=chartStart-1, min_col=2, max_col=2, max_row=rowNum-1)
s2 = Reference(ws, min_row=chartStart-1, min_col=3, max_col=3, max_row=rowNum-1)
s3 = Reference(ws, min_row=chartStart-1, min_col=4, max_col=4, max_row=rowNum-1)
c1.series.append(Series(s1, title_from_data=True))
c1.series.append(Series(s2, title_from_data=True))
c1.series.append(Series(s3, title_from_data=True))
c1.set_categories(dates)
c1.title = "Articles per year for search: %s" % options.searchTerm
c1.y_axis.title = 'Number of Articles'
c1.x_axis.title = 'Year'
ws.add_chart(c1, "A%d" % (rowNum+5))

orderedEntries = ['Full Author Names', 'Authors', 'Publication Date', 'Title', 'Publication Type', 'Journal Title', 'Source', 'Language', 'scopusID', 'PMID', 'Citations', 'Citations in Past Year', 'Citations Rate', 'Country of Origin']
saveWorksheet(wb, 'Pubmed Stats', scopusPublications, options.searchTerm, orderedEntries)

authorData = outputAuthors(scopusPublications)
saveWorksheet(wb, 'Authors', authorData)

journalData = outputJournals(scopusPublications)
saveWorksheet(wb, 'Journals', journalData)

saveWorksheet(wb, 'Gamma Knife', gammaPubs, gammaSearch, orderedEntries)
saveWorksheet(wb, 'Cyberknife', cyberPubs, cyberSearch, orderedEntries)

wb.save('data-%s.xlsx' % date)
