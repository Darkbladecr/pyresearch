from openpyxl import load_workbook, Workbook
from openpyxl.chart import LineChart, Reference, Series
from datetime import datetime
from sets import Set
from collections import OrderedDict
from operator import itemgetter
from Bio import Entrez
from iso639 import languages
from tqdm import tqdm


def excel2Dict(filename):
	wb = load_workbook(filename=filename, read_only=True)
	ws = wb.active
	records = []
	count = 0
	keys = []
	for row in ws.rows:
		temp = []
		for cell in row:
			if count == 0:
				keys.append(str(cell.value))
			else:
				temp.append(str(cell.value))
		if count > 0:
			records.append(dict(zip(keys, temp)))
		else:
			count += 1
	return records


def saveExcel(filename, title, data):
	wb = Workbook()
	ws = wb.active
	ws.title = title

	count = 0
	for item in data:
		if count == 0:
			ws.append(item.keys())
			count += 1
		ws.append(item.values())
	date = datetime.now().strftime('%Y-%m-%d')
	name = "%s-%s.xlsx" % (filename, date)
	wb.save(name)


def saveWorksheet(wb, title, data, searchTerm=None, orderedEntries=False, autoFilter=False):
	if orderedEntries is True:
		orderedEntries = ['Full Author Names', 'Authors', 'Publication Date', 'Title', 'Publication Type', 'Journal Title', 'Source', 'Language', 'scopusID', 'PMID', 'Citations', 'Citations in Past Year', 'Citations Rate', 'Country of Origin']
	ws = wb.create_sheet()
	ws.title = title
	count = 0
	for record in data:
		if orderedEntries:
			record = OrderedDict((k, record[k]) for k in orderedEntries)
		if count == 0:
			if searchTerm is not None:
				ws.append([searchTerm])
			ws.append(record.keys())
			count += 1
		output = list()
		for val in record.values():
			if isinstance(val, list):
				# dump = json.dumps(val)
				dump = ", ".join(val)
				output.append(dump)
			else:
				output.append(val)
		ws.append(output)
	if autoFilter and orderedEntries:
		ws.auto_filter.ref = "A2:N%d" % (len(data) + 1)
		for row in range(3, ws.max_row):
			cell = ws.cell(row=row, column=13)
			cell.number_format = "0.0"
	elif autoFilter:
		ws.auto_filter.ref = "A2:D%d" % (len(data) + 1)


def distinctSet(records, title):
	data = Set([])
	for record in records:
		if isinstance(record[title], list):
			for item in record[title]:
				data.add(item)
		else:
			data.add(record[title])
	return data


def languageParse(record):
	if len(record['LA']) > 1:
		return [languages.get(part2b=l.lower()).name for l in record['LA']]
	else:
		langCode = record['LA'][0].lower()
		return languages.get(part2b=langCode).name


def outputYearlyData(records):
	output = dict()
	for record in records:
		year = record['Publication Date'].year
		if year in output.keys():
			output[year] += 1
		else:
			output[year] = 1
	return output


def outputYearlyCitationData(records):
	output = dict()
	for record in records:
		year = record['Publication Date'].year
		if year in output.keys():
			output[year] += record['Citations']
		else:
			output[year] = record['Citations']
	return output
PToutput = ['ENGLISH ABSTRACT', 'Meta-Analysis', 'Controlled Clinical Trial', 'LETTER', 'Clinical Trial, Phase III', 'Review', "Research Support, U.S. Gov't, P.H.S.", 'Guideline', 'Interview', "Research Support, Non-U.S. Gov't", 'Consensus Development Conference', 'Lectures', "Research Support, U.S. Gov't, Non-P.H.S.", 'Published Erratum', 'Clinical Study', 'Overall', 'REVIEW', 'Letter', 'Observational Study', 'Comparative Study', 'Clinical Trial, Phase I', 'Comment', 'Multicenter Study', 'Validation Studies', 'Journal Article', 'JOURNAL ARTICLE', 'Historical Article', 'Evaluation Studies', 'Research Support, N.I.H., Extramural', 'Editorial', 'Retracted Publication', 'CASE REPORTS', 'Classical Article', 'Technical Report', 'Randomized Controlled Trial', 'Video-Audio Media', 'English Abstract', 'News', 'Portraits', 'Clinical Trial, Phase II', 'Research Support, N.I.H., Intramural', 'Introductory Journal Article', 'Congresses', 'Case Reports', 'Practice Guideline', 'Clinical Trial', 'Biography']
PTwant = ["Journal Article", "Letter", "Editorial", "Review", "Meta-Analysis", "Research Support, Non-U.S. Gov't", "Research Support, U.S. Gov't, P.H.S.", "Research Support, U.S. Gov't, Non-P.H.S.", "Research Support, N.I.H., Extramural", "Research Support, N.I.H., Intramural", "Clinical Study", "Observational Study", "Comparative Study", "Multicenter Study", "Validation Studies", "Controlled Clinical Trial", "Clinical Trial", "Clinical Trial, Phase I", "Clinical Trial, Phase II", "Clinical Trial, Phase III", "Randomized Controlled Trial", "Guideline", "Practice Guideline", "Lectures", "Case Reports", "Historical Article", "Evaluation Studies"]
PTdel = ['ENGLISH ABSTRACT', 'LETTER', 'Interview', 'Published Erratum', 'Overall', 'REVIEW', 'Comment', 'JOURNAL ARTICLE', 'CASE REPORTS', 'Classical Article', 'Technical Report', 'Video-Audio Media', 'English Abstract', 'News', 'Portraits', 'Introductory Journal Article']


def outputPubDataPubmed(records):
	output = dict()
	renamer = {'LETTER': 'Letter', 'REVIEW': 'Review', 'JOURNAL ARTICLE': 'Journal Article', 'CASE REPORTS': 'Case Reports'}
	for record in records:
		pubTypes = record['Publication Subset']
		pubTypesRenamed = [renamer.get(pub, pub) for pub in pubTypes]
		for pubType in pubTypesRenamed:
			if pubType in output.keys():
				output[pubType] += 1
			else:
				output[pubType] = 1
	output = {p: output.get(p, 0) for p in PTwant}
	return OrderedDict((k, output[k]) for k in PTwant)

PTscopus = ["Article", "Letter", "Editorial", "Review", "Conference Paper", "Chapter", "Note", "Short Survey"]


def outputPubDataScopus(records):
	output = OrderedDict((k, 0) for k in PTscopus)
	for r in records:
		try:
			pubType = r['Publication Type']
			output[pubType] += 1
		except:
			pass
	return output


def getPubmedIds(term, start=0):
	handle = Entrez.esearch(db="pubmed", retstart=start, retmax=20000, term=term)
	record = Entrez.read(handle)
	return record["IdList"]


def parsePubmed(records, total, array, pmids):
	for r in tqdm(records):
		if 'TI' in r:
			pub = dict()
			pub['Full Author Names'] = r.get('FAU', 'Unknown')
			pub['Authors'] = r.get('AU', 'Unknown')
			try:
				pub['Publication Date'] = datetime.strptime(r['EDAT'][:-6], "%Y/%M/%d").date()
			except ValueError:
				try:
					pub['Publication Date'] = datetime.strptime(r['EDAT'], "%Y/%M/%d").date()
				except ValueError:
					print("Date parsing error: %s") % r['EDAT']
					pub['Publication Date'] = r['EDAT']
			pub['Title'] = r['TI']
			pub['Publication Type'] = r['PT']
			pub['Journal Title'] = r.get('JT', 'Unknown')
			pub['Source'] = r.get('SO', 'Unknown')
			pub['Language'] = languageParse(r)
			pub['PMID'] = r['PMID']
			pub['Citations'] = 0
			pub['Citations in Past Year'] = 0
			pub['Citations Rate'] = 0
			array.append(pub)
			pmids.append(r['PMID'])


def pubmedData(wb, records, searchTerm, suffix):
	yearlyData = outputYearlyData(records)
	pubTypesData = outputPubDataPubmed(records)
	pubSubsetsData = outputPubDataScopus(records)

	print("Grabing Subsections")
	subsections = {}
	subsections['all'] = {
	    'query': searchTerm,
	    'records': records,
	    'yearlyData': yearlyData,
	    'pubSubsetsData': pubSubsetsData,
	    'pubTypesData': pubTypesData
	}

	ws = wb.active
	ws.title = 'Overview'
	ws.append([
	    'Search Term:',
	    searchTerm
	])
	ws.append([
	    'Total',
	    sum(yearlyData.values())
	])
	ws.append(['Publication Scopus Subsets'])
	rowNum = 3
	for subset, v in pubSubsetsData.items():
	    ws.append([
	        subset,
	        v
	    ])
	    rowNum += 1
	ws.append(['Publication Pubmed Subsets'])
	rowNum += 1
	for pubType, v in pubTypesData.items():
	    ws.append([
	        pubType,
	        v
	    ])
	    rowNum += 1
	ws.append([
	    'Year',
	    searchTerm
	])
	rowNum += 1
	chartStart = rowNum + 1
	oldestYear = min(yearlyData.keys())
	for year in range(oldestYear, 2017):
	    ws.append([
	        year,
	        yearlyData.get(year, 0)
	    ])
	    rowNum += 1

	c1 = LineChart()
	c1.width = 30
	c1.height = 15
	dates = Reference(ws, min_row=chartStart, min_col=1, max_col=1, max_row=rowNum)
	s1 = Reference(ws, min_row=chartStart - 1, min_col=2, max_col=2, max_row=rowNum)
	c1.series.append(Series(s1, title_from_data=True))
	c1.set_categories(dates)
	c1.title = "Articles per year for search: %s" % searchTerm
	c1.y_axis.title = 'Number of Articles'
	c1.x_axis.title = 'Year'
	ws.add_chart(c1, "A%d" % (rowNum + 5))

	saveWorksheet(wb, 'Pubmed Stats', records, searchTerm, orderedEntries=True, autoFilter=True)
	return subsections
