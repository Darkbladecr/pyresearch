from optparse import OptionParser
from Bio import Entrez, Medline
from pubmedHelpers import getPubmedIds, parsePubmed, cyberVgamma, outputYearlyData, outputPubDataScopus, outputPubDataPubmed, saveWorksheet
from authors import outputAuthors, outputCountries
from journals import outputJournals
from scopusAPI import searchScopus, parseScopus, citeMetadata
import numpy
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from datetime import datetime
from operator import itemgetter
from tqdm import tqdm
date = datetime.now().strftime('%Y-%m-%d')
with open('config.txt', 'r') as f:
    Entrez.email = f.readline()

parser = OptionParser()
parser.add_option("-s", "--search", dest="searchTerm", help="search term")
parser.add_option("-i", "--input", dest="input", help="input file")
parser.add_option("-p", "--pmid", dest="pmid", help="input file")
parser.add_option("-x", "--exclude", dest="exclude", action="store_true", default=False, help="exclude citations")
(options, args) = parser.parse_args()
if len(options.searchTerm) == 0:
    print("Please input a search term with -s or --search")
    exit(1)
print("Search Term: %s") % options.searchTerm

if options.input:
    # pmids = numpy.load(options.pmid)
    # publications = numpy.load(options.input)
    scopusPublications = numpy.load(options.input)
else:
    idlist = getPubmedIds(options.searchTerm, 0)
    print("Number of articles: %s") % len(idlist)
    records = list()
    while len(records) < len(idlist):
        handle = Entrez.efetch(db="pubmed", id=idlist, retstart=len(records), rettype="medline", retmode="text")
        temp = Medline.parse(handle)
        for i, r in tqdm(enumerate(temp)):
            if len(records) == len(idlist):
                break
            elif i == 10000:
                break
            else:
                records.append(r)

    count = 0
    pmids = list()
    publications = list()
    print("Pubmed gather started")
    parsePubmed(records, len(idlist), publications, pmids)
    print("Number of pubmed articles included: %s") % len(publications)
    numpy.save('pubmed_data.npy', publications)
    numpy.save('pmid_data.npy', pmids)

    scopusPublications = list()
    count = 0
    print("Scopus gather started")
    with tqdm(total=len(publications)) as pbar:
        while count < len(publications):
            query = ""
            for id in pmids[count:count + 200]:
                query += " OR PMID(%s)" % id
            query = query[4:]
            content = searchScopus(query, 0)
            if int(content['opensearch:itemsPerPage']) == 1:
                parseScopus(content['entry'], scopusPublications)
            else:
                for r in content['entry']:
                    parseScopus(r, scopusPublications)
            count += 200
            pbar.update(200)
        numpy.save('scopus_data.npy', scopusPublications)
    print("Number of scopus articles included: %s") % len(scopusPublications)

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
    numpy.save('merged_data-%s.npy' % date, scopusPublications)

    scopusids = [v['scopusID'] for v in scopusPublications]
    scopusCites = dict()
    count = 0
    print('Gathering Articles citation data')
    with tqdm(total=len(scopusPublications)) as pbar:
        while count < len(scopusPublications):
            query = ",".join(scopusids[count:count + 25])
            if options.exclude:
                content = citeMetadata(query, excludeSelf=True)
            else:
                content = citeMetadata(query)
            scopusCites.update(content)
            count += 25
            pbar.update(25)
    numpy.save('scopusCites.npy', scopusCites)

    for i, r in tqdm(enumerate(scopusPublications)):
        scopusPublications[i]['Authors'] = scopusCites[r['scopusID']]['scopusAuthors']
        try:
            scopusPublications[i]['Citations'] = scopusCites[r['scopusID']]['Citations']
            scopusPublications[i]['Citations in Past Year'] = scopusCites[r['scopusID']]['Citations in Past Year']
            scopusPublications[i]['Citations Rate'] = scopusCites[r['scopusID']]['Citations Rate']
        except TypeError as e:
            print("TypeError on article %d") % i
            print(e)
        except KeyError as e:
            print("KeyError on article %d") % i
            print(e)
        else:
            print("Cite metadata error on article %d") % i

    scopusPublications = sorted(scopusPublications, key=itemgetter('Citations Rate', 'Citations'), reverse=True)
    numpy.save('data-%s.npy' % date, scopusPublications)

yearlyData = outputYearlyData(scopusPublications)
pubTypesData = outputPubDataPubmed(scopusPublications)
pubSubsetsData = outputPubDataScopus(scopusPublications)

print("Grabing Subsections")
searchTerms = [
    "%s AND (Cyberknife OR cyber knife)" % options.searchTerm,
    "%s AND (gamma knife OR gammaknife)" % options.searchTerm,
    "%s AND linac" % options.searchTerm,
    "%s AND Novalis" % options.searchTerm,
    "%s AND TomoTherapy" % options.searchTerm]
subsections = cyberVgamma(scopusPublications, options.searchTerm)

wb = Workbook()
ws = wb.active
ws.title = 'Overview'
ws.append([
    'Search Term:',
    options.searchTerm,
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
    sum(subsections['cyber']['yearlyData'].values()),
    sum(subsections['linac']['yearlyData'].values()),
    sum(subsections['novalis']['yearlyData'].values()),
    sum(subsections['tomo']['yearlyData'].values())
])
rowNum = 3
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
    rowNum += 1
ws.append(['Publication Subsets'])
rowNum += 1
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
    rowNum += 1
ws.append([
    'Year',
    options.searchTerm,
    subsections['gamma']['query'],
    subsections['cyber']['query'],
    subsections['linac']['query'],
    subsections['novalis']['query'],
    subsections['tomo']['query']
])
rowNum += 1
chartStart = rowNum
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
    rowNum += 1

c1 = LineChart()
dates = Reference(ws, min_row=chartStart, min_col=1, max_col=1, max_row=rowNum)
s1 = Reference(ws, min_row=chartStart - 1, min_col=2, max_col=2, max_row=rowNum - 1)
s2 = Reference(ws, min_row=chartStart - 1, min_col=3, max_col=3, max_row=rowNum - 1)
s3 = Reference(ws, min_row=chartStart - 1, min_col=4, max_col=4, max_row=rowNum - 1)
c1.series.append(Series(s1, title_from_data=True))
c1.series.append(Series(s2, title_from_data=True))
c1.series.append(Series(s3, title_from_data=True))
c1.set_categories(dates)
c1.title = "Articles per year for search: %s" % options.searchTerm
c1.y_axis.title = 'Number of Articles'
c1.x_axis.title = 'Year'
ws.add_chart(c1, "A%d" % (rowNum + 5))

saveWorksheet(wb, 'Pubmed Stats', scopusPublications, options.searchTerm, orderedEntries=True, autoFilter=True)

authorData = outputAuthors(scopusPublications)
saveWorksheet(wb, 'Authors', authorData, searchTerm=options.searchTerm, autoFilter=True)

countryData = outputCountries(scopusPublications)
saveWorksheet(wb, 'Countries', countryData)

journalData = outputJournals(scopusPublications)
saveWorksheet(wb, 'Journals', journalData)

saveWorksheet(wb, 'Gamma Knife', subsections['gamma']['records'], subsections['gamma']['query'], orderedEntries=True, autoFilter=True)
saveWorksheet(wb, 'CyberKnife', subsections['cyber']['records'], subsections['cyber']['query'], orderedEntries=True, autoFilter=True)
saveWorksheet(wb, 'linac', subsections['linac']['records'], subsections['linac']['query'], orderedEntries=True, autoFilter=True)
saveWorksheet(wb, 'Novalis', subsections['novalis']['records'], subsections['novalis']['query'], orderedEntries=True, autoFilter=True)
saveWorksheet(wb, 'TomoTherapy', subsections['tomo']['records'], subsections['tomo']['query'], orderedEntries=True, autoFilter=True)

wb.save('data-%s.xlsx' % date)
