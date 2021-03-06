from optparse import OptionParser
from Bio import Entrez, Medline
from pubmedHelpers import getPubmedIds, parsePubmed, pubmedData, saveWorksheet
from authors import outputAuthors, outputCountries
from journals import outputJournals
from scopusAPI import searchScopus, parseScopus, citeMetadata
import numpy
from openpyxl import Workbook
from datetime import datetime
from operator import itemgetter
from tqdm import tqdm
with open('config.txt', 'r') as f:
    Entrez.email = f.readline()

parser = OptionParser()
parser.add_option("-s", "--search", dest="searchTerm", help="search term")
parser.add_option("-i", "--input", dest="input", help="input file")
parser.add_option("-p", "--pmid", dest="pmid", help="input file")
parser.add_option("-r", "--remove", dest="remove", help="Remove articles heavily self-cited")
parser.add_option("-x", "--exclude", dest="exclude", action="store_true", default=False, help="exclude citations")
parser.add_option("-f", "--force", dest="force", action="store_true", default=False, help="force all articles to be included")
(options, args) = parser.parse_args()
if len(options.searchTerm) == 0:
    print("Please input a search term with -s or --search")
    exit(1)
print("Search Term: %s") % options.searchTerm

date = datetime.now().strftime('%Y-%m-%d')
if options.exclude:
    suffix = "-%s-self_exclude" % date
else:
    suffix = "-%s" % date

if options.input:
    # pmids = numpy.load(options.pmid)
    # publications = numpy.load(options.input)
    scopusPublications = numpy.load(options.input).tolist()
    # scopusCites = numpy.load(options.pmid)
    # scopusCites = scopusCites.item()
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
    numpy.save('pubmed_data-%s.npy' % date, publications)
    numpy.save('pmid_data-%s.npy' % date, pmids)

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
        numpy.save('scopus_data-%s.npy' % date, scopusPublications)
    print("Number of scopus articles included: %s") % len(scopusPublications)

    print('Merging data from Pubmed to Scopus records')
    merged = list()
    for pub in tqdm(publications):
        found = False
        for r in scopusPublications:
            if r['PMID'] == pub['PMID']:
                found = True
                temp = r
                temp['Full Author Names'] = pub['Full Author Names']
                temp['Authors'] = pub['Authors']
                temp['Source'] = pub['Source']
                temp['Language'] = pub['Language']
                temp['Publication Date'] = pub['Publication Date']
                temp['Publication Subset'] = pub['Publication Type']
                merged.append(temp)
        if not found and options.force:
            temp = pub
            temp['Authors'] = None
            temp['scopusID'] = None
            temp['Country of Origin'] = 'Unknown'
            temp['Publication Subset'] = temp['Publication Type']
            temp['Publication Type'] = None
            merged.append(temp)
    print("Merged total: %d") % len(merged)
    numpy.save('merged_data-%s.npy' % date, merged)

    scopusPublications = merged
    scopusids = [v['scopusID'] for v in scopusPublications if v['scopusID'] is not None]
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
            if content is not None:
                scopusCites.update(content)
            count += 25
            pbar.update(25)
    numpy.save('scopusCites%s.npy' % suffix, scopusCites)
    if options.exclude:
        exit(0)
    for i, r in tqdm(enumerate(scopusPublications)):
        if scopusPublications[i]['Authors'] is not None:
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

print(len(scopusPublications))
if options.remove:
    fake = numpy.load(options.remove)
    for r in scopusPublications:
        if r['scopusID'] in fake:
            scopusPublications.remove(r)
    suffix = "-%s-corrected" % date
print(len(scopusPublications))
scopusPublications = sorted(scopusPublications, key=itemgetter('Citations Rate', 'Citations'), reverse=True)
numpy.save('data%s.npy' % suffix, scopusPublications)

wb = Workbook()
subsections = pubmedData(wb, scopusPublications, options.searchTerm, suffix)
numpy.save('subsections%s.npy' % suffix, subsections)

authorData = outputAuthors(scopusPublications)
saveWorksheet(wb, 'Authors', authorData, searchTerm=options.searchTerm, autoFilter=True)

countryData = outputCountries(scopusPublications)
saveWorksheet(wb, 'Countries', countryData)

journalData = outputJournals(scopusPublications)
saveWorksheet(wb, 'Journals', journalData)

wb.save('data%s.xlsx' % suffix)
