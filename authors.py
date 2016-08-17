from pubmedHelpers import distinctSet
from operator import itemgetter
from collections import OrderedDict
from tqdm import tqdm


def calc_hIndex(citations):
    N = len(citations)
    if N == 0:
        return 0
    i = 0
    while i < N and citations[i]['Citations'] >= (i + 1):
        i += 1
    return i


def calc_authorData(query, records, authorName):
    count = 0
    idlist = {query: {'articles': list()}}
    for record in records:
        if query in record[authorName]:
            count += 1
            idlist[query]['articles'].append(
                {'PMID': record['PMID'], 'Citations': int(record['Citations'])})
    idlist[query]['articles'] = sorted(
        idlist[query]['articles'], key=itemgetter('Citations'), reverse=True)
    hIndex = calc_hIndex(idlist[query]['articles'])
    idlist[query]['h-index'] = hIndex
    return idlist

# authorName = 'Author Names'
# authorName = 'Full Author Names'


def outputAuthors(records, authorName='Authors'):
    authorSet = distinctSet(records, authorName)
    print('Total number of authors: %d') % len(authorSet)
    authorData = list()
    for author in tqdm(authorSet):
        authorData.append(calc_authorData(author, records, authorName))
    authorData_flat = list()
    orderedEntries = ('Name', 'h-index', 'Articles', 'Cited')
    for author in tqdm(authorData):
        name = author.keys()[0]
        flat_data = author.values()[0]
        totalCitations = 0
        for article in flat_data['articles']:
            totalCitations += article['Citations']
        flat_data['Cited'] = totalCitations
        flat_data['Articles'] = int(len(flat_data['articles']))
        flat_data['Name'] = name
        orderedData = OrderedDict((k, flat_data[k]) for k in orderedEntries)
        authorData_flat.append(orderedData)
    return sorted(authorData_flat, key=itemgetter('h-index'), reverse=True)
