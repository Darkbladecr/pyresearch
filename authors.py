from pubmedHelpers import distinctSet
from operator import itemgetter
from collections import OrderedDict


def calc_hIndex(citations):
    N = len(citations)
    if N == 0:
        return 0
    i = 0
    while i < N and citations[i]['Citations'] >= (i + 1):
        i += 1
    return i
# authorName = 'Author Names'
# authorName = 'Full Author Names'


def calc_authorData(authorData, authorSet, records, authorName):
    for r in records:
        if r[authorName] is not None:
            for author in r[authorName]:
                authorData[author]['articles'].append({'PMID': r['PMID'], 'Citations': int(r['Citations'])})
    for a in authorSet:
        authorData[a]['articles'] = sorted(authorData[a]['articles'], key=itemgetter('Citations'), reverse=True)
        authorData[a]['h-index'] = calc_hIndex(authorData[a]['articles'])
    return authorData


def outputAuthors(records, authorName='Authors'):
    authorSet = distinctSet(records, authorName)
    print('Total number of authors: %d') % len(authorSet)
    authorData = {k: {'articles': list()} for k in authorSet}
    authorData = calc_authorData(authorData, authorSet, records, authorName)

    authorData_flat = list()
    orderedEntries = ('Name', 'h-index', 'Articles', 'Cited')
    for k, v in authorData.items():
        name = k
        flat_data = v
        totalCitations = 0
        for article in flat_data['articles']:
            totalCitations += article['Citations']
        flat_data['Cited'] = totalCitations
        flat_data['Articles'] = int(len(flat_data['articles']))
        flat_data['Name'] = name
        orderedData = OrderedDict((k, flat_data[k]) for k in orderedEntries)
        authorData_flat.append(orderedData)
    return sorted(authorData_flat, key=itemgetter('h-index', 'Cited'), reverse=True)


def outputCountries(records):
    countrySet = distinctSet(records, 'Country of Origin')
    print('Total number of countries: %d') % len(countrySet)
    countryData = {k: {'Country': k, 'Number of articles': 0} for k in countrySet}
    for r in records:
        if r['Country of Origin'] is not None:
            countryData[r['Country of Origin']]['Number of articles'] += 1
    countryData = sorted(countryData.values(), key=itemgetter('Number of articles'), reverse=True)
    return countryData
