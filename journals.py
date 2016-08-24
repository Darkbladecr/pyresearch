from pubmedHelpers import distinctSet
from operator import itemgetter


def calc_journalData(query, records):
	count = 0
	for record in records:
		if query in record['Journal Title']:
			count += 1
	return {'Journal title': query, 'Number of articles': count}


def outputJournals(records):
	journalSet = distinctSet(records, 'Journal Title')
	print('Total number of journals: %d') % len(journalSet)
	journalData = {k: {'Journal Title': k, 'Number of articles': 0} for k in journalSet}
	for r in records:
		journalData[r['Journal Title']]['Number of articles'] += 1
	journalData = sorted(journalData.values(), key=itemgetter('Number of articles'), reverse=True)
	return journalData
