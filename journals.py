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

	journalData = list()
	for journal in journalSet:
		journalData.append(calc_journalData(journal, records))

	journalData = sorted(journalData, key=itemgetter('Number of articles'), reverse=True)

	return journalData
