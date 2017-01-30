from __future__ import division
from urllib import urlencode
import urllib2
import xmltodict
import time
import requests
scopusIDPrefix = 10
with open('config.txt', 'r') as f:
    email = f.readline()
    scopusAPI = f.readline()


def countryOrigin(authorID):
    query = urlencode({'field': 'affiliation-country'})
    url = 'https://api.elsevier.com/content/author/author_id/%s?%s' % (
        authorID, query)
    req = urllib2.Request(url)
    req.add_header('X-ELS-APIKey', scopusAPI)
    req.add_header('Accept', 'text/xml')
    content = None
    while content is None:
        try:
            resp = urllib2.urlopen(req)
            content = xmltodict.parse(resp.read())
        except urllib2.HTTPError, e:
            print('No author metadata, error code - %s.') % e.code
            if e.code == 404:
                return 'Unknown'
            elif e.code == 401:
                print('Not authenticated, check API Key or VPN connection')
                exit(1)
        else:
            print('Error on Scopus API, retrying...')
            time.sleep(1)
    try:
        return content['author-retrieval-response']['affiliation-current']['affiliation-country']
    except:
        return 'Unknown'


def searchScopus(searchTerm, start=0):
    content = None
    while content is None:
        try:
            response = requests.get(
                url="http://api.elsevier.com/content/search/scopus",
                params={
                    "query": searchTerm,
                    "count": "200",
                    "sort": "-citedby-count",
                            "content": "all",
                            "start": str(start)
                },
                headers={
                    "Accept": "application/xml",
                    "X-ELS-APIKey": scopusAPI,
                },
            )
            res = xmltodict.parse(response.content)
            content = res['search-results']['entry']
            return res['search-results']
        except requests.exceptions.RequestException:
            print('HTTP Request failed')
        except KeyError:
            if res['service-error']['status']['statusText']:
                print("Error:")
                print(res['service-error']['status']['statusText'])
                exit(1)
            print('KeyError, retrying request:')
        else:
            print('Scopus Search Error, retrying...')


def parseScopus(r, array):
    pub = dict()
    if 'dc:creator' in r:
        pub['Authors'] = r['dc:creator']
        pub['Publication Date'] = r['prism:coverDisplayDate']
        pub['Title'] = r['dc:title']
        pub['Publication Type'] = r.get('subtypeDescription', 'Article')
        pub['Journal Title'] = r.get('prism:publicationName', 'Unknown')
        pub['Source'] = None
        pub['Language'] = 'Unknown'
        pub['scopusID'] = r['dc:identifier'][scopusIDPrefix:]
        pub['PMID'] = r['pubmed-id']
        pub['Citations'] = int(r.get('citedby-count', 0))
        pub['Citations in Past Year'] = 0
        pub['Citations Rate'] = 0
        try:
            pub['Country of Origin'] = r['affiliation']['affiliation-country']
        except TypeError:
            pub['Country of Origin'] = r['affiliation'][0]['affiliation-country']
        except KeyError:
            pub['Country of Origin'] = 'Unknown'
        array.append(pub)


def citeMetadata(query, excludeSelf=False):
    content = None
    params = {
        "scopus_id": query,
        "date": "1938-2016",
        "citation": "exclude-self"
    }
    if not excludeSelf:
        del params['citation']
    while content is None:
        try:
            response = requests.get(
                url="http://api.elsevier.com/content/abstract/citations",
                params=params,
                headers={
                    "Accept": "application/xml",
                    "X-ELS-APIKey": scopusAPI,
                },
            )
            content = xmltodict.parse(response.content)
        except requests.exceptions.RequestException:
            print('HTTP Request failed')
        else:
            pass
            # print('Scopus Search Error, retrying...')
    recordsCites = dict()
    try:
        citations = content['abstract-citations-response']['citeInfoMatrix'][
            'citeInfoMatrixXML']['citationMatrix']['citeInfo']
    except KeyError:
        return None
    for citation in citations:
        scopusId = citation['dc:identifier'][10:]
        temp = dict()
        authors = citation['author']
        try:
            temp['scopusAuthors'] = [a['index-name'] for a in authors]
        except TypeError:
            temp['scopusAuthors'] = [authors['index-name']]
        temp['scopusID'] = citation['dc:identifier'][10:]
        temp['Citations'] = int(citation['rowTotal'])
        try:
            temp['Citations in Past Year'] = int(citation['cc'][-1])
            yearCites = [int(i) for i in citation['cc']]
            yearCites = yearCites[int(citation['sort-year']) - 1938:]
            temp['Citations Rate'] = sum(yearCites) / len(yearCites)
        except KeyError:
            temp['Citations in Past Year'] = 0
            temp['Citations Rate'] = 0
        recordsCites[scopusId] = temp
    return recordsCites
