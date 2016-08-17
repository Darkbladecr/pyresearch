# PyResearch
![Python version](https://img.shields.io/badge/Python-2.7-brightgreen.svg) ![MIT Licence](https://img.shields.io/badge/Licence-MIT-blue.svg)
___
A Python script with Scopus and PubMed APIs built in to give you an overview of the most salient research in your search term, including:

- Article's per year
- Article subtypes
- Citation Counts
- Citation Index (citations per year)
- Author [*h*-index](https://en.wikipedia.org/wiki/H-index)
- Journal publications

## Download & Setup
1. Download this repository's [Latest Release](https://github.com/Darkbladecr/pyresearch/archive/master.zip) and navigate Terminal to the downloaded file:
```
$ cd Downloads
$ unzip pyresearch-master.zip 
$ cd pyresearch-master
```

2. Install all the python module requirements by using the command:
`sudo pip install -r requirements.txt`

3. Setup your `config.txt` file with your email address in the first line (required by PubMed's API) and the second line with your Scopus API key, which yu can get from [dev.elsevier.com](https://dev.elsevier.com/user/login). For example:

```
user@pyresearch.com
f024b78838b3a32f966feaf30f4361dc97c840e7
```



## Basic Usage
Open Terminal and go to the folder where you unziped the files in step 1. Then use the command line with `python pyresearch` for the search with the argument`-s OR --search "Search Term Here"`:

For example a search for the articles on Focused Ultrasound in the Brain could be achived with:

`python pyresearch -s "Focused Ultrasound AND Brain"`

![Example](https://github.com/Darkbladecr/pyresearch/blob/master/example.png?raw=true)

The script will go through its various steps and let you know the progress. When it is complete it will save an excel document into the same directory as the script.

Note it takes about 1 minute per 300 results, you can cancel the script at any time with `CTRL + C`.

## Troubleshooting
If you get an error such as:
```
Error: Exceeds the maximum number allowed for the service level and/or view
OR
Error: APIKey 123456789abcdef with IP address 100.100.100.80 is unrecognized or has insufficient privileges for access to this resource
```
Double check that your Scopus API Key is correctly setup or has appropriate privlidges. You can try this directly on the [dev.elsevier.com](http://dev.elsevier.com/metadata.html#!/Citations_Overview/CitationsOverview) website.
