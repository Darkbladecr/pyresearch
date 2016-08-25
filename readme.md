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
- Article country of origin

![Example Data](https://github.com/Darkbladecr/pyresearch/blob/master/screenshots/example_data.png?raw=true)
![Example Author Data](https://github.com/Darkbladecr/pyresearch/blob/master/screenshots/example_authors.png?raw=true)

## Download & Setup
1. Download this repository's [Latest Release](https://github.com/Darkbladecr/pyresearch/archive/master.zip) and navigate Terminal to the downloaded file:
```
$ cd Downloads
$ unzip pyresearch-master.zip 
$ cd pyresearch-master
```

2. Install all the python module requirements by using the command:
`sudo pip install -r requirements.txt`

3. Setup your `config.txt` file with your email address in the first line (required by PubMed's API) and the second line with your Scopus API key, which you can get from [dev.elsevier.com](https://dev.elsevier.com/user/login). For example:

```
user@pyresearch.com
f024b78838b3a32f966feaf30f4361dc97c840e7
```



## Basic Usage
Open Terminal and go to the folder where you unziped the files in step 1. Then use the command line with `python pyresearch` for the search with the argument`-s OR --search "Search Term Here"`:

For example a search for the articles on Focused Ultrasound in the Brain could be achived with:

`python pyresearch.py -s "Focused Ultrasound AND Brain"`

`-s or --search` is used for your search query.

![Example Terminal](https://github.com/Darkbladecr/pyresearch/blob/master/screenshots/terminal.png?raw=true)

The script will go through its various steps and let you know the progress. When it is complete it will save an excel document into the same directory as the script.

Note it takes about 1 minute per 300 results, you can cancel the script at any time with `CTRL + C`.

### Adding Subsearches
Once the general search is complete, you are able to do subsection searches by narrowing your original search term. The great benefit of this subsearch is that it reuses the data saved locally (*.npy file) which will make the search much faster. The articles found matching your query will be added as an extra worksheet onto the original excel file. Additionally, more detailed information about the subsearch (authors, journals etc.) will be outputed to the folder `subsections`. 

For example, if I wanted to find all the articles based on Focused Ultrasound in the Brain for Essential Tremor I would use the following command:

`python pyresearch_subsection.py -s "Focused Ultrasound AND Brain AND Essential Tremor" -t "Essential Tremor" -i "data-2016-08-23"`

`-s or --search` is used for your search query.

`-t or --title` is used for the worksheet title in your excel file and for the name of the ouput excel file in the subsections folder.

`-i or --input` is used for the filename of your `*.xlsx` file and your `*.npy` note they must share the same name. The inital pyresearch script outputs these files, however note if you rename only one then the subsection script will not work.

![Example Subsection Terminal](https://github.com/Darkbladecr/pyresearch/blob/master/screenshots/example_subsection.png?raw=true)

## Troubleshooting
If you get an error such as:
```
Error: Exceeds the maximum number allowed for the service level and/or view
OR
Error: APIKey 123456789abcdef with IP address 100.100.100.80 is unrecognized or has insufficient privileges for access to this resource
```
Double check that your Scopus API Key is correctly setup or has appropriate privlidges. You can try this directly on the [dev.elsevier.com](http://dev.elsevier.com/metadata.html#!/Citations_Overview/CitationsOverview) website.
