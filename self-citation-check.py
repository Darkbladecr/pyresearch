from __future__ import division
from optparse import OptionParser
import numpy

parser = OptionParser()
parser.add_option("-i", "--input", dest="input", help="input file")
(options, args) = parser.parse_args()


def numpyImport(file):
    output = numpy.load("%s.npy" % file)
    output = output.item()
    return output


allCites = numpyImport(options.input)
exCites = numpyImport("%s-self_exclude" % options.input)
fake = list()
for scopusId in allCites.keys():
    if allCites[scopusId]['Citations'] > 0 and exCites[scopusId]['Citations'] > 0:
        if 1 - (exCites[scopusId]['Citations'] / allCites[scopusId]['Citations']) > 0.4:
            fake.append(scopusId)
            print("%d vs %d" % (allCites[scopusId]['Citations'], exCites[scopusId]['Citations']))
print("%d articles were found to be heavily self-cited" % len(fake))
numpy.save('self-cited.npy', fake)
