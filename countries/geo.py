import csv
from geopy.geocoders import Nominatim
data = csv.DictReader(open("countries.csv"))
geolocator = Nominatim()
countries = list()
for country in data:
    temp = country
    location = geolocator.geocode(temp['Country'])
    temp['Coordinates'] = "%s %s" % (location.latitude, location.longitude)
    countries.append(temp)

with open('countries_done.csv', 'w') as csvfile:
    fieldnames = ['Country', 'Number of articles', 'Coordinates']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for row in countries:
        writer.writerow(row)
