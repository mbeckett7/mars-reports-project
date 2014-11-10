#!/usr/bin/env python
'''
Script for enhancing MARS reports with data from the HOLLIS Presto API and the MARS transactions reports.
Created for the Harvard Library ITS MARS Reports Pilot Project, 2014.
'''
import csv
import glob
import requests
import time
from lxml import html

bib_dict = {} # Dictionary of HOLLIS bib numbers

# Get bib numbers from all of the current report CSV files
for file in glob.glob('*.csv'):
    with open(file, 'rb') as mars_csv:
        reader = csv.reader(mars_csv)
        bib_row = ''
        for index, row in enumerate(reader):
            if index == 0:
                if 'Bib No' in row[0]: # Check column 1
                    bib_row = 0
                elif 'Bib No' in row[1]: # Check column 2
                    bib_row = 1
            else:
                try:
                    if ',' in row[bib_row]: # Get only first bib number if there are multiple ones (e.g., in R00)
                        bib_dict.setdefault(row[bib_row].split(',')[0], None)
                    else: # Otherwise, get the single bib number
                        bib_dict.setdefault(row[bib_row], None)
                except:
                    pass

# Get data from HOLLIS Presto API
# Current settings: LDR 06 (type of record), 008 35-37 (language code), and sublibraries and collection codes
print 'Waiting for Presto API to process', len(bib_dict), 'records ...'
for bib, fields in bib_dict.items(): # bib, fields as key, value
    libraries = []
    marc_url = 'http://webservices.lib.harvard.edu/rest/marc/hollis/' + bib
    presto = requests.get(marc_url)
    marc_xml = presto.content.replace('<record xmlns="http://www.loc.gov/MARC21/slim" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/MARC21/slim   http://www.loc.gov/standards/marcxml/schema/MARC21slim.xsd">','<record>')
    marc_record = html.fromstring(marc_xml)
    ldr06 = marc_record.xpath('//leader/text()')[0][6] #Get LDR byte 06 (type of record)
    language = marc_record.xpath('//controlfield[@tag="008"]/text()')[0][35:38] #Get 008 bytes 35-37 (language code)
    own = marc_record.xpath('//datafield[@tag="OWN"]/subfield/text()') #Get list of OWN (holding library) fields
    collection = marc_record.xpath('//datafield[@tag="852"]/subfield[@code="c"]/text()') #Get collection code from 852 $c
    for (i, j) in zip(own, collection): #Combine own field and collection code and format as a text string
        libraries.append( i + ' (' + j + ')')
    libraries = '; '.join(libraries)
    bib_dict[bib] = [ldr06, language, libraries]
    if index % 25 == 0: # Delay to avoid being locked out of API
        time.sleep(2)
    else:
        time.sleep(1)


enhanced_dict = {} # Dictionary of enhanced data

# Add data to CSV files
for file in glob.glob('*.csv'):
    with open(file, 'rb') as mars_csv:
        reader = csv.reader(mars_csv)

        enhanced_rows = []
        enhanced_file = file[:-4] + '_enhanced.csv'

        for index, row in enumerate(reader):
            if index == 0:
                row[:-3] += ['LDR 06','Language','Libraries']
            elif row[0] in bib_dict:
                row[:-3] += bib_dict[row[0]]
            elif row[1] in bib_dict:
                row[:-3] += bib_dict[row[1]]
            else:
                row[:-3] += ['','','']
            enhanced_rows.append(row)
        enhanced_dict[enhanced_file] = enhanced_rows

# TO DO: Exclude reports without bib numbers (e.g., authority reports) from enhancement
# TO DO: Split music records (LDR 06) into separate files
# TO DO: Split 'No Replacement Found' (R04) into separate files

for csv_file, csv_data in enhanced_dict.items():
    with open(csv_file, 'wb') as enhanced_csv:
            writer = csv.writer(enhanced_csv)
            writer.writerows(csv_data)
            #TO DO: Encoding is not correct for enhanced files; to be fixed
    print csv_file, 'has been created'
