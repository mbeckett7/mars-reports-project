#!/usr/bin/env python
'''
Script for removing Harmatheque e-book records from MARS R00 report.
Created for the Harvard Library ITS MARS Reports Pilot Project, 2014.
'''
import codecs
import csv
import glob
import requests
import time
from lxml import html

nets = []
not_nets = []


counter = 0

for file in glob.glob('*.csv'):
    with open(file, 'rb') as mars_csv:
        reader = csv.reader(mars_csv)
        for row in reader:
            if row[15] == 'NET (GEN)':
                bib = row[1]
                marc_url = 'http://webservices.lib.harvard.edu/rest/marc/hollis/' + bib
                presto = requests.get(marc_url)
                marc_xml = presto.content.replace('<record xmlns="http://www.loc.gov/MARC21/slim" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/MARC21/slim   http://www.loc.gov/standards/marcxml/schema/MARC21slim.xsd">','<record>')
                marc_record = html.fromstring(marc_xml)
                if marc_record.xpath('//leader/text()'):
                    h09 = marc_record.xpath('//datafield[@tag="H09"]/subfield[@code="m"]/text()') # Get list of H09 subfield m fields
                if 'harmatheque' in h09:
                    nets.append(row)
                    
                else:
                    not_nets.append(row)
                time.sleep(1)
            else:
                not_nets.append(row)

if len(nets) > 0:
    with open('r00_harmatheque.csv', 'wb') as output:
        output.write(codecs.BOM_UTF8)
        writer = csv.writer(output, quoting=csv.QUOTE_ALL,quotechar='"')
        writer.writerows(nets)
if len(not_nets) > 0:
    with open('r00_not_harmatheque.csv', 'wb') as output:
        output.write(codecs.BOM_UTF8)
        writer = csv.writer(output, quoting=csv.QUOTE_ALL,quotechar='"')
        writer.writerows(not_nets)
