#!/usr/bin/env python
'''
Script for enhancing MARS reports with data from the HOLLIS Presto API and the MARS transactions reports.
Created for the Harvard Library ITS MARS Reports Pilot Project, 2014.
'''
import codecs
import csv
import glob
import requests
import time
from lxml import html

bib_dict = {} # Dictionary of HOLLIS bib numbers
enhanced_dict = {} # Dictionary of enhanced data
music_reports = ['R00','R06','R07','R11'] # List of reports to check for music headings
no_replace_reports = ['R04'] # List of reports to check for 'No Replacement Found' records
no_enhance_reports = ['R03', 'R04'] # List of reports that cannot or will not be enhanced
# Authority reports without bib numbers cannot be enhanced by the HOLLIS Presto API

# Get bib numbers from all of the current report CSV files
for file in glob.glob('*.csv'):
    with open(file, 'rb') as mars_csv:
        reader = csv.reader(mars_csv)
        bib_row = ''
        report_no = file[:4].replace('_','')

        if report_no not in no_enhance_reports: # Only enhance reports that can be enhanced
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
if len(bib_dict) > 0:
    print 'Waiting for Presto API to process', len(bib_dict), 'records ...'
for bib, fields in bib_dict.items(): # bib, fields as key, value
    libraries = []
    marc_url = 'http://webservices.lib.harvard.edu/rest/marc/hollis/' + bib
    presto = requests.get(marc_url)
    marc_xml = presto.content.replace('<record xmlns="http://www.loc.gov/MARC21/slim" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/MARC21/slim   http://www.loc.gov/standards/marcxml/schema/MARC21slim.xsd">','<record>')
    marc_record = html.fromstring(marc_xml)
    if marc_record.xpath('//leader/text()'):
        ldr06 = marc_record.xpath('//leader/text()')[0][6] # Get LDR byte 06 (type of record)
        language = marc_record.xpath('//controlfield[@tag="008"]/text()')[0][35:38] # Get 008 bytes 35-37 (language code)
        own = marc_record.xpath('//datafield[@tag="OWN"]/subfield/text()') # Get list of OWN (holding library) fields
        collection = marc_record.xpath('//datafield[@tag="852"]/subfield[@code="c"]/text()') # Get collection code from 852 $c
        for (i, j) in zip(own, collection): # Combine own field and collection code and format as a text string
            libraries.append( i + ' (' + j + ')')
        libraries = '; '.join(libraries)
        bib_dict[bib] = [ldr06, language, libraries] # Add HOLLIS data to dictionary of bib numbers      
    else:
        bib_dict[bib] = ['','','']
    time.sleep(1)

# Add HOLLIS data to CSV files
for file in glob.glob('*.csv'):
    with open(file, 'rb') as mars_csv:
        reader = csv.reader(mars_csv)

        enhanced_rows = []
        music_rows = [] # For music reports
        no_replace_rows = [] # For 'No Replacement Found' R04 report
        report_no = file[:4].replace('_','')

        if report_no not in no_enhance_reports: # Only enhance reports that can be enhanced
            for index, row in enumerate(reader):
                if ',' in row[0]:
                    compare_col = row[0].split(',')[0]
                else:
                    compare_col = row[0]
                if ',' in row[1]:
                    compare_col_2 = row[1].split(',')[0]
                else:
                    compare_col_2 = row[1]
                    
                if index == 0: # Get header row
                    row[0] = row[0].replace('"', '')
                    row[:-3] += ['LDR 06','Language','Libraries']
                    enhanced_rows.append(row)
                elif compare_col in bib_dict and bib_dict[compare_col] != '': # Check first column
                    row[:-3] += bib_dict[compare_col]
                    if report_no in music_reports: # For music reports, check for LDR 06 c, d, or j 
                        if bib_dict[compare_col][0] == 'c' or bib_dict[compare_col][0] == 'd' or bib_dict[compare_col][0] == 'j': 
                            music_rows.append(row) # Put in music report
                        else:
                            enhanced_rows.append(row) # Put in non-music report
                    else:
                        enhanced_rows.append(row)
                elif compare_col_2 in bib_dict and bib_dict[compare_col_2] != '': # Check second column
                    row[:-3] += bib_dict[compare_col_2]
                    if report_no in music_reports: # For music reports, check for LDR 06 c, d, or j
                        if bib_dict[compare_col_2][0] == 'c' or bib_dict[compare_col_2][0] == 'd' or bib_dict[compare_col_2][0] == 'j':
                            music_rows.append(row) # Put in music report
                        else:
                            enhanced_rows.append(row) # Put in non-music report
                    else:
                        enhanced_rows.append(row)
                else:
                    row[:-3] += ['','','']
                    enhanced_rows.append(row)
        elif report_no in no_replace_reports:
            for index, row in enumerate(reader):
                if index == 0: # Get header row
                    row[0] = row[0].replace('"', '')
                    enhanced_rows.append(row)
                else:
                    if row[1] == 'No replacement found':
                        no_replace_rows.append(enhanced_rows[-1]) # Add previous (i.e., Old) row to no replacement report
                        del enhanced_rows[-1] # Delete Old row from report with replacement rows
                        no_replace_rows.append(row) # Add current (i.e., New) row to no replacement report
                    else:
                        enhanced_rows.append(row)

        if enhanced_rows or music_rows or no_replace_rows:
            enhanced_dict[file[:-4] + '_enhanced.csv'] = enhanced_rows # Create enhanced report

            if music_rows: # Create music report
                music_rows.insert(0, enhanced_rows[0])
                enhanced_dict[file[:-4] + '_enhanced_music.csv'] = music_rows

            if no_replace_rows: # Create music report
                no_replace_rows.insert(0, enhanced_rows[0])
                enhanced_dict[file[:-4] + '_enhanced_noreplace.csv'] = no_replace_rows
            # TO DO: Row numbers are not correct for separated reports; this should be fixed if it matters to the processing team
        else:
            print file, 'was not enhanced'
                
            
# Create CSV files for enhanced reports
for csv_file, csv_data in enhanced_dict.items():
    with open(csv_file, 'wb') as output:
            output.write(codecs.BOM_UTF8)
            writer = csv.writer(output, quoting=csv.QUOTE_ALL,quotechar='"')
            writer.writerows(csv_data)
    print csv_file, 'has been created'
