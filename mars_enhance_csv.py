
def cataloger_assignment(report_no, language): #function for determining cataloger auto-assignment, if any
    
    cataloger = ''
    # reports in which catalogers are auto-assigned
    #random_assignment_reports = ['R00', 'R03', 'R04', 'R06', 'R07', 'R25'] #
    random_assignment_reports = [] #
    language_assignment_reports = ['R13', 'R14']
    
    # lists of catalogers for assigning report rows
#==============================================================================
#     cataloger_by_language = {'lat':['John','Bruce','Anthony'], 'ita':['Anthony','Mary Jane','Karen'], 'nap':['Anthony','Mary Jane','Karen'], 
#                              'spa':['Isabel','Anthony'], 'cat':['Isabel','Anthony'], 'glg':['Isabel','Anthony'], 'gag':['Isabel','Anthony'], 'por':['John','Isabel'],
#                              'fre':['John','Anthony','Bruce','Mary Jane'], 'frm':['John','Anthony','Bruce','Mary Jane'], 'fro':['John','Anthony','Bruce','Mary Jane'], 
#                              'ger':['John','Bruce', 'Mary Jane'], 'goh':['John','Bruce', 'Mary Jane'], 'gmh':['John','Bruce', 'Mary Jane'], 'gem':['John','Bruce', 'Mary Jane'], 
#                              'dut':['John','Mary Jane','Bruce'], 'dan':['John','Bruce'], 'nor':['John','Bruce'], 'swe':['John','Bruce'], 'ice':['John','Bruce'], 
#                              'chi':'Jia Lin', 'jpn':'Jia Lin', 'myn':'Isabel', 'afr':'Mary Jane', 
#                              'eng':['John', 'Jia Lin', 'Mary Jane', 'Isabel', 'Anthony', 'Bruce', 'Michael','Karen']} # current language assignments
#==============================================================================

    # lists of catalogers for assigning report rows
    cataloger_by_language = {'ita':'Rebecca', 'nap':'Rebecca', 
                             'spa':['Pam', 'Ann'], 'cat':['Pam', 'Ann'], 'glg':['Pam', 'Ann'], 'gag':['Pam', 'Ann'],
                             'fre':['Rebecca', 'Ann'], 'frm':['Rebecca', 'Ann'], 'fro':['Rebecca', 'Ann'], 
                             'ger':['Rebecca','Pam'], 'goh':['Rebecca','Pam'], 'gmh':['Rebecca','Pam'], 'gem':['Rebecca','Pam'],
                             'eng':['Rebecca', 'Pam', 'Ann', 'Delana' ],
                             'other':['Delana', 'Delana', 'Delana', 'Delana', 'Ann', 'Ann', 'Pam', 'Pam', 'Rebecca']} # current language assignments



    if report_no in random_assignment_reports:
        cataloger = random.choice(cataloger_by_language['eng'])
        #print report_no + ' ' + language + ' ' + cataloger + ' (random)'
    elif report_no in language_assignment_reports:
        if language in cataloger_by_language.keys():
            if type(cataloger_by_language[language]) is str: 
                cataloger = cataloger_by_language[language]
            elif type(cataloger_by_language[language]) is list:
                cataloger = random.choice(cataloger_by_language[language])
            else: 
                cataloger = random.choice(cataloger_by_language['other'])
            #print report_no + ' ' + language + ' ' + cataloger + ' (non-random)'
        else:
            cataloger = random.choice(cataloger_by_language['other'])
    else:
        cataloger = random.choice(cataloger_by_language['other']) # assigning all other languages randomly
        #print report_no + ' ' + language + ' no assignment'
        
    return cataloger #to do: determine why this sometimes throws an error if cataloger is not defined as blank at beginning of function


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
import os
from lxml import html
import random # for assigning catalogers where language expertise overlaps

# Query server in order to get report_date variable for use in changing to working directory with filtered reports (from Create script) - improve this!
import datetime
from pyquery import PyQuery as pq
base_url = 'http://lms01.harvard.edu/mars-reports/' # top-level directory page
r = requests.get(base_url)
d = pq(r.content)
month_url = base_url + d('a')[0].text  # Gets first linked url from page; change index number to get reports for an earlier month
report_date = datetime.datetime.strptime(d('a')[0].text[:-5], '%b-%y').strftime('%Y_%m') # Convert date in URL to numeric date for file naming later in script
os.chdir('C:\\Users\\beckett\\MARS Filtered Reports\\MARS Filtered Reports ' + report_date)

bib_dict = {} # Dictionary of HOLLIS bib numbers -- example key/value: {'009151020': ['c', 'ita', 'MUS (ISHAM); MUS (HD)', '']}
enhanced_dict = {} # Dictionary of enhanced data
#music_reports = ['R00','R06','R07','R11', 'R28', 'R42', 'R119'] # List of reports to check for music headings pre-Feb 2015
#music_reports = ['R00','R06','R07','R11','R13','R14','R119'] # List of reports to check for music headings Feb 2015 forward # Added R13 and R14 for March 2015 forward # added R119 for April 2015 forward
music_reports = ['R00','R06','R11','R13','R14','R119'] #removed R07 for enhancement of R07 and R30 series reports
net_reports = ['R13','R14'] # List of reports to check for NET-only records
no_replace_reports = ['R04'] # List of reports to check for 'No Replacement Found' records
no_enhance_reports = ['R03','R04'] # List of reports that cannot or will not be enhanced
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
    time.sleep(.2)

# Add HOLLIS data to CSV files
for file in glob.glob('*.csv'):
    with open(file, 'rb') as mars_csv:
        reader = csv.reader(mars_csv)

        enhanced_rows = []
        music_rows = [] # For music reports
        net_rows = [] # for net-only rows
        no_replace_rows = [] # For 'No Replacement Found' R04 report
        report_no = file[:4].replace('_','')

        #print report_no # debugging
        
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
                    if report_no == 'R00':
                        row[:-3] += ['LDR 06','Language','Libraries','Assigned To']
                    else:
                        row += ['LDR 06','Language','Libraries','Assigned To']
                    row += ['No change needed?','Notes','For Amy','Time Spent']
                    if report_no == 'R00':
                        row += ['Heading Matches Near Match?','# Headings Attached','Time Spent']
                    enhanced_rows.append(row)
                elif compare_col in bib_dict and bib_dict[compare_col] != '': # Check first column
                    language = bib_dict[compare_col][1] #get language fron bib_dict
                    row[:-3] += bib_dict[compare_col] #add Hollis data to report row
                    row[:-3] += [cataloger_assignment(report_no, language)] #add cataloger assignment based on language
                    if report_no in music_reports: # For music reports, check for LDR 06 c, d, or j 
                        if bib_dict[compare_col][0] == 'c' or bib_dict[compare_col][0] == 'd' or bib_dict[compare_col][0] == 'j': 
                            music_rows.append(row) # Put in music report
                        elif any('libretto' in row_string.lower() for row_string in row): # check if any lower-cased string in the list = "libretto"
                            music_rows.append(row) # Put in music report 
                        elif report_no in net_reports:
                            if ('NET' in bib_dict[compare_col][2]) and (';' not in bib_dict[compare_col][2]):
                                net_rows.append(row)
                            else:
                                enhanced_rows.append(row)                        
                        else:
                            enhanced_rows.append(row) # Put in non-music report					
                    else:
                        enhanced_rows.append(row)
                elif compare_col_2 in bib_dict and bib_dict[compare_col_2] != '': # Check second column
                    language = bib_dict[compare_col_2][1] #get language fron bib_dict
                    row[:-3] += bib_dict[compare_col_2] #add Hollis data to report row
                    row[:-3] += [cataloger_assignment(report_no, language)] #add cataloger assignment based on language
                    if report_no in music_reports: # For music reports, check for LDR 06 c, d, or j
                        if bib_dict[compare_col_2][0] == 'c' or bib_dict[compare_col_2][0] == 'd' or bib_dict[compare_col_2][0] == 'j':
                            music_rows.append(row) # Put in music report
                        elif any('libretto' in row_string.lower() for row_string in row): # check if any lower-cased string in the list = "libretto"
                            music_rows.append(row) # Put in music report 
                        elif report_no in net_reports:
                            if ('NET' in bib_dict[compare_col_2][2]) and (';' not in bib_dict[compare_col_2][2]):
                                net_rows.append(row)                        
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

        if enhanced_rows or music_rows or net_rows or no_replace_rows:
            enhanced_dict[file[:-4] + '_enhanced.csv'] = enhanced_rows # Create enhanced report

            if music_rows: # Create music report
                music_rows.insert(0, enhanced_rows[0])
                enhanced_dict[file[:-4] + '_enhanced_music.csv'] = music_rows
                
            if net_rows: # Create music report
                net_rows.insert(0, enhanced_rows[0])
                enhanced_dict[file[:-4] + '_enhanced_net.csv'] = net_rows

            if no_replace_rows: # Create music report
                no_replace_rows.insert(0, enhanced_rows[0])
                enhanced_dict[file[:-4] + '_enhanced_noreplace.csv'] = no_replace_rows
            # TO DO: Row numbers are not correct for separated reports; this should be fixed if it matters to the processing team
        else:
            print file, 'was not enhanced'
                
            
# Create CSV files for enhanced reports
            
# Create new working directory and change to that directory. If directory already exists for that date, change to overflow directory
try:
    os.mkdir('C:\\Users\\beckett\\MARS Filtered Reports\\MARS Enhanced Reports ' + report_date)
    os.chdir('C:\\Users\\beckett\\MARS Filtered Reports\\MARS Enhanced Reports ' + report_date)
except:
    os.chdir('C:\\Users\\beckett\\MARS Filtered Reports\\MARS Enhanced Reports ' + report_date)
    print report_date + ' directory already exists! Enhanced reports may be duplicated in ' + os.getcwd()  

for csv_file, csv_data in enhanced_dict.items():
    with open(csv_file, 'wb') as output:
            output.write(codecs.BOM_UTF8)
            writer = csv.writer(output, quoting=csv.QUOTE_ALL,quotechar='"')
            writer.writerows(csv_data)
    print csv_file, 'has been created'


