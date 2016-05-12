# -*- coding: utf-8 -*-
#!/usr/bin/env python
"""
Created on Mon May 09 15:14:49 2016
Script for converting MARS report HTML pages to CSV files.
Created for the Harvard Library ITS MARS Reports Project, 2016.
@author: beckett
"""

import codecs
import csv
import datetime
import requests # pip install requests
import StringIO
import sys
import xlrd # pip install xlrd
from pyquery import PyQuery as pq # pip install pyquery

print sys.argv[0], 'is running ...' # produces status message, where sys.argv[0] is the script currently running

# Dictionary of reports to be processed

##reports = {'R03_C1XX':[], 'R04':[], 'R06 LC_Subjects': [], 'R07 LC_Subjects': [], 'R09 LC_Subjects': [], 'R11':[], 'R13':[], 'R14':[], 'R25':[], 'R87':[]} #  Current reports set - February 2016 forward
reports = {'R13':[], 'R14':[]} #  Reports needed from archive for retrospective processing - May 2016 forward

# Locate reports
base_url = 'http://lms01.harvard.edu/mars-reports/' # top-level directory page

r = requests.get(base_url)
d = pq(r.content)

for month_index in range (5,10):
    # To Do: Add error checking in case top link is not a batch of monthly reports; currently assumes link contains date
    month_url = base_url + d('a')[month_index].text  # Gets first linked url from page; change index number to get reports for an earlier month
    report_date = datetime.datetime.strptime(d('a')[month_index].text[:-5], '%b-%y').strftime('%Y_%m') # Convert date in URL to numeric date for file naming later in script
    
    r = requests.get(month_url)
    d = pq(r.content)
    
    links = [] # list will get urls of all reports for most recent month
    # Iteration was not working as expected with PyQuery object so items have been converted to a Python list
    for link in d('a').items():
            links.append(link.text())
    
    # Retrieve HTML pages or XLS files for each report
    for report in reports.keys(): # uses report names generated from the key values from the dictionary of reports
            report_pages = []
          
            file_extension = '.htm'
    
            report_name = report.split() # Split report name on space
            # report_name[0] is BSLW report number; report_name[1], etc. is descriptive part of report name
            if len(report_name) < 2:
                    report_name.append('') # If name is report number only, add a blank value for descriptive part to avoid error messages
    
            for link in links: # match report from production report list to url(s) from list of links from monthly report page with correct extension
                    if link.startswith(report_name[0] + '_') and link.endswith(file_extension) and report_name[1] in link:
                            r = requests.get(month_url + '/' + link) # get individual report page content
                            if file_extension == '.xls': #PyQuery does not work with XLS files; use xlrd instead
                                    report_pages.append(xlrd.open_workbook(file_contents=r.content)) # append content to list of report page content
                            else:
                                    report_pages.append(pq(r.content)) # append content to list of report page content
    
            reports[report] = report_pages # Add all pages (from list of report page content) to reports dictionary
            # TO DO: Add caching to above code so multiple calls will not be needed if script fails  
    
# Extract data from the report pages
for report, pages in reports.items(): # using report, pages for key, value
    rows = []
    # header = [] # TO DO: Find way (other than manually) to get header data

    if pages:
        for page in pages:

                # R06, R07, R09, R11, R13, R14, R25, R39, R42 -- "field-info" table reports
                if page('table:last').attr['class'].startswith('field-info'):
                        for tr in page('table.field-info').find('tr').items():
                                if not tr('th'): # Skip any BSLW header/legend rows
                                        if len(tr('.ctl_no')) > 1: # For reports with multiple ctl_no fields (e.g., summary reports with old/new columns)
                                                for ctl_index, ctl in enumerate(tr('.ctl_no')):
                                                        if ctl_index == 0:
                                                                rows.append([ctl.text])
                                                        else: # For reports with a single ctl_no field (usually the Aleph bib. no.)
                                                                rows[-1] += [ctl.text]
                                        else:
                                                rows.append([tr('.ctl_no').text()])
                                        rows[-1] += [tr('.tag').text()]
                                        rows[-1] += [tr('.ind').text()]
                                        rows[-1] += [tr('.fielddata').text()]
                                        rows[-1] += [tr('.fielddata b').text()] #R25 (unrecognized $z), R06 (data matched authority record)
                                        rows[-1] += [tr('.fielddata span.invalid').text()] #R06
                                        rows[-1] += [tr('.fielddata span.partly_valid').text()] #R06



        reports[report] = rows # Add reformatted data to reports dictionary

# Filter report data (based on MARS processing team requests)
for report, lines in reports.items():
        filtered_lines = []

        if lines:
            
                filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data']) # Add header
                filtered_lines += lines

                for line_index, line in enumerate(filtered_lines): # Add row numbers
                        if line_index > 0:
                                if line[1] == 'New' or line[1] == 'Old': # Do not number summary reports with old/new lines
                                        pass
                                elif line[0] == 'New' or line[0] == 'Old': # Also do not number authority reports with old/new lines (e.g., R03 C1XX)
                                        pass
                                else:
                                        filtered_lines[line_index].insert(0, str(line_index))

              # Create CSV File
                out_file = report.replace(' ','_') + '_' + report_date + '.csv'
                with open(out_file,'wb') as output:
                        output.write(codecs.BOM_UTF8)
                        writer = csv.writer(output,quoting=csv.QUOTE_ALL,quotechar='"')
                        for line in filtered_lines:
                                writer.writerow([i.encode('utf-8') for i in line])
                print out_file, 'created'
        else:
              print 'No', report, 'report for this month' # TO DO: This would be better as part of a log file
              # TO DO: Log could also give info about number of rows in each report
