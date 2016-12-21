#!/usr/bin/env python
'''
Script for converting MARS report HTML pages (or Excel files for R00) to CSV files.
Created for the Harvard Library ITS MARS Reports Pilot Project, 2014.
'''
import codecs
import csv
import datetime
import requests # pip install requests
import StringIO
import sys
import xlrd # pip install xlrd
import os
from pyquery import PyQuery as pq # pip install pyquery

print sys.argv[0], 'is running ...' # produces status message, where sys.argv[0] is the script currently running

# Dictionary of reports to be processed
##reports = {'R04':[],'R06 LC_Subjects':[], 'R07 LC_Subjects':[], 'R13':[], 'R14':[], 'R25':[]} # Six original project reports
##reports = {'R00':[], 'R03_C1XX':[], 'R04':[], 'R06 LC_Subjects': [], 'R06 Series':[], 'R07 LC_Subjects': [], 'R09 LC_Subjects': [], 'R11':[], 'R14':[], 'R28 LC_Subjects':[], 'R39':[], 'R42':[], 'R119':[]} # Test reports for November 2014
##reports = {'R00':[], 'R03_C1XX':[], 'R04':[], 'R06 LC_Subjects': [], 'R07 LC_Subjects': [], 'R09 LC_Subjects': [], 'R11':[], 'R13':[], 'R14':[], 'R25':[], 'R87':[], 'R119':[]} #  Reports for April 2015 (added R119)
##reports = {'R03_C1XX':[], 'R04':[], 'R06 LC_Subjects': [], 'R07 LC_Subjects': [], 'R09 LC_Subjects': [], 'R11':[], 'R13':[], 'R14':[], 'R25':[], 'R87':[]} #  Current reports set - February 2016 forward
reports = {'R03_C1XX':[], 'R04':[], 'R13':[], 'R14':[]} #  Reports for processing - May 2016 forward

# Locate most recent reports
base_url = 'http://lms01.harvard.edu/mars-reports/' # top-level directory page

r = requests.get(base_url)
d = pq(r.content)
# To Do: Add error checking in case top link is not a batch of monthly reports; currently assumes link contains date
month_url = base_url + d('a')[0].text  # Gets first linked url from page; change index number to get reports for an earlier month
report_date = datetime.datetime.strptime(d('a')[0].text[:-5], '%b-%y').strftime('%Y_%m') # Convert date in URL to numeric date for file naming later in script

# Create new working directory and change to that directory. If directory already exists for that date, change to overflow directory
try:
    os.mkdir('I:\MARS\MARS Reports\MARS Filtered Reports\\MARS Filtered Reports ' + report_date)
    os.chdir('I:\MARS\MARS Reports\MARS Filtered Reports\\MARS Filtered Reports ' + report_date)
except:
    os.chdir('I:\MARS\MARS Reports\MARS Filtered Reports\\MARS Filtered Reports Overflow')
    print report_date + ' directory already exists! Reports saved in: ' + os.getcwd()                

r = requests.get(month_url)
d = pq(r.content)

links = [] # list will get urls of all reports for most recent month
# Iteration was not working as expected with PyQuery object so items have been converted to a Python list
for link in d('a').items():
        links.append(link.text())

# Retrieve HTML pages or XLS files for each report
for report in reports.keys(): # uses report names generated from the key values from the dictionary of reports
        report_pages = []
        if report.startswith('R00'):
                file_extension = '.xls'
        else:
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

                if report.startswith('R00'):
                        sheet_rows = []
                        for sheet_name in page.sheet_names():
                                sh = page.sheet_by_name(sheet_name)
                                for row in range(sh.nrows):
                                        if sh.row_values(row)[-1] != '': # Skip any BSLW header/subheader rows
                                                sheet_rows.append(sh.row_values(row))
                        header = sheet_rows[0]
                        rows += sheet_rows[1:]
                                # For various reasons (including bad HTML that was not parsing correctly), it is better to work with the XLS version of this report
                                # In older reports, R00 HTML and XLS reports contained different data; this may have been fixed--but watch out for inconsistencies

                # R06, R07, R09, R11, R13, R14, R25, R39, R42 -- "field-info" table reports
                elif page('table:last').attr['class'].startswith('field-info'):
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

                # R02, R03, R03 C1XX, R04 -- "rec-compare" table reports
                elif page('table').attr['class'].startswith('rec-compare'):
                        for tr in page('table.rec-compare').find('tr.rec-compare-row').items():
                                tr_rows = []
                                for td in tr('td').items():
                                        for div in td('div').items():
                                                tr_rows.append([td.attr['class'], div.attr['class']])
                                                tr_rows[-1] += [div('span.tag').text()] #extact content based on html tag
                                                tr_rows[-1] += [div('span.ind').text()]
                                                tr_rows[-1] += [div('span.fielddata').text()]
                                rows.append(tr_rows)

        reports[report] = rows # Add reformatted data to reports dictionary

# Filter report data (based on MARS processing team requests)
for report, lines in reports.items():
        filtered_lines = []

        if lines:
                if report.startswith('R00'):
			# Keep rows that have a match number greater than 99%
                        filtered_lines.append(['Row No','Bib No','Tag','Ind','Unmatched Heading','Match %-1','Tag-1','Near Match-1',
                        'Authority-1','Match %-2','Tag-2','Near Match-2','Authority-2']) # Add header
                        for line in lines:
                                if float(line[5][:-1]) >= 99 or float(line[9][:-1]) >= 99: # For 99% or greater matches
                                        del line[0] # Remove BSLW row number
                                        #line += ['','',''] # Add blank ('Assigned To', 'Notes', and 'For Amy') columns
                                        filtered_lines.append(line)
                elif report.startswith('R03_C1XX'):
                        # Keep only rows with changes to 010 or 1XX fields; ignore indicator (and tag?) changes
                        # TO DO: Should old/new rows be on the same row or separate rows? (Currently separate.)
                        filtered_lines.append(['Old/New','Ctrl No (010)','Tag','Ind','Heading','Assigned To','Notes','# Headings Attached','Time Spent','For Amy']) # Add header
                        changed_lines = []
                        for record_lines in lines:
                                for record_line in record_lines:
                                        if record_line[2] == '010':
                                                del record_line[3] # Remove empty/unnecessary columns
                                                del record_line[2]
                                                record_line.insert(1,record_line[0].replace('-record','').title()) # Change HTML attributes to old/new
                                                del record_line[0]
                                                changed_lines.append(record_line)
                                        else:
                                                del record_line[0]
                                                changed_lines[-1] += record_line
                        for index, line in enumerate(changed_lines):
                                if index % 2 == 0: # If even row (e.g., old row)
                                        if line[1] != changed_lines[index + 1][1] or line[6] != changed_lines[index + 1][6]:
                                                del line[3] # Remove field/chgfield columns
                                                del line[1]
                                                del changed_lines[index + 1][3]
                                                del changed_lines[index + 1][1]
                                                line += ['','','','',''] # Add blank ('Assigned To', 'Notes', 'For Amy', '# Headings Attached', and 'Time Spent') columns
                                                filtered_lines.append(line)
                                                changed_lines[index + 1] += ['','','','',''] # Add blank ('Assigned To', 'Notes', 'For Amy', '# Headings Attached', and 'Time Spent') columns
                                                filtered_lines.append(changed_lines[index + 1])
                elif report.startswith('R04'):
						# Keep only the 001 and the highlighted changed field
                        # Keep No Replacement Found rows but put in separate report
                        # Use 008 byte 32 to remove undifferentiated name records; imperfect since byte is often changed from 'b' when last name is removed
						# Earlier versions of this script, used the 010 field instead of the 001; however, some records do not have an 010 field
                        # filtered_lines.append(['Old/New','Ctrl No (001)','Tag','Ind','Heading','Assigned To','Notes','For Amy']) # Add header with extra fields
                        filtered_lines.append(['Old/New','Ctrl No (010)','Tag','Ind','Heading','Assigned To','Notes','# Headings Attached','Time Spent','For Amy']) # Add header
                        for record_lines in lines: #For each old/new record pair
                                changes = 0 # counter for number of changed fields
                                changed_lines = []
                                for record_line in record_lines:
                                        if record_line[2] == '008':
                                                undiff = record_line[4][32] # Get 008 byte 32 ('b' = undifferentiated name)
                                        elif record_line[2] == '001':
                                                changed_lines.append(record_line)
                                        elif record_line[1] == 'chgfield':
                                                changes += 1
                                                changed_lines[-1] += record_line
                                if undiff != 'b':
                                        if changes == 1:
                                                changed_lines.append(['new-record','','','','No replacement found','','','','',''])
                                                filtered_lines += changed_lines
                                        elif changes == 2:
                                                filtered_lines += changed_lines
                        for index, line in enumerate(filtered_lines):
                                if index > 0:
                                        del line[6] # Remove field/chgfield columns and other unnecessary columns
                                        del line[5]
                                        del line[3]
                                        del line[2]
                                        del line[1]
                                        line.insert(1,line[0].replace('-record','').title()) # Change HTML attribute to old/new
                                        del line[0]
                                        line += ['','','']  # Add blank ('Assigned To', 'Notes', and 'For Amy') columns # TO DO: is this creating extraneous columns?                  
                elif report.startswith('R06') and 'Series' in report:
                        # Keep rows with 'needs review' (i.e., not bold) data
                        # TO DO: Confirm that this report does not contain invalid or partly valid fields like other R06 reports
                        filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data','Needs Review']) # Add header
                        for line_index, line in enumerate(lines):
                                line.insert(5,line[3].replace(line[4],'').strip()) # Get 'needs review' fields by removing bold fields
                                # Assumption is that R06 series only has 'needs review' fields, not invalid and partly valid
                                # This assumption may be wrong
                                del line[4] # Remove column with bold (valid) fields
                                line += [''] # Add blank ('For Amy') column
                                filtered_lines.append(line)
                elif report.startswith('R06'):
                        # Only keep rows that have invalid or partly valid fields
                        # TO DO: Do we want rows with 'unknown/may need review' (i.e., not bold) data also?
                        filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data','Invalid','Partly Valid']) # Add header
                        for line in lines:
                                if line[5] != '' or line[6] != '': # Skip rows that have neither invalid nor partly valid fields
                                        del line[4] # Delete valid (bold) data column
                                        line += ['','',''] # Add blank ('Assigned To', 'Notes', and 'For Amy') columns
                                        filtered_lines.append(line)
                elif report.startswith('R07') and 'LC_Subjects' in report:
                        # Only keep 650 fields
                        # TO DO: Are we still removing everything but the 650 fields?
                        filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data']) # Add header
                        for line in lines:
                                if line[1] == '650': # Only keep 650s
                                        filtered_lines.append(line)
                
                elif report.startswith('R14'):
                        #filter out 246 fields and 830 fields w/o $v
                        filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data']) # Add header
                        for line in lines:
                                if not (line[1] == '246' or ((line[1] == '830' or line[1] == '440') and '$v' not in line[3])):
                                        filtered_lines.append(line)
                            
                elif report.startswith('R25'):
                        # Only keep rows that have an unrecognized $z
                        filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data','Unrecognized $z']) # Add header
                        for line_index, line in enumerate(lines):
                                if line[4] != '': # Skip rows without unrecognized $z
                                        lines[line_index].insert(-1, '') # Add blank ("For Amy") column to end
                                        filtered_lines.append(line)
                else: 
                        if lines[0][1] == 'Old': # For summary (e.g. Old/New) reports
                        # TO DO: Should old/new rows be on the same row or separate rows? (Currently separate.)
                                filtered_lines.append(['Bib No', 'Old/New','Tag','Ind','Field Data']) # Add header
                                for line_index, line in enumerate(lines):
                                        if line[0] is None:
                                                del line[0]
                                                line.insert(0, lines[line_index-1][0])
                                        filtered_lines.append(line)
                        else: # All other reports (e.g., R09, R11, R13, R14)
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
