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
from pyquery import PyQuery as pq # pip install pyquery

print sys.argv[0], 'is running ...'

# Dictionary of reports to be processed
# reports = {'R04':[],'R06 LC_Subjects':[], 'R07 LC_Subjects':[], 'R13':[], 'R14':[], 'R25':[]} # Six original project reports
reports = {'R03_C1XX 001':[]}
# reports = {'R00':[], 'R03_C1XX':[], 'R06 Series':[], 'R09 LC_Subjects': [], 'R11':[], 'R14':[], 'R39':[]} # Test reports for November 2014

#October test: R06 Series (30 rows), R09 (no report); R11 (16 rows), R14 (157 rows), R39 (600 rows--300 old/new pairs)
# R00 (2084 rows), 
#October old reports: R06 Subjects (891 rows), R07 Subjects (317 rows), R13 (288 rows), R14 (157 rows), R25 (198 rows)

# Locate most recent reports
base_url = 'http://lms01.harvard.edu/mars-reports/'

r = requests.get(base_url)
d = pq(r.content)
month_url = base_url + d('a')[0].text # Change index number to get reports for an earlier month
report_date = datetime.datetime.strptime(d('a')[0].text[:-5], '%b-%y').strftime('%Y_%m') # Convert date in URL to numeric date for file naming later in script

r = requests.get(month_url)
d = pq(r.content)

links = []
# Iteration was not working as expected with PyQuery object so items have been converted to a Python list
for link in d('a').items():
	links.append(link.text())

# Retrieve HTML pages or XLS files for each report
for report in reports.keys():
	report_pages = []
	if report.startswith('R00'):
		file_extension = '.xls'
	else:
		file_extension = '.htm'

	report_name = report.split() # Split report name on space
	# report_name[0] is BSLW report number; report_name[1], etc. is descriptive part of report name
	if len(report_name) < 2:
		report_name.append('') # If name is report number only, add a blank value for descriptive part to avoid error messages

	for link in links:
		if link.startswith(report_name[0] + '_') and link.endswith(file_extension) and report_name[1] in link:
			r = requests.get(month_url + '/' + link)
			if file_extension == '.xls': #PyQuery does not work with XLS files; use xlrd instead
				report_pages.append(xlrd.open_workbook(file_contents=r.content))
			else:
				report_pages.append(pq(r.content))

	reports[report] = report_pages # Add all pages to reports dictionary
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

    		# R06, R07, R09, R11, R13, R14, R25, R39 -- "field-info" table reports
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
    						tr_rows[-1] += [div('span.tag').text()]
    						tr_rows[-1] += [div('span.ind').text()]
    						tr_rows[-1] += [div('span.fielddata').text()]
    				rows.append(tr_rows)

	reports[report] = rows # Add reformatted data to reports dictionary
	# for row in rows: # test
	# 	print row # test 

# Filter report data (based on MARS processing team requests)
for report, lines in reports.items():
	filtered_lines = []

	if lines:
		if report.startswith('R00'):
			# Keep rows that have a match number greater than 90%
			filtered_lines.append(['Row No','Bib No','Tag','Ind','Unmatched Heading','Match %-1','Tag-1','Near Match-1',
			'Authority-1','Match %-2','Tag-2','Near Match-2','Authority-2']) # Add header
			for line in lines:
				if float(line[5][:-1]) >= 90 or float(line[9][:-1]) >= 90: # For 90% or greater matches
					del line[0] # Remove BSLW row number
					filtered_lines.append(line)
		elif report.startswith('R03_C1XX'):
			filtered_lines.append(['Row No','Tag 010 (O)','Ctl No (O)','Tag (O)','Ind (O)','Field (Old)',
			'Tag 010 (N)','Ctl No (N)','Tag (N)','Ind (N)','Field (New)','Assigned To','Notes','For Amy']) # Add header
			for record_lines in lines:
				changes = 0
				differences = 0
				changed_records = []
				zip_1 = zip(record_lines[0], record_lines[2])
				zip_2 = zip(record_lines[1], record_lines[3])
				if 'chgfield' in zip_1[1] or 'chgfield' in zip_2[1]:
					changes += 1
				if zip_1[4][0] != zip_1[4][1] or zip_2[4][0] != zip_2[4][1]:
					differences += 1
				if changes > 0 and differences > 0:
					for index, record_line in enumerate(record_lines):
						# TO DO: THIS NEEDS A LOT OF WORK
						if record_line[3] == '':
							del record_line[3]
						del record_line[1]
						del record_line[0]
						record_line += ['','','']
						if index < 3:
							changed_records += record_line
						else:
							changed_records += record_line
							filtered_lines.append(changed_records)

		elif report.startswith('R06') and 'Series' in report:
			# Keep rows with 'needs review' (i.e., not bold) data
			# TO DO: Confirm that this report does not contain invalid or partly valid fields like other R06 reports
			filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data','Needs Review','Assigned To','Notes','For Amy']) # Add header
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
			filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data','Invalid','Partly Valid','Assigned To','Notes','For Amy']) # Add header
			for line in lines:
				if line[5] != '' or line[6] != '': # Skip rows that have neither invalid nor partly valid fields
					del line[4] # Delete valid (bold) data column
					line += ['','',''] # Add blank ('Assigned To', 'Notes', and 'For Amy') columns
					filtered_lines.append(line)
		elif report.startswith('R07') and 'LC_Subjects' in report:
			# Only keep 650 fields
			# TO DO: Are we still removing everything but the 650 fields?
			filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data','Assigned To','Notes','For Amy']) # Add header
			for line in lines:
				if line[1] == '650': # Only keep 650s
					filtered_lines.append(line)
		elif report.startswith('R25'):
			# Only keep rows that have an unrecognized $z
			filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data','Unrecognized $z','Assigned To','Notes','For Amy']) # Add header
			for line_index, line in enumerate(lines):
				if line[4] != '': # Skip rows without unrecognized $z
					lines[line_index].insert(-1, '') # Add blank ("For Amy") column to end
					filtered_lines.append(line)
		else: 
			if lines[0][1] == 'Old': # For summary (e.g. Old/New) reports
			# TO DO: Do we want old and new on separate rows or single rows?
			# They are separate for now
				filtered_lines.append(['Row No','Bib No', 'Old/New','Tag','Ind','Field Data','Assigned To','Notes','For Amy']) # Add header
				for line_index, line in enumerate(lines):
					if line[0] is None:
						del line[0]
						line.insert(0, lines[line_index-1][0])
					filtered_lines.append(line)
			else: # All other reports (e.g., R09, R11, R13, R14)
				filtered_lines.append(['Row No','Bib No','Tag','Ind','Field Data','Assigned To','Notes','For Amy']) # Add header
				filtered_lines += lines

		for line_index, line in enumerate(filtered_lines): # Add row numbers
			if line_index > 0:
				if line[1] == 'New' or line[1] == 'Old': # Do not number summary reports with Old/New lines
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