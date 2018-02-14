from __future__ import print_function 
import sys 
import os, errno 
import re

from lxml import etree 
import xml.etree.ElementTree as et 
import lxml.etree as et

from xmlutils.xml2csv import xml2csv
import glob
import csv
import xlrd
import xlwt

print("XML converter version 1.3")
print("Please follow the directions below")
print("For technical assistance contact mcohen@krl.org")
print("")
print("Input filename should include source as well")
print("Example format: /Users/bill/Desktop/quartertwo.xml")
print("")
fileIn=raw_input("Type input filename:")
print("")
print("Output path is just the path not filename")
print("Example format: /Users/bill/Desktop/")
print("")
fileOut=raw_input("Type output destination:")

###### Start XLS ######
globeXLS = '''\
	<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
		<xsl:output method="xml" indent="yes" omit-xml-declaration="no"/>
		<xsl:strip-space elements="*"/>
		<xsl:template match="node()|@*">
			<xsl:copy>
				<xsl:apply-templates select="node()|@*"/>
			</xsl:copy>
		</xsl:template>
		<xsl:key name="k" match="event" use="concat(title, '|', RelatedLocations)"/>
		<xsl:template match="events">
			<xsl:copy>
				<xsl:for-each select="event[count(. | key('k', concat(title, '|', RelatedLocations))[1]) = 1]">
					<xsl:sort select="title" />
					<event>
						<xsl:apply-templates select="EventType" />
						<title>
						<xsl:value-of select="title" />
						</title><title></title>
						<xsl:for-each select="key('k', concat(title, '|', RelatedLocations))">
							<xsl:sort select="RelatedLocations" />
							<RelatedLocations>
							<xsl:value-of select="RelatedLocations" />
							</RelatedLocations>
						</xsl:for-each>
						<xsl:apply-templates select="Date" />                
						<xsl:apply-templates select="DateYear" />
						<xsl:apply-templates select="DateMonth" />
						<xsl:apply-templates select="DateDay" />
						<DateDay>NA</DateDay>
						<xsl:apply-templates select="Body" />
						<Body></Body>
						<xsl:apply-templates select="AgeRanges" />
						<AgeRanges>NA</AgeRanges>
						<xsl:apply-templates select="RegistrationRequired" />
						<RegistrationRequired></RegistrationRequired>
						<xsl:apply-templates select="Location" />
						<Location>NA</Location>
					</event>                
				</xsl:for-each>
			</xsl:copy>		
		</xsl:template>
		<xsl:template match="*/text()[normalize-space()]">
			<xsl:value-of select="normalize-space()"/>
		</xsl:template>
		<xsl:template match="*/text()[not(normalize-space())]" />
	</xsl:stylesheet> '''
	
storyXLS = '''\
		<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
		<xsl:output encoding="UTF-8" indent="yes" method="xml" />
		<xsl:strip-space elements="*"/>
			<xsl:template match="node()|@*">
				<xsl:copy>
					<xsl:apply-templates select="node()|@*">
						<xsl:sort select="RelatedLocations" />
					</xsl:apply-templates>
				</xsl:copy>
			</xsl:template>
			<xsl:template match="event">
				<xsl:copy>
					<xsl:apply-templates select="@*" />
					<xsl:apply-templates select="RelatedLocations" />
					<xsl:apply-templates select="Date" />
					<xsl:apply-templates select="title" />
					<xsl:apply-templates select="DateYear" />
					<xsl:apply-templates select="DateMonth" />
					<xsl:apply-templates select="DateDay" />
				</xsl:copy>
			</xsl:template> 
		</xsl:stylesheet> ''' 

kidsXLS = '''\
		<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
			<xsl:output method="xml" indent="yes" omit-xml-declaration="no"/>
			<xsl:strip-space elements="*"/>
			<xsl:template match="node()|@*">
				<xsl:copy>
					<xsl:apply-templates select="node()|@*"/>
				</xsl:copy>
			</xsl:template>
			<xsl:key name="k" match="event" use="concat(title, '|', RelatedLocations)"/>
			<xsl:template match="events">
				<xsl:copy>
					<xsl:for-each select="event[count(. | key('k', concat(title, '|', RelatedLocations))[1]) = 1]">
						<xsl:sort select="title" />
						<event>
							<xsl:apply-templates select="EventType" />
							<title>
							<xsl:value-of select="title" />
							</title>
							<xsl:for-each select="key('k', concat(title, '|', RelatedLocations))">
								<xsl:sort select="RelatedLocations" />
								<RelatedLocations>
								<xsl:value-of select="RelatedLocations" />
								</RelatedLocations>
							</xsl:for-each>
							<xsl:apply-templates select="Date" />                
							<xsl:apply-templates select="DateYear" />
							<xsl:apply-templates select="DateMonth" />
							<xsl:apply-templates select="DateDay" />
							<DateDay>NA</DateDay>
							<xsl:apply-templates select="Body" />
							<Body></Body>
							<xsl:apply-templates select="AgeRanges" />
							<AgeRanges>NA</AgeRanges>
							<xsl:apply-templates select="RegistrationRequired" />
							<RegistrationRequired></RegistrationRequired>
							<xsl:apply-templates select="RecommendedFor" />
							<RecommendedFor>NA</RecommendedFor>
							<xsl:apply-templates select="Location" />
							<Location>NA</Location>
						</event>                
					</xsl:for-each>
				</xsl:copy>		
			</xsl:template>
			<xsl:template match="*/text()[normalize-space()]">
				<xsl:value-of select="normalize-space()"/>
			</xsl:template>
			<xsl:template match="*/text()[not(normalize-space())]" />
		</xsl:stylesheet> '''
		
sortXLS = '''\
	<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output encoding="UTF-8" indent="yes" method="xml" />
	<xsl:strip-space elements="*"/>
		<xsl:template match="node()|@*">
			<xsl:copy>
				<xsl:apply-templates select="node()|@*">
					<xsl:sort select="EventType" />
				</xsl:apply-templates>
			</xsl:copy>
		</xsl:template>
	</xsl:stylesheet> '''
###### End XLS ######

###### Global Functions Start #######
def silentRemove(filename):
    try:
        os.remove(filename)
    except OSError as e:
        if e.errno != errno.ENOENT: # errno.ENOENT = no such file or directory
            raise silentRemove(filename)

def silentCreate(filename):
	events = et.Element("events") #starts the file with <events> and ends with </events> because the rest of the script creates only <event> to </event>
	tree = et.ElementTree(events)
	tree.write(filename) 
	
silentCreate("temp.xml") #creates the master temporary XML file

def globalClean(): #Cleans up some wording and formatting
	src_tree = et.parse(fileIn) #parse source
	src_root = src_tree.getroot() #Get's the root of each node i.e. everything between <event> and </event>
	for event in src_root.findall('event'): #Basic start to most of these functions. Locates the node that starts with <event> i.e. all of them
			month = ['January', 'February', 'August', 'September', 'October', 'November', 'December']
			monthAbrv = ['Jan.', 'Feb.', 'Aug.', 'Sept.', 'Oct.', 'Nov.', 'Dec.']
			date = event.find('Date') #Checks for child note <Date>
			dates = date.text #Changes that note into a text string
			title = event.find('title')
			titles = title.text
			body = event.find('Body')
			if body is None: #If the node is empty this will pass over that node instead of throwing an error
				continue
			bodies = body.text		
			for month, monthAbrv in zip(month, monthAbrv):
				if month in dates:
					date.text = date.text.replace(month, monthAbrv)
				if month in bodies:
					body.text = body.text.replace(month, monthAbrv)
			if 'p.m.-' in dates:
				date.text = date.text.replace('p.m.-', '-')
			if '2017' in dates:
				date.text = date.text.replace(' 2017 -', '')
			if ' -' in dates:
				date.text = date.text.replace(' -', '-')
			if '12 p.m.' in dates:
				date.text = date.text.replace('12 p.m.', 'noon')
			if '  ' in dates:
				date.text = date.text.replace('  ', ' ')
			if '&#039;' in titles:
				title.text = title.text.replace('&#039;', '\'')
			if '  ' in bodies:
				body.text = body.text.replace('  ', ' ')
			if '&#039;' in bodies:
				body.text = body.text.replace('&#039;', '\'')
			location = event.find('Location')
			if location is None:
				continue
			locations = location.text
			if '&#039;' in locations:
				location.text = location.text.replace('&#039;', '\'')		
	et.ElementTree(src_root).write("temp.xml") #Writes out to temp.xml which is the base of everything
globalClean()

def ageGrab(age, filename): #Grabs all events related to a certain age group
	dest_tree = et.parse("temp.xml") #Parses temp.xml
	dest_root = dest_tree.getroot() 
	for event in dest_root.findall('event'):
		agerange = event.find('AgeRanges')
		if agerange is None: 
			dest_root.remove(event)
			continue
		ageranges = agerange.text
		if ageranges != age:
			dest_root.remove(event) #.remove will delete all events that don't contain whatever AgeRange we're looking for i.e. adult, child etc
	et.ElementTree(dest_root).write(filename)	
	
def clean(filename): #Time to get rid of any element with some tags
	dest_tree = et.parse(filename)
	dest_root = dest_tree.getroot()
	for event in dest_root.findall('event'):
		book = event.find('EventType') #Find all EventType tags in the new destination source
		books = book.text
		if books == 'Book Groups': #anything that has this in the EventType tag is gone
			dest_root.remove(event)
		elif books == 'Book Sales':
			dest_root.remove(event)
		elif books == 'Bookmobile Stop':
			dest_root.remove(event)
		elif books == 'Friends of the Library':
			dest_root.remove(event)
		elif books == 'Friends of the Library Book Sale':
			dest_root.remove(event)
		elif books == 'Storytimes':
			dest_root.remove(event)
		agerange = event.find('AgeRanges')
	et.ElementTree(dest_root).write(filename)
	
def dateClean(filename): #Changing dates from numbers to APA style
	src_tree = et.parse(filename) #parse source
	src_root = src_tree.getroot()
	for event in src_root.findall('event'):
		monthNum = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
		month = ['Jan.', 'Feb.', 'March', 'April', 'May', 'June', 'July', 'Aug.', 'Sept.', 'Oct.', 'Nov.', 'Dec.']
		date = event.find('DateMonth')
		dates = date.text
		for monthNum, month in zip(monthNum, month):
			if monthNum in dates:
				date.text = date.text.replace(monthNum, month)	
	et.ElementTree(src_root).write(filename)	
	
def cleanNodes(filename): #Deleting the notes field
	dest_tree = et.parse(filename)
	dest_root = dest_tree.getroot()
	foos = dest_tree.findall('event') 
	for event in foos:
		bars = event.findall('Notes') 
		for Notes in bars: 
			event.remove(Notes) 
	et.ElementTree(dest_root).write(filename)

def change(filename, xlsfile): #Used to change and sort XML files
	dom = et.parse(filename) #Same parse as before
	xslt_root = etree.XML(xlsfile) #Reads the XLS code
	transform = et.XSLT(xslt_root) #Sets XLS code to be called
	newdom = transform(dom) #Transforming our .xml file with our XLS code
	log = open(filename, 'w') #Opens the temp*.xml file as a writeable
	print(str(newdom), file = log) #Writes out to the tempfile
	
def deleteDupes(filename): #Deletes duplicate XML nodes
	path = (filename)
	tree = et.parse(path)
	root = tree.getroot() 
	prev = None
	def elements_equal(e1, e2):
		if type(e1) != type(e2):
			return False
		if e1.tag != e1.tag: return False
		if e1.text != e2.text: return False
		if e1.tail != e2.tail: return False
		if e1.attrib != e2.attrib: return False
		if len(e1) != len(e2): return False
		return all([elements_equal(c1, c2) for c1, c2 in zip(e1, e2)])

	for page in root:
		elems_to_remove = []
		for elem in page:
			if elements_equal(elem, prev):
				elems_to_remove.append(elem)
				continue
			prev = elem
		for elem_to_remove in elems_to_remove:
			page.remove(elem_to_remove)

	et.ElementTree(root).write(filename)
	
def finalTouch(filename): #Writes XML header in temp file
	tree = et.parse(filename)
	root = tree.getroot()
	et.ElementTree(root).write(filename, xml_declaration=True, method='xml', encoding='UTF-8')
###### Global Functions End ######

###### Create Temporary XML Files ######
silentCreate("tempadult.xml")
silentCreate("tempkids.xml")
silentCreate("tempteen.xml")
silentCreate("tempstory.xml")
silentCreate("tempbook.xml")
silentCreate("tempfriend.xml")

###### Start of Individual Functions ######

###### Full Adult Tab Function ######
def adultSmash():
	ageGrab("Adult", "tempadult.xml")
	clean("tempadult.xml")
	cleanNodes("tempadult.xml")
	dateClean("tempadult.xml")
	change("tempadult.xml", globeXLS)
	change("tempadult.xml", sortXLS)
	deleteDupes("tempadult.xml")
	finalTouch("tempadult.xml")
adultSmash()

###### Kids and Teens ######
def ageSmash():
	array = ["Kids", "Teen"]
	array2 = ["tempkids.xml", "tempteen.xml"]
	for array, array2 in zip(array, array2):
		ageGrab(array, array2)
		clean(array2)
		cleanNodes(array2)
		dateClean(array2)
		change(array2, kidsXLS)
		change(array2, sortXLS)
		deleteDupes(array2)
		finalTouch(array2)
ageSmash()

###### Book Groups ######
def bookSmash():
	def bookGrab():
		dest_tree = et.parse("temp.xml")
		dest_root = dest_tree.getroot()
		for event in dest_root.findall('event'):
			agerange = event.find('EventType')
			if agerange is None: 
				dest_root.remove(event)
				continue
			ageranges = agerange.text
			if ageranges != 'Book Groups':
				dest_root.remove(event)
		et.ElementTree(dest_root).write("tempbook.xml")	
	bookGrab()
	def bookFix(): #Changing "Book Groups:" in titles to RelatedLocations
		dest_tree = et.parse("tempbook.xml")
		dest_root = dest_tree.getroot()
		for event in dest_root.findall('event'):
			location = event.find('RelatedLocations')
			locations = location.text
			title = event.find('title')
			titles = title.text
			if 'Book Group:' in titles:
				title.text = title.text.replace('Book Group:', locations)
		et.ElementTree(dest_root).write("tempbook.xml")
	bookFix()
	cleanNodes("tempbook.xml")
	dateClean("tempbook.xml")
	change("tempbook.xml", globeXLS)
	deleteDupes("tempbook.xml")
	finalTouch("tempbook.xml")
bookSmash()				

###### Friends of the Library Events ######
def friendSmash():
	def friendGrab():
		dest_tree = et.parse("temp.xml")
		dest_root = dest_tree.getroot()
		for event in dest_root.findall('event'):
			eventT =["Arts", "Business", "DIY", "Films", "Games", "Genealogy", "Health", "Hot", "Trustees", "U", "Author", "Environment", "Special", "Technology", "Writing", "Groups", "Bookmobile Stop", "BiblioTEC", "Leadership", "Legos", "School", "STEM", "Storytimes"]
			for eventT in eventT:
				agerange = event.find('EventType')
				if agerange is None: 
					continue
				ageranges = agerange.text
				if eventT in ageranges:
					dest_root.remove(event)
		et.ElementTree(dest_root).write("tempfriend.xml")
	
	friendGrab()
	cleanNodes("tempfriend.xml")
	dateClean("tempfriend.xml")
	change("tempfriend.xml", globeXLS)
	change("tempfriend.xml", sortXLS)
	deleteDupes("tempfriend.xml")
	finalTouch("tempfriend.xml")
friendSmash()	

###### Storytimes ######
def storySmash():
	def storyGrab():
		dest_tree = et.parse("temp.xml") 
		dest_root = dest_tree.getroot() 
		for event in dest_root.findall('event'):
			eventT =["Arts", "Business", "DIY", "Films", "Games", "Genealogy", "Health", "Hot", "Trustees", "U", "Author", "Environment", "Special", "Technology", "Writing", "Groups", "Bookmobile Stop", "BiblioTEC", "Leadership", "Legos", "School", "STEM"]
			for eventT in eventT:
				agerange = event.find('EventType')
				if agerange is None: 
					continue
				ageranges = agerange.text
				if eventT in ageranges:
					dest_root.remove(event)
		et.ElementTree(dest_root).write("tempstory.xml")
	def storyClean(filename):
		dest_tree = et.parse(filename)
		dest_root = dest_tree.getroot()
		for event in dest_root.findall('event'):
			book = event.find('EventType')
			books = book.text
			if books == 'Book Groups':
				dest_root.remove(event)
			elif books == 'Book Sales':
				dest_root.remove(event)
			elif books == 'Bookmobile Stop':
				dest_root.remove(event)
			elif books == 'Friends of the Library':
				dest_root.remove(event)
			elif books == 'Friends of the Library Book Sale':
				dest_root.remove(event)
			agerange = event.find('AgeRanges')
		et.ElementTree(dest_root).write(filename)	
	storyGrab()
	storyClean("tempstory.xml")
	cleanNodes("tempstory.xml")
	change("tempstory.xml", storyXLS)
	dateClean("tempstory.xml")
	deleteDupes("tempstory.xml")
	finalTouch("tempstory.xml")
storySmash()

###### Converting XML to CSV files ######
def convertCSV():
	array = ["tempadult.xml", "tempkids.xml", "tempteen.xml", "tempstory.xml", "tempbook.xml", "tempfriend.xml", fileIn]
	array2 = ["tempadults.csv", "tempkids.csv", "tempteen.csv", "tempstory.csv", "tempbook.csv", "tempfriend.csv", "First_Pull.csv"]
	for array, array2 in zip(array, array2):
		converter = xml2csv(array, array2, encoding="utf-8")
		converter.convert(tag="event")
convertCSV()

###### Cleaning up CSVs to reflect proper sorting and fix column titles ######
def csvClean():
	abf = ['EventType', 'Title', 'RelatedLocation', 'Date', 'DateYear', 'DateMonth', 'DateDay', 'Body', 'AgeRanges', 'RegistrationRequired', 'Location']
	kt = ['EventType', 'Title', 'RelatedLocation', 'Date', 'DateYear', 'DateMonth', 'DateDay', 'Body', 'AgeRanges', 'RegistrationRequired', 'RecommendedFor', 'Location']
	s = ['RelatedLocation', 'Date', 'Title', 'DateYear', 'DateMonth', 'DateDay']
	array = ["tempadults.csv", "tempbook.csv", "tempfriend.csv", "tempkids.csv", "tempteen.csv", "tempstory.csv"]
	array2 = ["adults.csv", "book.csv", "friend.csv", "kids.csv", "teen.csv","story.csv"]
	array3 = [abf, abf, abf, kt, kt, s]
	for array, array2, array3 in zip(array, array2, array3):
		inputFileName = array
		outputFileName= array2	
		with open(inputFileName, 'rb') as inFile, open(outputFileName, 'wb') as outfile:
			r = csv.reader(inFile)
			w = csv.writer(outfile)

			next(r, None)  # skip the first row from the reader, the old header
			# write new header
			w.writerow(array3)

			# copy the rest
			for row in r:
				w.writerow(row)
csvClean()

###### Clean up the temporary CSVs ######
def destroy():
	array = ["tempadults.csv", "tempkids.csv", "tempteen.csv", "tempstory.csv", "tempbook.csv", "tempfriend.csv"]
	for array in array:
		silentRemove(array)		
destroy()

###### Put all the CSV's together into the output file ######
def combiner():
	wb = xlwt.Workbook(encoding='UTF-8')
	for csvfile in glob.glob(os.path.join('.', '*.csv')):
			fpath = csvfile.split("/", 1)
			fname = fpath[1].split(".", 1)
			ws = wb.add_sheet(fname[0])
			with open(csvfile, 'rb') as f:
					reader = csv.reader(f)
					for r, row in enumerate(reader):
							for c, col in enumerate(row):
									ws.write(r, c, col)
	wb.save(fileOut + 'output.xls')
combiner()

###### Get rid of all of the leftover XML and CSV files ######
def destroyer():
	array = ["tempadult.xml", "tempkids.xml", "tempteen.xml", "tempbook.xml", "tempfriend.xml", "adults.csv", "kids.csv", "teen.csv", "story.csv", "book.csv", "friend.csv", "temp.xml"]
	for array in array:
		silentRemove(array)
destroyer()
