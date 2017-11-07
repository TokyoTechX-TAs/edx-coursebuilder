# coding=utf-8

import os, tarfile, shutil, xlrd, datetime
from lxml import etree

from library.video_comp import *
from library.problem_comp import *
from library.html_comp import *



"""
	sheet->coursestructure
"""
STURCSHEET = "coursestructure"
STRUCID = 0
STRUCSECTION = 1
STRUCSUBSECTION = 2
STRUCUNIT = 3
STRUCCOMPONENT = 4
STRUCTYPECOMPONENT = 5



"""
hardcoded xlsmpath must change to a parameter
"""

course_path= 'course'
xlsmPath = "course_info.xlsm"
wb = xlrd.open_workbook(xlsmPath)
sheetstruc = wb.sheet_by_name(STURCSHEET)
sheetvideo = wb.sheet_by_name(VIDEOSHEET)
sheetproblem = wb.sheet_by_name(PROBLEMSHEET)
sheethtml = wb.sheet_by_name(HTMLSHEET)













class Course_extraction:

	def __init__(self):
				
		print (os.getcwd())

		if not os.path.exists(os.path.join(course_path,"chapter")):
			os.makedirs(os.path.join(course_path,"chapter"))
		if not os.path.exists(os.path.join(course_path,"sequential")):
			os.makedirs(os.path.join(course_path,"sequential"))
		if not os.path.exists(os.path.join(course_path,"vertical")):
			os.makedirs(os.path.join(course_path,"vertical"))
		if not os.path.exists(os.path.join(course_path,"video")):
			os.makedirs(os.path.join(course_path,"video"))
		if not os.path.exists(os.path.join(course_path,"problem")):
			os.makedirs(os.path.join(course_path,"problem"))
		if not os.path.exists(os.path.join(course_path,"html")):
			os.makedirs(os.path.join(course_path,"html"))
		if not os.path.exists(os.path.join(course_path,"static")):
			os.makedirs(os.path.join(course_path,"static"))

		self.section_path = os.path.join(course_path,'chapter')
		self.subsection_path = os.path.join(course_path,'sequential')
		self.unit_path = os.path.join(course_path,'vertical')
		self.problme_path = os.path.join(course_path,'problem')
		self.course = os.path.join(course_path,'course')
	


	def course_(self):
		section_file = os.path.join(self.course,'course.xml')
		tree = etree.parse(section_file)
		root = tree.getroot()
		section_ls = root.findall('.chapter')
		temp_url = []
		for section in section_ls:
			temp_url.append(section.get('url_name'))
		self.section_url = {'section_url':temp_url}

		return(self.section_url)


	def sections(self):
		sections_files = os.listdir(self.section_path)
		self.all_section = []
		for section_file in sections_files:
			tree = etree.parse(os.path.join(self.section_path,section_file))
			root = tree.getroot()
			section_name = root.get('display_name')
			section_link = section_file.replace('.xml', '')
			subsection_objs = root.findall(".sequential")
			subsection_url = []
			for subsection_obj in subsection_objs:	
				subsection_url.append(subsection_obj.get('url_name'))

			#print "section: " + str(block_name)
			#print subblock_url
			self.all_section.append({'section_link':section_link,
				'section_name':section_name,
				'assoc_subsection_url':subsection_url})
		return(self.all_section)

	def subsections(self):
		subsections_files = os.listdir(self.subsection_path)
		self.all_subsection = []
		for subsection_file in subsections_files:
			tree = etree.parse(os.path.join(self.subsection_path,subsection_file))
			root = tree.getroot()
			subsection_name = root.get('display_name')
			subsection_link = subsection_file.replace( '.xml', '')
			unit_objs = root.findall(".vertical")
			unit_url = []
			for unit_obj in unit_objs:	
				unit_url.append(unit_obj.get('url_name'))

			self.all_subsection.append({'subsection_link':subsection_link,
				'subsection_name':subsection_name,
				'assoc_unit_url':unit_url})
		return(self.all_subsection)


	def units(self):
		units_files = os.listdir(self.unit_path)
		self.all_unit = []
		for unit_file in units_files:
			tree = etree.parse(os.path.join(self.unit_path,unit_file))
			root = tree.getroot()
			unit_name = root.get('display_name')
			unit_link = unit_file.replace('.xml', '')
			
			
			self.all_unit.append({'unit_link':unit_link,'unit_name':unit_name})
		return(self.all_unit)
	










def create_course():
	"""
	create course.xml file
	"""

	from_course = Course_extraction()
	list_section = from_course.course_()
	tree = etree.parse(os.path.join(course_path,'course','course.xml'))
	root = tree.getroot()
	

	currentsection = ''
	section_idx = 1


	for row in range(1, sheetstruc.nrows):

		if currentsection != sheetstruc.cell_value(row, STRUCSECTION ):
			currentsection = sheetstruc.cell_value(row, STRUCSECTION )
			urlName = "section" +  "{0:0=2d}".format(section_idx)

			if urlName not in list_section['section_url']:
				print('no section: "'+ urlName +'"" in course. Add link to course.xml')
				etree.SubElement(root, 'chapter',url_name=urlName)
			else:
				print('section: "'+ urlName +'" exists in course.')

			section_idx+=1


	
	doc = etree.ElementTree(root)
	doc.write(os.path.join(course_path,'course','course.xml'), pretty_print=True, xml_declaration=False, encoding='utf-8')





def create_section():
	"""
	creates a section file
	
	"""

	currentsection = sheetstruc.cell_value(1, STRUCSECTION )
	currentsubsection = sheetstruc.cell_value(1,STRUCSUBSECTION)
	section_idx = 1
	subsection_idx = 1
	filename = 'section' +  '{0:0=2d}'.format(section_idx) + '.xml'
	page = etree.Element('chapter', display_name= currentsection)
	subsection_url_name = 'subsection' +  '{0:0=2d}'.format(subsection_idx) 
	etree.SubElement(page, 'sequential',url_name=subsection_url_name)
	subsection_idx += 1
	print('added new section: "'+ filename +'" file at chapter directory')
	print('      added new subsection link"'+ subsection_url_name +'"" in section:' +filename )

	for row in range(2, sheetstruc.nrows):

		if currentsection != sheetstruc.cell_value(row, STRUCSECTION ):

			doc = etree.ElementTree(page)
			doc.write(os.path.join(course_path,'chapter',filename), pretty_print=True, xml_declaration=False, encoding='utf-8')
			print('added new section: "'+ filename +'" file at chapter directory')
			section_idx +=1
			currentsection = sheetstruc.cell_value(row, STRUCSECTION )
			currentsubsection = sheetstruc.cell_value(row,STRUCSUBSECTION)
			filename = 'section' +  '{0:0=2d}'.format(section_idx) + '.xml'
			page = etree.Element('chapter', display_name= currentsection)
			subsection_url_name = 'subsection' +  '{0:0=2d}'.format(subsection_idx) 
			etree.SubElement(page, 'sequential',url_name=subsection_url_name)
			subsection_idx += 1
			print('      added new subsection link"'+ subsection_url_name +'"" in file: ' +filename )
			
		else:
			if currentsubsection != sheetstruc.cell_value(row,STRUCSUBSECTION):
				currentsubsection = sheetstruc.cell_value(row,STRUCSUBSECTION)
				subsection_url_name = 'subsection' +  '{0:0=2d}'.format(subsection_idx) 
				etree.SubElement(page, 'sequential',url_name=subsection_url_name)
				subsection_idx += 1
				print('      added new subsection link "'+ subsection_url_name +'"" in file: ' +filename )
			

	doc = etree.ElementTree(page)
	doc.write(os.path.join(course_path,'chapter',filename), pretty_print=True, xml_declaration=False, encoding='utf-8')
	print('added new section: "'+ filename +'" file at chapter directory')

def create_subsection():
	"""
	creates subsection files
	
	"""
	currentsubsection = sheetstruc.cell_value(1,STRUCSUBSECTION)
	currentunit = sheetstruc.cell_value(1, STRUCUNIT)
	subsection_idx = 1
	unit_idx = 1
	filename = 'subsection' +  '{0:0=2d}'.format(subsection_idx) + '.xml'
	page = etree.Element('sequential', display_name= currentsubsection)
	unit_url_name = 'unit' +  '{0:0=2d}'.format(subsection_idx) 
	etree.SubElement(page, 'vertical',url_name=unit_url_name)
	unit_idx += 1
	print('added new subsection: "'+ filename +'" file at sequential directory')
	print('      added new unit link "'+ unit_url_name +'"" in subsection:' +filename )

	for row in range(2, sheetstruc.nrows):

		if currentsubsection != sheetstruc.cell_value(row, STRUCSUBSECTION ):

			doc = etree.ElementTree(page)
			doc.write(os.path.join(course_path,'sequential',filename), pretty_print=True, xml_declaration=False, encoding='utf-8')
			print('added new subsection: "'+ filename +'" file at sequential directory')
			subsection_idx +=1
			currentsubsection = sheetstruc.cell_value(row, STRUCSUBSECTION )
			currentunit = sheetstruc.cell_value(row,STRUCUNIT)
			filename = 'subsection' +  '{0:0=2d}'.format(subsection_idx) + '.xml'
			page = etree.Element('sequential', display_name= currentsubsection)
			unit_url_name = 'unit' +  '{0:0=2d}'.format(unit_idx) 
			etree.SubElement(page, 'vertical',url_name=unit_url_name)
			unit_idx += 1
			print('      added new unit "'+  unit_url_name +'"" in file: ' +filename )
			
		else:
			if currentunit != sheetstruc.cell_value(row,STRUCUNIT):
				currentunit = sheetstruc.cell_value(row,STRUCUNIT)
				unit_url_name = 'unit' +  '{0:0=2d}'.format(unit_idx) 
				etree.SubElement(page, 'vertical',url_name=unit_url_name)
				unit_idx += 1
				print('      added new unit "'+ unit_url_name +'"" in file: ' +filename )
			

	doc = etree.ElementTree(page)
	doc.write(os.path.join(course_path,'sequential',filename), pretty_print=True, xml_declaration=False, encoding='utf-8')
	print('added new subsection: "'+ filename +'" file at sequential directory')
			




def create_unit():
	"""
	creates unit files

	"""

	currentunit = sheetstruc.cell_value(1, STRUCUNIT)
	unit_idx = 1
	filename = 'unit' +  '{0:0=2d}'.format(unit_idx) +'.xml'
	page = etree.Element('vertical', display_name= currentunit)
	print('added new unit: "'+ filename +'" file at vertical directory')
	
	for row in range(2, sheetstruc.nrows):

		if currentunit != sheetstruc.cell_value(row, STRUCUNIT):

			doc = etree.ElementTree(page)
			doc.write(os.path.join(course_path,'vertical',filename), pretty_print=True, xml_declaration=False, encoding='utf-8')
			print('added new unit: "'+ filename +'" file at vertical directory')
			unit_idx +=1
			currentunit = sheetstruc.cell_value(row, STRUCUNIT )
			filename = 'unit' +  '{0:0=2d}'.format(unit_idx) + '.xml'
			page = etree.Element('chapter', display_name= currentunit)
				

	doc = etree.ElementTree(page)
	doc.write(os.path.join(course_path,'vertical',filename), pretty_print=True, xml_declaration=False, encoding='utf-8')
	print('added new unit: "'+ filename +'" file at vertical directory')
			


def add_component():


	'''
	start adding component with respect to macro excel
	'''
	video_idx = 1
	text_idx = 1
	problem_idx = 1
	for row in range(1, sheetstruc.nrows):
		comp_type = sheetstruc.cell_value(row,STRUCTYPECOMPONENT)
		print(str(row),str(text_idx))
		if comp_type == 'video':
			search_video_in_course(video_excel2list(video_idx,sheetvideo),Course_extraction(),course_path)
			video_idx +=1
		elif comp_type == 'problem':
			search_problem_in_course(problem_excel2list(problem_idx,sheetproblem),Course_extraction(),course_path)
			problem_idx +=1
		elif comp_type == 'text':
			search_html_in_course(html_excel2list(text_idx,sheethtml),Course_extraction(),course_path,sheethtml)
			text_idx +=1
			
	



def make_tarfile():
	
	# compress course content in a targz file and ready to import.
	print("file is being compressed as tar.gz ")
	with tarfile.open(course_path + '/' + course_path + '.tar.gz', 'w:gz') as tar:
		for f in os.listdir(course_path):
			tar.add(course_path + "/" + f, arcname=os.path.basename(f))
		tar.close()
	print("uploadable file is created at " + course_path + '/' + course_path + '.tar.gz')


	"""
	addpath = 'set PATH=%PATH%;C:\Program Files\7-Zip\ ' 
	compress_tar = '7z a course.tar course\ '
	compress_targz = '7z a course.tar.gz course.tar'
	os.system(addpath)
	os.system(compress_tar)
	os.system(compress_targz)
	os.remove('course.tar')
	"""


def main():

	flag = 0
	global file
	
	'''
	select options
	1) Create course outline --> create section,subsection,unit
	2) Add course content --> add components (text,quiz,video)
	3) upload video to youtube --> upload video file to youtube account

	'''

	while(flag==0):
		command = input("enter [1-3]\n1.Create course outline \n2.Add course contents\n3.Upload video to Youtube\n")
		if command == '1':
			print ('Create course outline is chosen')
			create_course()
			create_section()
			create_subsection()
			create_unit()
			flag = 1
		elif command == '2':
			print ('Add course contents is chosen')
			add_component()
			make_tarfile()
			flag = 1
		elif command == '3':
			os.system('python video2youtube.py')
			flag = 1
		else:
			print ('wrong command, try again!!!!')

	
	

	






if __name__ == '__main__':
	try:
		main()
	except KeyboardInterrupt:
		logging.warn("\n\nCTRL-C detected, shutting down....")
		sys.exit(ExitCode.OK)