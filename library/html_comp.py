# coding=utf-8


import os, tarfile, shutil, xlrd,xlwt, datetime,sys
import json,urllib


from bs4 import BeautifulSoup
from lxml import etree
from six.moves import html_parser

"""
	sheet-> video component
"""

HTMLSHEET = "text"
HTMLINDEX = 0
HTMLSECTION = 1
HTMLSUBSECTION = 2
HTMLUNIT = 3
HTMLCOMPNAME = 4
HTMLLOC = 5
HTMLFILE = 6






def html_excel2list(row,sheethtml):
	
	html_info = []
	#for row in range(1, sheetvideo.nrows):

	html_idx = sheethtml.cell_value(row,HTMLINDEX)
	html_section = sheethtml.cell_value(row,HTMLSECTION)
	html_subsection = sheethtml.cell_value(row,HTMLSUBSECTION)
	html_unit = sheethtml.cell_value(row,HTMLUNIT)
	html_displayname = sheethtml.cell_value(row,HTMLCOMPNAME)
	html_dir = sheethtml.cell_value(row,HTMLLOC)
	html_file = sheethtml.cell_value(row,HTMLFILE)


	html_info.append({'idx':html_idx,
		'file_row':row,
		'section':html_section,
		'subsection':html_subsection,
		'unit':html_unit,
		'html_display':html_displayname,
		'html_loc':html_dir,
		'html_file':html_file}) 
	#print(html_info)
	return(html_info[0])






def find_section_name(row_section,course_section):
	
	for course_sec_row in course_section:
		course_sec_row['section_name'] = course_sec_row['section_name'].rstrip()
		row_section['section'] = row_section['section'].rstrip()
		if course_sec_row['section_name']== row_section['section']:
			print ('found section: ' + (row_section['section'])+ ' in the exported course')
			selected_section = course_sec_row
			return selected_section

	print ('no section: ' + (row_section['section']) + ' in the exported course')
	exit()

	return selected_section

def find_subsection_name(row_subsection,course_subsection,selected_section):
	
	for course_subsec_row in course_subsection:
		course_subsec_row['subsection_name'] = course_subsec_row['subsection_name'].rstrip()
		row_subsection['subsection'] = row_subsection['subsection'].rstrip()
		if course_subsec_row['subsection_name']== row_subsection['subsection']:
			if course_subsec_row['subsection_link'] in selected_section['assoc_subsection_url']:
				print ('found subsection: ' + (row_subsection['subsection'])+ ' in the exported course')
				selected_subsection = course_subsec_row
				return selected_subsection

	print ('no subsection: ' + (row_subsection['subsection']) + ' in the exported course')
	exit()


def find_unit_name(row_unit,course_unit,selected_subsection,course_path):
	
	for course_unit_row in course_unit:
		course_unit_row['unit_name'] = course_unit_row['unit_name'].rstrip()
		row_unit['unit'] = row_unit['unit'].rstrip()
		if course_unit_row['unit_name'] == row_unit['unit']:
			if course_unit_row['unit_link'] in selected_subsection['assoc_unit_url']:
				print('found unit: ' + (row_unit['unit'])+ ' in the exported course')
				tree = etree.parse(os.path.join(course_path,'vertical',course_unit_row['unit_link']+'.xml'))
				root = tree.getroot()
				new_html_link = "text_content" +  "{0:0=2d}".format(int(row_unit['idx']))
				etree.SubElement(root, 'html',url_name=new_html_link)
				doc = etree.ElementTree(root)
				doc.write(os.path.join(course_path,'vertical',course_unit_row['unit_link']+'.xml'), pretty_print=True, xml_declaration=False, encoding='utf-8')
				selected_unit = {'unit_link':course_unit_row['unit_link'],'unit_name':course_unit_row['unit_name'],'assoc_html_url':new_html_link}
				print('      added text link: '+new_html_link)
				return selected_unit

	print ('no unit: ' + (row_unit['unit']) + ' in the exported course')
	exit()








############################for editing video text component 	###############################################

################################################################################################################



def add_html(row_html,selected_unit,course_path,sheethtml):
	html_file = selected_unit['assoc_html_url']+ '.xml'
	html_path =  os.path.join(course_path, 'html', html_file)
	page = etree.Element('html', display_name=row_html['html_display'], filename = selected_unit['assoc_html_url'])
	doc = etree.ElementTree(page)
	doc.write(html_path, pretty_print=True, xml_declaration=False, encoding='utf-8')
	print ('crate a new html component: ' + (row_html['html_display']) + ' in the exported course')

	source_loc = (sheethtml.cell_value(row_html['file_row'],HTMLLOC))
	source = (sheethtml.cell_value(row_html['file_row'],HTMLFILE))
	source_file = os.path.join(source_loc, source)


	#html_path = path + "/html"
	des_file = os.path.join(course_path,'html',selected_unit['assoc_html_url']+ '.html')
	shutil.copyfile(source_file, des_file)
	print ('copy text content: ' + (row_html['html_display']) + ' to the exported course')
	modify_figure_src(des_file,selected_unit['assoc_html_url'],course_path,row_html['html_loc'])


	print ("------------------------------------------------------------\n\n")





def modify_figure_src(_des_path,_filename,_course_path,html_loc):



	htmlfile_read = open(_des_path, 'r') 
	text_dump = htmlfile_read.read() 
	tag_html = BeautifulSoup(text_dump, 'html.parser')
	img_tag = tag_html.find_all('img')
	htmlfile_read.close()
	if len(img_tag) != 0:


		for i in range(len(img_tag)):

			figure_name, file_extension = os.path.splitext(img_tag[i].attrs['src'])
			new_figure_name = _filename + '_fig_' + str(i) + file_extension
			new_figure_path = '/static/'+  new_figure_name 
			img_tag[i].attrs['src'] = new_figure_path

			mod_text = tag_html.prettify()
			htmlfile_write = open(_des_path, 'w',encoding="utf-8") 
			htmlfile_write.write(mod_text) 
			htmlfile_write.close()
			
			path_fig = urllib.parse.unquote(os.path.join(html_loc,figure_name+ file_extension))
			shutil.copyfile( path_fig, os.path.join(_course_path,'static',new_figure_name))


		print('figure sources are all modified')



	else:
		print("No figure in this text component")


	




def search_html_in_course(row_from_excel,course_structure,course_path,sheethtml):
	
	selected_section = find_section_name(row_from_excel,course_structure.sections())
	selected_subsection = find_subsection_name(row_from_excel,course_structure.subsections(),selected_section)
	selected_unit=find_unit_name(row_from_excel,course_structure.units(),selected_subsection,course_path)
	add_html(row_from_excel,selected_unit,course_path,sheethtml)