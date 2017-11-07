#!/usr/bin/env python
# -*- coding: utf-8 -*-


import os, tarfile, shutil, xlrd,xlwt, datetime,sys
import json
import string
from lxml import etree
from six.moves import html_parser

"""
	sheet->Probem
"""
PROBLEMSHEET = "problem"
PROBLEMINDEX = 0
PROBLEMSECTION = 1
PROBLEMSUBSECTION = 2
PROBLEMUNIT = 3
PROBLEMDIR = 4
PROBLEMNAME = 5
PROBLEMSHEETNAME = 6
PROBLEMDISPLAYNAME = 7
PROBLEMTYPE = 8



################################# class of problem type ##################################





class Problem_droplist:

	def __init__(self,info):
		prob_detail_path = os.path.join(info['dir'],info['filename'])
		#print(prob_detail_path)
		wb_prob = xlrd.open_workbook(prob_detail_path)
		self.sheetstruc = wb_prob.sheet_by_name( info['sheet'])
		self.n_droplists= (self.sheetstruc.ncols-4)//3
		self.prob_disp = 0
		self.prob_weight = 1
		self.prob_attempt = 2
		self.prob_hint = 3
		self.question_col = 4
		self.choice_col = 5
		self.ans_col = 6
		print('number of question is '+str(self.n_droplists))
		if self.n_droplists is float:
			print('number of column does not match with number of droplist')
			exit()

		
	def display_name(self):
		return(self.sheetstruc.cell_value(1,self.prob_disp))

	def hint(self):
		return(self.sheetstruc.cell_value(1,self.prob_hint))


	def weight(self):
		weigth_per_question = self.sheetstruc.cell_value(1,self.prob_weight)
		if weigth_per_question == '':
			return('')
		else:
			total_weight=float(float(weigth_per_question)*self.n_droplists)
			return(str(total_weight))

	def attempt(self):
		if self.sheetstruc.cell_value(1,self.prob_attempt) == '':
			return('')
		else:
			attempt = int(self.sheetstruc.cell_value(1,self.prob_attempt))
			return(str(attempt))

	
	def droplists(self,element_obj):

		for droplist_idx in range(0,self.n_droplists):
			print(droplist_idx)
			q_text = ''
			a_text = '<optionresponse><optioninput>'

			for row_ in range(1,self.sheetstruc.nrows):
				tmp = self.sheetstruc.cell_value(row_,self.question_col)
				if tmp != '':

					question_page=etree.SubElement(element_obj,'p')
					question_page.text = tmp
					

			opt_response_page = etree.SubElement(element_obj,'optionresponse')
			sub_opt_page= etree.SubElement(opt_response_page,'optioninput')
			for row_ in range(1,self.sheetstruc.nrows):
				answer_text = self.sheetstruc.cell_value(row_,self.ans_col)
				choice_text = self.sheetstruc.cell_value(row_,self.choice_col)

				if answer_text == '':
					continue

				if answer_text.lower() == 't'.lower():
					choice = etree.SubElement(sub_opt_page,'option',correct='True')
					choice.text = str(choice_text)
				else:
					choice = etree.SubElement(sub_opt_page,'option',correct='False')
					choice.text = str(choice_text)

			self.question_col = self.question_col + 3
			self.choice_col = self.choice_col+3
			self.ans_col = self.ans_col+3

		if self.hint() != '':
			demand_hint = etree.SubElement(element_obj,'demandhint') 
			hint = etree.SubElement(demand_hint,'hint')
			hint.text = self.hint() 
			
		print(etree.tostring(element_obj))
		return(element_obj)


	def create_file(self,filename,course_path):
		new_problem_file = os.path.join(course_path,'problem',filename)
		
		page = etree.Element('problem', display_name=self.display_name()) 
		if self.weight() != '':
			page.set('weight',self.weight())
		if self.attempt() != '':
			page.set('max_attempts',self.attempt())

		full_xml_obj = self.droplists(page)
		doc = etree.ElementTree(page)
		doc.write(new_problem_file, pretty_print=True, xml_declaration=False, encoding='utf-8')
	


class Problem_multichoice:

	def __init__(self,info):
		prob_detail_path = os.path.join(info['dir'],info['filename'])
		wb_prob = xlrd.open_workbook(prob_detail_path)
		self.sheetstruc = wb_prob.sheet_by_name( info['sheet'])
		self.n_multichoice= (self.sheetstruc.ncols-4)//3
		self.prob_disp = 0
		self.prob_weight = 1
		self.prob_attempt = 2
		self.prob_hint = 3
		self.question_col = 4
		self.multichoice_col = 5
		self.ans_col = 6
		print('number of question is '+str(self.n_multichoice))
		if self.n_multichoice is float:
			print('number of column does not match with number of droplist')
			exit()

		
	def display_name(self):
		return(self.sheetstruc.cell_value(1,self.prob_disp))

	def hint(self):
		return(self.sheetstruc.cell_value(1,self.prob_hint))


	def weight(self):
		weigth_per_question = self.sheetstruc.cell_value(1,self.prob_weight)
		if weigth_per_question == '':
			return('')
		else:
			total_weight=float(float(weigth_per_question)*self.n_multichoice)
			return(str(total_weight))

	def attempt(self):
		if self.sheetstruc.cell_value(1,self.prob_attempt) == '':
			return('')
		else:
			attempt = int(self.sheetstruc.cell_value(1,self.prob_attempt))
			return(str(attempt))

	
	def multichoice(self,element_obj):



		for multichoice_idx in range(0,self.n_multichoice):
			print(multichoice_idx)
			
			for row_ in range(1,self.sheetstruc.nrows):
				tmp = self.sheetstruc.cell_value(row_,self.question_col)
				if tmp != '':

					question_page=etree.SubElement(element_obj,'p')
					question_page.text = tmp
					

			multi_response_page = etree.SubElement(element_obj,'multiplechoiceresponse')
			choice_group_page= etree.SubElement(multi_response_page,'choicegroup',type='MultipleChoice')
			for row_ in range(1,self.sheetstruc.nrows):
				answer_text = self.sheetstruc.cell_value(row_,self.ans_col)
				choice_text = self.sheetstruc.cell_value(row_,self.multichoice_col)

				if answer_text == '':
					continue

				if answer_text.lower() == 't'.lower():
					choice = etree.SubElement(choice_group_page,'choice',correct='True')
					choice.text = str(choice_text)
				else:
					choice = etree.SubElement(choice_group_page,'choice',correct='False')
					choice.text = str(choice_text)

			self.question_col = self.question_col + 3
			self.multichoice_col = self.multichoice_col+3
			self.ans_col = self.ans_col+3

		if self.hint() != '':
			demand_hint = etree.SubElement(element_obj,'demandhint') 
			hint = etree.SubElement(demand_hint,'hint')
			hint.text = self.hint() 
			
		print(etree.tostring(element_obj))
		return(element_obj)


	def create_file(self,filename,course_path):
		new_problem_file = os.path.join(course_path,'problem',filename)
		page = etree.Element('problem', display_name=self.display_name()) 
		if self.weight() != '':
			page.set('weight',self.weight())
		if self.attempt() != '':
			page.set('max_attempts',self.attempt())

		full_xml_obj = self.multichoice(page)
		doc = etree.ElementTree(page)
		doc.write(new_problem_file, pretty_print=True, xml_declaration=False, encoding='utf-8')




class Problem_checkbox:

	def __init__(self,info):
		prob_detail_path = os.path.join(info['dir'],info['filename'])
		wb_prob = xlrd.open_workbook(prob_detail_path)
		self.sheetstruc = wb_prob.sheet_by_name( info['sheet'])
		self.n_checkbox= (self.sheetstruc.ncols-4)//3
		self.prob_disp = 0
		self.prob_weight = 1
		self.prob_attempt = 2
		self.prob_hint = 3
		self.question_col = 4
		self.checkbox_col = 5
		self.ans_col = 6
		print('number of question is '+str(self.n_checkbox))
		if self.n_checkbox is float:
			print('number of column does not match with number of droplist')
			exit()

		
	def display_name(self):
		return(self.sheetstruc.cell_value(1,self.prob_disp))

	def hint(self):
		return(self.sheetstruc.cell_value(1,self.prob_hint))


	def weight(self):
		weigth_per_question = self.sheetstruc.cell_value(1,self.prob_weight)
		if weigth_per_question == '':
			return('')
		else:
			total_weight=float(float(weigth_per_question)*self.n_checkbox)
			return(str(total_weight))

	def attempt(self):
		if self.sheetstruc.cell_value(1,self.prob_attempt) == '':
			return('')
		else:
			attempt = int(self.sheetstruc.cell_value(1,self.prob_attempt))
			return(str(attempt))

	
	def checkbox(self,element_obj):



		for checkboxs_idx in range(0,self.n_checkbox):
			print(checkboxs_idx)
			
			for row_ in range(1,self.sheetstruc.nrows):
				tmp = self.sheetstruc.cell_value(row_,self.question_col)
				if tmp != '':

					question_page=etree.SubElement(element_obj,'p')
					question_page.text = tmp
					

			choice_response_page = etree.SubElement(element_obj,'choiceresponse')
			checkbox_group_page= etree.SubElement(choice_response_page,'checkboxgroup')
			for row_ in range(1,self.sheetstruc.nrows):
				answer_text = self.sheetstruc.cell_value(row_,self.ans_col)
				checkbox_text = self.sheetstruc.cell_value(row_,self.checkbox_col)

				if answer_text == '':
					continue

				if answer_text.lower() == 't'.lower():
					checkbox_obj = etree.SubElement(checkbox_group_page,'choice',correct='True')
					checkbox_obj.text = checkbox_text
				else:
					checkbox_obj = etree.SubElement(checkbox_group_page,'choice',correct='False')
					checkbox_obj.text = checkbox_text

			self.question_col = self.question_col + 3
			self.checkbox_col = self.checkbox_col+3
			self.ans_col = self.ans_col+3

		if self.hint() != '':
			demand_hint = etree.SubElement(element_obj,'demandhint') 
			hint = etree.SubElement(demand_hint,'hint')
			hint.text = self.hint() 
			
		print(etree.tostring(element_obj))
		return(element_obj)


	def create_file(self,filename,course_path):
		new_problem_file = os.path.join(course_path,'problem',filename)
		
		page = etree.Element('problem', display_name=self.display_name()) 
		if self.weight() != '':
			page.set('weight',self.weight())
		if self.attempt() != '':
			page.set('max_attempts',self.attempt())

		full_xml_obj = self.checkbox(page)
		doc = etree.ElementTree(page)
		doc.write(new_problem_file, pretty_print=True, xml_declaration=False, encoding='utf-8')




class Problem_fillblank:

	def __init__(self,info):
		prob_detail_path = os.path.join(info['dir'],info['filename'])
		wb_prob = xlrd.open_workbook(prob_detail_path)
		self.sheetstruc = wb_prob.sheet_by_name( info['sheet'])
		self.n_fillblank= (self.sheetstruc.ncols-4)//2
		self.prob_disp = 0
		self.prob_weight = 1
		self.prob_attempt = 2
		self.prob_hint = 3
		self.question_col = 4
		self.ans_col = 5

		print('number of question is '+str(self.n_fillblank))
		if self.n_fillblank is float:
			print('number of column does not match with number of droplist')
			exit()

		
	def display_name(self):
		return(self.sheetstruc.cell_value(1,self.prob_disp))

	def hint(self):
		return(self.sheetstruc.cell_value(1,self.prob_hint))


	def weight(self):
		weigth_per_question = self.sheetstruc.cell_value(1,self.prob_weight)
		if weigth_per_question == '':
			return('')
		else:
			total_weight=float(float(weigth_per_question)*self.n_multiquestion)
			return(str(total_weight))

	def attempt(self):
		if self.sheetstruc.cell_value(1,self.prob_attempt) == '':
			return('')
		else:
			attempt = int(self.sheetstruc.cell_value(1,self.prob_attempt))
			return(str(attempt))

	
	def fillblank(self,element_obj):



		for fillblank_idx in range(0,self.n_fillblank):
			print(fillblank_idx)
			
			for row_ in range(1,self.sheetstruc.nrows):
				tmp = self.sheetstruc.cell_value(row_,self.question_col)
				if tmp != '':

					question_page=etree.SubElement(element_obj,'p')
					question_page.text = tmp
					
			answer_text = self.sheetstruc.cell_value(1,self.ans_col)
			string_response_page = etree.SubElement(element_obj,'stringresponse',answer=str(answer_text),type='ci')
			etree.SubElement(string_response_page,'textline',size='20')
	

			self.question_col = self.question_col + 2
			self.ans_col = self.ans_col+2

		if self.hint() != '':
			demand_hint = etree.SubElement(element_obj,'demandhint') 
			hint = etree.SubElement(demand_hint,'hint')
			hint.text = self.hint() 
			
		print(etree.tostring(element_obj))
		return(element_obj)


	def create_file(self,filename,course_path):
		new_problem_file = os.path.join(course_path,'problem',filename)
		page = etree.Element('problem', display_name=self.display_name()) 
		if self.weight() != '':
			page.set('weight',self.weight())
		if self.attempt() != '':
			page.set('max_attempts',self.attempt())

		full_xml_obj = self.fillblank(page)
		doc = etree.ElementTree(page)
		doc.write(new_problem_file, pretty_print=True, xml_declaration=False, encoding='utf-8')



		########################################################################################






def problem_excel2list(row,sheetproblem):

	problem_info = []

	problem_idx = sheetproblem.cell_value(row,PROBLEMINDEX)
	problem_section = sheetproblem.cell_value(row,PROBLEMSECTION)
	problem_subsection = sheetproblem.cell_value(row,PROBLEMSUBSECTION)
	problem_unit = sheetproblem.cell_value(row,PROBLEMUNIT)
	problem_dir = sheetproblem.cell_value(row,PROBLEMDIR)
	problem_name = sheetproblem.cell_value(row,PROBLEMNAME)
	problem_sheet = sheetproblem.cell_value(row,PROBLEMSHEETNAME)
	problem_type = sheetproblem.cell_value(row,PROBLEMTYPE)

	problem_info.append({'idx':int(problem_idx),
		'section':problem_section,
		'subsection':problem_subsection,
		'unit':problem_unit,
		'dir':problem_dir,
		'filename':problem_name,
		'sheet':problem_sheet,
		'type':problem_type}) 
	return(problem_info[0])





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
				print ('found unit: ' + row_unit['unit']+ ' in the exported course')
				tree = etree.parse(os.path.join(course_path,'vertical',course_unit_row['unit_link']+'.xml'))
				root = tree.getroot()
				new_problem_link = "problem" +  "{0:0=2d}".format(int(row_unit['idx']))
				etree.SubElement(root, 'problem',url_name=new_problem_link)
				doc = etree.ElementTree(root)
				doc.write(os.path.join(course_path,'vertical',course_unit_row['unit_link']+'.xml'), pretty_print=True, xml_declaration=False, encoding='utf-8')
				selected_unit = {'unit_link':course_unit_row['unit_link'],'unit_name':course_unit_row['unit_name'],'assoc_problem_url':new_problem_link}
				print('      added problem link: '+new_problem_link)

				return selected_unit

	print ('no unit: ' + (row_unit['unit']) + 'in the exported course')
	exit()

	return selected_unit





def add_problem(problem_source_info,selected_unit,course_path):
	new_problem_link = "problem" +  "{0:0=2d}".format(int(problem_source_info['idx']))

	#link2unit(selected_unit_path,new_problem_link)
	print(problem_source_info['type'])
	if problem_source_info['type'] == 'droplist':
		problem_instance = Problem_droplist(problem_source_info)

	elif problem_source_info['type'] == 'multiple_choice':
		problem_instance = Problem_multichoice(problem_source_info)

	elif problem_source_info['type'] == 'checkbox':
		problem_instance = Problem_checkbox(problem_source_info)

	elif problem_source_info['type'] == 'fill_blank':
		problem_instance = Problem_fillblank(problem_source_info)
	else:
		print('no problem type available')
		exit()
	
	problem_instance.create_file(new_problem_link+'.xml',course_path)
		



def search_problem_in_course(row_from_excel,course_structure,course_path):
	
	selected_section = find_section_name(row_from_excel,course_structure.sections())
	selected_subsection = find_subsection_name(row_from_excel,course_structure.subsections(),selected_section)
	selected_unit=find_unit_name(row_from_excel,course_structure.units(),selected_subsection,course_path)
	add_problem(row_from_excel,selected_unit,course_path)