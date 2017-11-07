# coding=utf-8


import os, tarfile, shutil, xlrd,xlwt, datetime,sys
import json,pysrt
from lxml import etree
from six.moves import html_parser


"""
	sheet-> video component
"""
VIDEOSHEET = "video"
VIDEOINDEX = 0
VIDEOSECTION = 1
VIDEOSUBSECTION = 2
VIDEOUNIT = 3
VIDEOURL = 4
VIDEONAME = 5
TRANSCRIPTDIR = 6
ENTRANSCRIPTFILE = 7
JPTRANSCRIPTFILE = 8





def video_excel2list(row,sheetvideo):
	
	video_info = []
	#for row in range(1, sheetvideo.nrows):

	video_idx = sheetvideo.cell_value(row,VIDEOINDEX)
	video_section = sheetvideo.cell_value(row,VIDEOSECTION)
	video_subsection = sheetvideo.cell_value(row,VIDEOSUBSECTION)
	video_unit = sheetvideo.cell_value(row,VIDEOUNIT)
	video_url = sheetvideo.cell_value(row,VIDEOURL)
	video_url_id = video_url.rsplit('https://youtu.be/', 1)[1]
	video_name = sheetvideo.cell_value(row,VIDEONAME)
	transcript_dir = sheetvideo.cell_value(row,TRANSCRIPTDIR)
	en_transcript_file = sheetvideo.cell_value(row,ENTRANSCRIPTFILE)
	jp_transcript_file = sheetvideo.cell_value(row,JPTRANSCRIPTFILE)


	video_info.append({'idx':video_idx,
		'section':video_section,
		'subsection':video_subsection,
		'unit':video_unit,
		'video_link':video_url,
		'video_id':video_url_id,
		'video_name':video_name,
		'transcript_dir':transcript_dir,
		'en_transcript_file':en_transcript_file,
		'jp_transcript_file':jp_transcript_file}) 
	return(video_info[0])






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
				new_video_link = "video" +  "{0:0=2d}".format(int(row_unit['idx']))
				etree.SubElement(root, 'video',url_name=new_video_link)
				doc = etree.ElementTree(root)
				doc.write(os.path.join(course_path,'vertical',course_unit_row['unit_link']+'.xml'), pretty_print=True, xml_declaration=False, encoding='utf-8')
				selected_unit = {'unit_link':course_unit_row['unit_link'],'unit_name':course_unit_row['unit_name'],'assoc_video_url':new_video_link}
				print('      added video link: '+new_video_link)
				return selected_unit

	print ('no unit: ' + (row_unit['unit']) + ' in the exported course')
	exit()








############################for editing video video_component 	###############################################

def modify_video(row_video,selected_unit,course_path):
	


	

	video_file = selected_unit['assoc_video_url']+ '.xml'
	video_path =  os.path.join(course_path, 'video', video_file)
	youtube = '1.00:'+ row_video['video_id'].rstrip()
	youtube_id_1_0 = row_video['video_id'].rstrip()
	display_name = row_video['video_name']
	urlname = selected_unit['assoc_video_url']
	#print(type(selected_unit['assoc_video_url']))
	download_TF = "false"
	edx_video_id = ""
	url_source = "[]"
	link_sub = row_video['video_id']
	transcripts = transcript2static(row_video,course_path)
	#transcripts_name, file_extension = os.path.splitext(transcripts['en']) 

	if transcripts != []:
		page = etree.Element('video', youtube=youtube, url_name = urlname, display_name = display_name, download_video=download_TF,download_track="true",edx_video_id =edx_video_id, html5_sources=url_source, youtube_id_1_0=youtube_id_1_0,transcripts=str(json.dumps(transcripts)), sub=youtube_id_1_0) 
		for key, value in transcripts.items():
			etree.SubElement(page, 'transcript',language=key,src=value)
	else: 
		page = etree.Element('video', youtube=youtube, url_name = urlname, display_name = display_name, download_video=download_TF,edx_video_id =edx_video_id, html5_sources=url_source, youtube_id_1_0=youtube_id_1_0,transcripts="") 
	doc = etree.ElementTree(page)
	doc.write(video_path, pretty_print=True, xml_declaration=False, encoding='utf-8')
	print ('crate a new video: ' + (row_video['video_name']) + ' in the exported course')
	print ("------------------------------------------------------------\n\n")



def transcript2static(video_info,course_path):
	static_path = os.path.join(course_path,'static')
	
	transcripts = dict()
	if video_info['en_transcript_file'] != '':
		transcript_path = os.path.join(video_info['transcript_dir'],video_info['en_transcript_file'])
		shutil.copy(transcript_path, static_path)
		json_filename = 'subs_' + video_info['video_id'].rstrip() + '.srt.sjson'
		convert_srt2json(transcript_path,json_filename,static_path)
		transcripts['en'] = str(video_info['en_transcript_file'])



	if video_info['jp_transcript_file'] != '':
		transcript_path = os.path.join(video_info['transcript_dir'],video_info['jp_transcript_file'])
		shutil.copy(transcript_path, static_path)
		json_filename = 'ja_subs_' + video_info['video_id'].rstrip() + '.srt.sjson'
		convert_srt2json(transcript_path,json_filename,static_path)
		transcripts['ja'] = str(video_info['jp_transcript_file'])

	return transcripts


def convert_srt2json(srt_file,json_filename,course_static_path):
	subs = pysrt.open(srt_file)
	t_start_milli = []
	t_end_milli = []
	text = []
	for line_sub in subs:
	    h2milli = [line_sub.start.hours*3600*1000 , line_sub.end.hours*3600*1000 ]
	    m2milli = [line_sub.start.minutes*60*1000 , line_sub.end.minutes*60*1000 ]
	    s2milli = [line_sub.start.seconds*1000 , line_sub.end.seconds*1000 ]
	    milli = [line_sub.start.milliseconds , line_sub.end.milliseconds*1000 ]
	    
	    t_start_milli.append(h2milli[0] + m2milli[0] + s2milli[0] + milli[0])
	    t_end_milli.append(h2milli[1] + m2milli[1] + s2milli[1] + milli[1])
	    text.append(line_sub.text)
	    
	json_str = json.dumps({"start":t_start_milli,"end":t_end_milli,"text":text}, sort_keys=False,indent=4, separators=(',', ': '))
	dest_file = os.path.join(course_static_path,json_filename)
	json_file = open(dest_file, 'w',encoding="utf-8") 
	json_file.write(json_str) 
	json_file.close()


################################################################################################################





def search_video_in_course(row_from_excel,course_structure,course_path):
	
	selected_section = find_section_name(row_from_excel,course_structure.sections())
	selected_subsection = find_subsection_name(row_from_excel,course_structure.subsections(),selected_section)
	selected_unit=find_unit_name(row_from_excel,course_structure.units(),selected_subsection,course_path)
	modify_video(row_from_excel,selected_unit,course_path)