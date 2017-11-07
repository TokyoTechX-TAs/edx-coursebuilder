# edx-course-builder

This tool is developed to provide an alternative approach for designing course outlines and adding course materials offline. The output (modified course) can be easily used to upload full-content course instance back to edx.studio with import function.    

# This tool consists of 4 parts

1) a Python script which was developed based on [EdX Open Learning XML(OLX) format](http://edx.readthedocs.io/projects/edx-open-learning-xml/en/latest/). It consists of 3 functions.
  * create course outline. 
  * add course contents. 
  * upload video to YouTube. Detail avaiable [here](https://github.com/KeNopphon/edx-YouTube-video-uploading-tool)
2) an macro-excel file where information about course outline (Section, Subsection, and Unit), and course contents (text, quiz, and video) are contained. User requires to fill these information as a mean to design a course before running the code 
3) source files of course contents (text file, image file, video files, subtitle file, excel file) 
4) exmpty edX course instance exported from edx-studio 

In a nutshell, this tool creates the course offline according to course information input into an Excel sheet which can be uploaded to the Open edX Studio using the export and then the import function. For example there are four sheets in an Excel file: course outline, text content blocks, assessment items, and videos. The names of course sections are typed in the order they appear in the course outline, HTML version of the text blocks and Excel cells containing the text and assessments respectively are input into the separate sheets, and a list of video files are input into a sheet consisting of YouTube video links and a list of closed captions .srt files that will be imported into the Open edX Studio course instance. 

Please see the [pptx](https://github.com/KeNopphon/edx-course-builder/blob/master/course_builder.pptx) or [pdf](https://github.com/KeNopphon/edx-course-builder/blob/master/course_builder.pdf) file for further instruction.


# dependencies
- xlrd, xlwt, pysrt


# test environment
Python 3 run on Windows 10 




