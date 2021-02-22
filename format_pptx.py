#!python
# read_pptx.py - read content from powerpoints

import os
from math import inf
from pptx import Presentation
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.shared import OxmlElement
import xml.etree.ElementTree as gfg
from xml.dom import minidom


class Course:
    
    def __init__(self, pptx_file, file_id):
        self.file_id = file_id
        self.pptx_file = pptx_file
    
    @property
    def pptx_file(self):
        return self._pptx_file

    @property
    def pres(self):
        # return presentation
        return Presentation(self.pptx_file)

    @property
    def course_title(self):
        # title slide - first slide of powerpoint
        title_slide = self.pres.slides[0]

        for shape in title_slide.shapes:
            if not shape.has_text_frame:
                continue
            return shape.text_frame.text

    @property
    def course_id(self):
        # title slide - first slide of powerpoint
        title_slide = self.pres.slides[0]

        # skip course title (first item)
        itershapes = iter(title_slide.shapes)
        next(itershapes)
        
        for shape in itershapes:
            if not shape.has_text_frame:
                continue
            return shape.text_frame.text

    @pptx_file.setter
    def pptx_file(self, f):
        # check if file exists
        if not os.path.isfile(f): raise FileExistsError

        # check if file is pptx file
        if os.path.splitext(f)[-1].lower() != '.pptx':
            raise Exception('must provide .pptx file!')
        self._pptx_file = f

    def get_pptx_dict(self):
        # return dictionary of pptx
        pptx = []
        for n, _ in enumerate(self.pres):
            pptx.append({
                'slide_number': n + 1,
                'slide_text': self.get_slide_text(n+1),
                'slide_notes': self.get_slide_notes(n+1)
            })
        return pptx

    def get_notes(self, *args):
        # return list of tuples: [(page_num, notes), (...)]
        narration = []
        
        # iterate through pptx slides
        for n, _ in enumerate(self.pres.slides):
            slide_text = self.get_slide_text(n+1)
            slide_notes = self.get_slide_notes(n+1)

            # apply filters
            for arg in args:
                slide_notes = arg(slide_notes, slide_text, n+1)

            # add slide notes to narration list
            if slide_notes:
                narration.append((n+1, slide_notes.replace('\n', '')))

        return narration
    
    def get_slide_text(self, slide_num=1):
        # return text in slide
        try:
            slide = self.pres.slides[slide_num-1]
        except IndexError:
            raise Exception('Slide number must be less than slide length')
        text_runs = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text_runs.append(run.text.replace('\n', ''))

        return '\n'.join(text_runs)

    def get_slide_notes(self, slide_num=1):
        # return slide notes section
        try:
            slide = self.pres.slides[slide_num-1]
        except IndexError:
            raise Exception('Slide number must be less than slide length')
        if slide.has_notes_slide:
            slide_notes = slide.notes_slide.notes_text_frame.text
            return slide_notes.replace('\n', '')

        return ''

    def mk_narration_docx(self, *args):

        def prevent_doc_breakup(doc):
            # prevents document tables from spanning 2 pages
            tags = doc.element.xpath('//w:tr')
            rows = len(tags)
            for row in range(0,rows):
                tag = tags[row]                     # Specify which <w:r> tag you want
                child = OxmlElement('w:cantSplit')  # Create arbitrary tag
                tag.append(child)                   # Append in the new tag
            
        doc = Document()
        style = doc.styles['Normal']
        style.font.size, style.font.name = Pt(11), 'Calibri'
        doc.add_paragraph(f'{self.course_id} - Narration Script')
        doc.add_paragraph(self.course_title)
        table = doc.add_table(rows=0, cols=2)
        table.style = 'Table Grid'
        for note in self.get_notes(*args):
            row_cells = table.add_row().cells
            row_cells[0].text = f'{self.file_id}_{note[0]}'
            row_cells[1].text = note[1]
        for cell in table.columns[0].cells:
            cell.width = Inches(0.5)
        for cell in table.columns[1].cells:
            cell.width = Inches(7)
        for section in doc.sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        prevent_doc_breakup(doc)
        doc_file = f'{self.course_title} narration script_01.docx'
        try:
            doc.save(doc_file)
            print(f'{doc_file} written!')
        except PermissionError:
            print(f'{doc_file} exists! or invalid permissions!')
        return doc_file

    def mk_narration_xml(self, *args, section_headers=None, course_menu=False):
        
        def setup_course_xml(course_menu):
            # create course element
            the_course = gfg.Element('theCourse')
            
            # write comment
            comment = gfg.Comment('Generated with py_course')
            the_course.append(comment)

            # write course title
            course_title = gfg.SubElement(the_course, 'myCourseTitle')
            course_title.text = self.course_title

            # write study guide file
            study_guide = gfg.SubElement(the_course, 'studyGuidePDF')
            study_guide.text = f'{self.course_id}_508.pdf'

            # write study guide print
            study_guide_print = gfg.SubElement(the_course, 'studyGuidePrint')
            study_guide_print.text = f'{self.course_id}_StudyGuide.pdf'

            # write courseMenu option
            if course_menu:
                course_menu = gfg.SubElement(the_course, 'courseMenu')
                course_menu.text = 'YES'

            return the_course
        
        # create xml file with notes
        if not section_headers:
            section_headers = [('Main', 1)]

        # course xml
        the_course = setup_course_xml(course_menu)

        # narration text
        notes = self.get_notes(*args)

        current_header_num = 0              # section_header iterator
        file_num = 1                        # section file iterator
        next_header = inf                   # next section_header slide start number
        if len(section_headers) > 1:
            next_header = section_headers[current_header_num + 1][1]
        
        the_sections = gfg.SubElement(the_course, 'theSections', {'title': section_headers[current_header_num][0]})
        for note in notes:
            # note passes next_header
            if note[0] >= next_header:
                # move current_header_num
                current_header_num += 1

                # create new theSections elem
                the_sections = gfg.SubElement(the_course, 'theSections', {'title': section_headers[current_header_num][0]})

                # move next_header
                if len(section_headers) > current_header_num + 1:
                    next_header = section_headers[current_header_num + 1][1]
                else:
                    next_header = inf

                # set file_num to 0
                file_num = 1
            
            # add sectionNumber elem within theSections elem
            section_num = gfg.SubElement(the_sections, 'sectionNumber')

            # add theFIleToLoad elem within sectionNumber elem
            the_file_to_load = gfg.SubElement(section_num, 'theFileToLoad')
            the_file_to_load.text = f'{self.file_id}_s{current_header_num}_{file_num}.html'
            file_num += 1

            # add closedCaption elem within sectionNumber elem
            closed_caption = gfg.SubElement(section_num, 'closedCaptionText')
            closed_caption.text = note[1]
        
        xml_file = f'{self.course_title}_narration.xml'
        reparsed = minidom.parseString(gfg.tostring(the_course).decode('utf-8'))
        with open(xml_file, 'w') as f:
            f.write(reparsed.toprettyxml(indent='   '))
        print(f'{xml_file} written!')

    def mk_narration_txt(self, *args):
        txt_file = f'{self.course_title}_narration.txt'
        with open(txt_file, 'w') as f:
            for note in self.get_notes(*args):
                f.write(f'{note}\n\n')
        
        print(f'{txt_file} written!')
        return True


# custom slide filters
def filter_kc(slide_notes, slide_text, slide_num):
    # remove knowlege check notes slides
    if slide_text.lower().replace(' ', '').startswith('knowledgecheck'):
        print(f'Slide {str(slide_num)} | Knowledge Check skipped: {slide_notes}')
        return ''
    return slide_notes

def filter_ai(slide_notes, _, slide_num):
    # omit slide notes after "Aditional Information" in notes
    if 'Additional Information' in slide_notes:
        split_notes = slide_notes.split('Additional Information')
        if split_notes[0].replace('\n', ''):
            print(f'Slide {slide_num} | Additional Information skipped: ' + split_notes[1].replace('\n', ''))
            return split_notes[0].replace('\n', '')
        return ''
    return slide_notes


if __name__ == '__main__':
    pres_file = r'C:\Users\wbuehl\Documents\python_stuff\powerpoint_automation\SMA-HQ-WBT-108.pptx'
    #pres_file = r'C:\Users\wbuehl\Documents\python_stuff\powerpoint_automation\SMA-SS-WBT-0013_RIDM.pptx'

    course = Course(pres_file, 'HQ108')
    # course.mk_narration_txt(filter_kc, filter_ai)
    # course.mk_narration_docx(filter_kc, filter_ai)
    section_headers = [
        ('Introduction', 1),
        ('COPV Basic', 6),
        ('COPV Damage and Testing', 15),
        ('Safety', 28),
        ('Damage Essentials', 37),
        ('Summary', 45)
    ]
    course.mk_narration_xml(filter_kc, filter_ai, section_headers=section_headers, course_menu=True)
    # course.mk_narration_docx(filter_kc, filter_ai)
