#!python
# read_pptx.py - read content from powerpoints

import json
import re
import sys
from pptx import Presentation
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.shared import OxmlElement
from lxml import etree


class Slide:

    def __init__(self, course, slide_id=None, slide_num=None):
        self.course = course
        self._slide_id = slide_id
        self._slide_num = slide_num
        self.slide = self.course.pres.slides.get(self.slide_id, default=None)

    @property
    def slide_id(self):
        if not self._slide_id:
            if self._slide_num:
                return self.course.slide_ids[self._slide_num - 1]
            else:
                raise Exception('slide_id or slide_num not provided')
        else:
            return self._slide_id

    @property
    def slide_num(self):
        if not self._slide_num:
            if self._slide_id:
                return self.course.slide_ids.index(self._slide_id) + 1
            else:
                raise Exception('slide_id or slide_num not provided')
        else:
            return self._slide_num

    @property
    def slide_type(self):
        # return slide's type
        if self.slide_id in self.course.title_slides:
            return 'title'
        elif self.slide_id in self.course.section_header_slides:
            return 'section_header'
        elif self.slide_id in self.course.menu_slides:
            return 'menu'
        else:
            return 'standard'
    
    @property
    def slide_text(self):
        # return text in slide
        text_runs = []
        slide = self.slide
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text_runs.append(self.format_string(run.text))

        return text_runs
    
    @property
    def slide_notes(self):
        # return slide notes section (narration)
        if self.slide.has_notes_slide:
            slide_notes = self.slide.notes_slide.notes_text_frame.text
            return self.format_string(slide_notes)

        return ''

    @staticmethod
    def format_string(string):
        string = string.replace('\n', ' ')
        string = string.replace('’', "'")
        string = string.replace('‘', "'")
        string = string.replace('  ', ' ')
        string = string.replace('“', '"')
        string = string.replace('”', '"')
        string = string.replace('–', '-')
        string = string.replace('—', '-')
        if string.endswith('\n'):
            string = string[:-2]
        string = string.replace('‐', '')
        string = string.replace('…', '...')
        string = string.replace('˚', ' degrees ')
        if not string.replace(' ', ''):
            return ''
        return string
    
    def __repr__(self):
        return f'Slide({self.course.title}, {self.slide_id}, {self.slide_num})'


class Course:

    def __init__(self, pptx_file, course_id=None, file_id=None, course_title=None):
        self.pptx_file = pptx_file
        self._course_id = course_id
        self._file_id = file_id
        self._course_title = course_title
        self.pres = Presentation(self.pptx_file)
    
    @property
    def slide_ids(self):
        return [slide.slide_id for slide in self.pres.slides]

    @property
    def title_slides(self):
        title_layout_index = 0
        layout = self.pres.slide_master.slide_layouts[title_layout_index]
        return [slide.slide_id for slide in layout.used_by_slides]

    @property
    def section_header_slides(self):
        section_header_layout_index = 2
        layout = self.pres.slide_master.slide_layouts[section_header_layout_index]
        return [slide.slide_id for slide in layout.used_by_slides]
    
    @property
    def menu_slides(self):
        menu_slide_layout_index = 5
        layout = self.pres.slide_master.slide_layouts[menu_slide_layout_index]
        return [slide.slide_id for slide in layout.used_by_slides]

    @property
    def standard_slides(self):
        normal_slides = []
        for n, layout in enumerate(self.pres.slide_layouts):
            # if title slide or slide header
            if n in [0, 2, 5]:
                continue
            for slide in layout.used_by_slides:
                normal_slides.append(slide.slide_id)
        return normal_slides

    @property
    def has_menu(self):
        for s in self.slide_ids:
            slide = Slide(self, s)
            if ' '.join(slide.slide_text).replace(' ', '').lower().startswith('mainmenu'):
                return True
        return False

    @property
    def course_title(self):
        # title slide - first text of first slide of powerpoint
        if not self._course_title:
            slide = Slide(self, self.slide_ids[0])
            return(slide.slide_text[0])
        else:
            return self._course_title

    @property
    def course_id(self):
        # title slide - last text of first slide of powerpoint
        if not self._course_id:
            slide = Slide(self, self.slide_ids[0])
            return(slide.slide_text[-1])
        else:
            return self._course_id
    
    @property
    def file_id(self):
        # course_id - filtered text
        if not self._file_id:
            try:
                code = re.search('-(.*?)-', self.course_id).group(1)
                num = self.course_id[self.course_id.rfind('-')+1:]
                return f'{code}{num}'
            except Exception:
                return 'FILE_ID_ERROR'
        else:
            return self._file_id

    @property
    def course(self):

        pptx = {}   # return dictionary of pptx

        pptx['course_title'] = self.course_title
        pptx['course_id'] = self.course_id
        pptx['study_guide_pdf'] = f'{self.course_id}_508.pdf'
        pptx['study_guide_print'] = f'{self.course_id}_StudyGuide.pdf'
        pptx['sections'] = [{'section_title': 'Introduction', 'slides': []}]


        for slide_id in self.slide_ids:
            slide = Slide(self, slide_id)
            text = slide.slide_text
            if slide.slide_type == 'section_header':
                section = {'section_title': text[0], 'slides': []}
                pptx['sections'].append(section)
            pptx['sections'][-1]['slides'].append(slide_id)

        return pptx

    def write_json(self):
        # write course dict to json
        filename = f'{self.course_id}.json'
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.course, f, ensure_ascii=False, indent=4)
            print(f'JSON written: {filename}')
        return filename

    def write_docx(self, *args):

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

        for item in self.slide_ids:
            # set slide
            slide = Slide(self, item)

            # skip main menu
            if slide.slide_type == 'menu' or slide.slide_type == 'section_header':
                continue
            
            # apply filters
            narration = slide.slide_notes
            for func in args:
                narration = func(slide)
            if not narration.replace(' ', ''):
                continue
            
            # add narration to table
            if narration:
                row_cells = table.add_row().cells
                row_cells[0].text = f'{self.file_id}_{slide.slide_num}'
                row_cells[1].text = narration

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
        doc_file = f'{self.course_id}_narration.docx'
        try:
            doc.save(doc_file)
            print(f'DOCX written: {doc_file}')
        except Exception:
            print(f'{doc_file} exists! or invalid permissions!')
            return False
        return doc_file

    def write_xml(self, *args):
        # write course narration XML
        # course xml
        the_course = etree.Element('theCourse')
        the_course.append(etree.Comment('Generated with pycourse!'))
        etree.SubElement(the_course, 'myCourseTitle').text = self.course_title
        etree.SubElement(the_course, 'studyGuidePDF').text = f'{self.course_id}.pdf'
        etree.SubElement(the_course, 'studyGuidePrint').text = f'{self.course_id}_StudyGuide.pdf'

        menu = 'NO'
        if self.has_menu:
            menu = 'YES'

        etree.SubElement(the_course, 'courseMenu').text = menu

        for n, section in enumerate(self.course.get('sections')):
            # theSections elem
            the_sections = etree.SubElement(the_course, 'theSections', {'title': section.get('section_title')})
            
            file_num = 0
            for s in section.get('slides'):
                # slide
                slide = Slide(self, s)
                
                # skip main menu
                if slide.slide_type == 'menu' or slide.slide_type == 'section_header':
                    continue

                # apply filters
                narration = slide.slide_notes
                for func in args:
                    narration = func(narration, slide.slide_text, slide.slide_num, slide.slide_type)

                # sectionNumber elem
                section_num = etree.SubElement(the_sections, 'sectionNumber')

                # theFileToLoad elem
                the_file_to_load = etree.SubElement(section_num, 'theFileToLoad')
                the_file_to_load.text = f'{self.file_id}_s{n}_{file_num+1}.html'

                # closed caption elem
                closed_caption = etree.SubElement(section_num, 'closedCaptionText')
                
                closed_caption.text = narration

                # increment filenum
                file_num += 1

        xml_file = f'{self.course_id}_narration.xml'
        tree_string = etree.tostring(the_course, pretty_print=True).decode('utf-8')
        with open(xml_file, 'w') as f:
            f.write(tree_string)
            print(f'XML written: {xml_file}')

        return xml_file

    def write_txt(self):
        txt_file = f'{self.course_id}_narration.txt'
        with open(txt_file, 'w') as f:
            for item in self.slide_ids:
                slide = Slide(self, item)
                if slide.slide_notes:
                    note = slide.slide_notes
                    f.write(f'{note}\n\n')
            print(f'TXT written: {txt_file}')

        return txt_file

    def __repr__(self):
        return f'Course({self.course_id}, {self.course_title})'


# custom slide filters
def filter_kc(slide_notes, slide_text, slide_num, slide_type):
    # Capitalize and add knowledge check narration in () knowlege check notes slides
    del slide_type
    if ' '.join(slide_text).lower().replace(' ', '').startswith('knowledgecheck'):
        print(f'Slide {str(slide_num)} | Knowledge Check skipped: {slide_notes}')
        return ''
        return f'KNOWLEDGE CHECK ({slide_notes})'
    return slide_notes

def filter_ai(slide_notes, slide_text, slide_num, slide_type):
    # omit slide notes after "Aditional Information" in notes
    del slide_text, slide_type
    if 'Additional Information' in slide_notes:
        split_notes = slide_notes.split('Additional Information')
        if split_notes[0].replace('\n', ''):
            print(f'Slide {slide_num} | Additional Information skipped: ' + split_notes[1].replace('\n', ''))
            return split_notes[0].replace('\n', '')
        return ''
    return slide_notes

def to_caps(slide_notes, slide_text, slide_num, slide_type):
    del slide_text, slide_num, slide_type
    return slide_notes.upper()


if __name__ == '__main__':
    
    # get sys arguments
    # l = len(sys.argv)
    # if l <= 1:
    #     print('Provide arguments: <pptx file> <course_id:optional> <course_title:optional> <file_id:optional>')
    #     sys.exit()
    # elif l >= 2:
    #     pptx_file = sys.argv[1]
    #     course_id = None
    #     file_id = None
    #     course_title = None
    #     if l >= 3:
    #         course_id = sys.argv[2]
    #         if l >= 4:
    #             course_title = sys.argv[3]
    #             if l >= 5:
    #                 file_id = sys.argv[4]
    
    pptx_file = r'testFiles/SMA-HQ-WBT-108.pptx'
    course_id = 'SMA-SS-WBT-400'
    course_title = 'System Safety Analysis Relationships with Single Point of Failure Analysis'
    file_id = 'SS400'

    # course object
    course = Course(pptx_file, course_id=course_id, course_title=course_title, file_id=file_id)

    # write files
    course.write_docx(filter_kc, filter_ai)
