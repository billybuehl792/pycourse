#!python
# read_pptx.py - read content from powerpoints

import os
from pptx import Presentation
from docx import Document


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
        pptx = []
        for n, _ in enumerate(self.pres):
            pptx.append({
                'slide_number': n + 1,
                'slide_text': self.get_slide_text(n+1),
                'slide_notes': self.get_slide_notes(n+1)
            })
        return pptx

    def get_notes(self):
        narration = []
        
        for n, _ in enumerate(self.pres.slides):
            slide_text = self.get_slide_text(n+1)
            slide_notes = self.get_slide_notes(n+1)
            if slide_notes:
                # omit knowlege checks
                if slide_text.lower().replace(' ', '').startswith('knowledgecheck'):
                    print(f'knowledge check skipped: {slide_notes}')
                    continue
                    
                # omit slide notes after "Aditional Information"
                if 'Additional Information' in slide_notes:
                    print(slide_notes)
                    split_notes = slide_notes.split('Additional Information')
                    if split_notes[0].replace('\n', ''):
                        narration.append(split_notes[0].replace('\n', ''))
                    print(split_notes[1].replace('\n', ''))
                    continue
                    
                narration.append(slide_notes.replace('\n', ''))

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

        return None


def mk_narration_docx(course_id, file_id, course_title, notes):
    doc = Document()
    doc.add_heading(f'{course_title} - Narration Script', level=1)
    table = doc.add_table(rows=0, cols=2)
    for n, note in enumerate(notes):
        row_cells = table.add_row().cells
        row_cells[0].text = f'{file_id}_{n}'
        row_cells[1].text = note
    
    doc_file = f'{course_title} narration script_01.docx'
    doc.save(doc_file)
    return doc_file


def mk_narration_txt(course_title, notes):
    txt_file = f'{course_title}_narration.txt'
    with open(txt_file, 'w') as f:
        for note in notes:
            f.write(f'{note}\n\n')
    
    print(f'{txt_file} written!')
    return True

def mk_narration_xml(course_id, file_id, course_title, pptx_dict, section_headers=None):
    # create xml file with notes
    if not section_headers:
        section_headers = []


if __name__ == '__main__':
    pres_file = r'C:\Users\wbuehl\Documents\python_stuff\powerpoint_automation\SMA-HQ-WBT-108.pptx'
    # pres_file = r'C:\Users\wbuehl\Documents\python_stuff\powerpoint_automation\SMA-SS-WBT-0013_RIDM.pptx'

    course = Course(pres_file, 'HQ108')
    course_notes = course.get_notes()
    mk_narration_docx(course.course_id, course.file_id, course.course_title, course_notes)
