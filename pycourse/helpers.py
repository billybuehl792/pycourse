# custom slide filters

# return KNOWLEDGE CHECK (slide_notes) if knowledge check slide
def filter_kc(slide_notes, slide_text, slide_num, slide_type):
    del slide_type

    if ' '.join(slide_text).lower().replace(' ', '').startswith('knowledgecheck'):
        print(f'Slide {str(slide_num)} | Knowledge Check skipped: {slide_notes}')
        return f'KNOWLEDGE CHECK ({slide_notes})'

    return slide_notes

# return slide notes before "Aditional Information" in slide notes
def filter_ai(slide_notes, slide_text, slide_num, slide_type):
    
    del slide_text, slide_type

    if 'Additional Information' in slide_notes:
        split_notes = slide_notes.split('Additional Information')
        if split_notes[0].replace('\n', ''):
            print(f'Slide {slide_num} | Additional Information skipped: {split_notes[1].replace("\n", "")}')
            return split_notes[0].replace('\n', '')
        return ''

    return slide_notes

# return slide notes but capitalized
def to_caps(slide_notes, slide_text, slide_num, slide_type):

    del slide_text, slide_num, slide_type

    return slide_notes.upper()
