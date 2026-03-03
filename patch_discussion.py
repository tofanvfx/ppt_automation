with open('docx_to_ppt.py', 'r', encoding='utf-8') as f:
    text = f.read()

changes = [
    # 1. Routing 
    ('''
    if name.replace("_", "").lower() in ('homeworkpage', 'ssthomeworkpage', 'ssthomework', 'homework'):
        for layout in prs.slide_layouts:
            if layout.name == 'LAYOUT_sst_homework_page':
                return layout

    return None''',
    '''
    if name.replace("_", "").lower() in ('homeworkpage', 'ssthomeworkpage', 'ssthomework', 'homework'):
        for layout in prs.slide_layouts:
            if layout.name == 'LAYOUT_sst_homework_page':
                return layout

    if name.replace("_", "").lower() in ('discussionpage', 'sstdiscussionpage', 'sstdiscussion', 'sst_discussion_page'):
        for layout in prs.slide_layouts:
            if layout.name == 'LAYOUT_sst_discussion_page':
                return layout

    return None'''),

    # 2. Add question1 matching in loop
    ('''
                txt = shape.text.lower()
                if layout.name.startswith('LAYOUT_sst_quiztime_page'):
                    if 'quiz time' in txt:
                        templates['title'] = copy.deepcopy(shape.element)
                    elif 'question' in txt:
                        templates['question'] = copy.deepcopy(shape.element)
                    elif 'options' in txt:
                        templates['options'].append(copy.deepcopy(shape.element))
                    elif txt.strip(): 
                        if not templates['title']: templates['title'] = copy.deepcopy(shape.element)
                else:
                    templates['text'] = copy.deepcopy(shape.element)
                ''',
    '''
                txt = shape.text.lower()
                if layout.name.startswith('LAYOUT_sst_quiztime_page'):
                    if 'quiz time' in txt:
                        templates['title'] = copy.deepcopy(shape.element)
                    elif 'question' in txt:
                        templates['question'] = copy.deepcopy(shape.element)
                    elif 'options' in txt:
                        templates['options'].append(copy.deepcopy(shape.element))
                    elif txt.strip(): 
                        if not templates['title']: templates['title'] = copy.deepcopy(shape.element)
                elif layout.name == 'LAYOUT_sst_discussion_page':
                    if 'question1' in txt:
                        templates['question1'] = copy.deepcopy(shape.element)
                else:
                    templates['text'] = copy.deepcopy(shape.element)
                '''),

    # 3. Add dictionary to store mapping
    ('''    'LAYOUT_sst_quiztime_page_02': {'title': None, 'question': None, 'options': [], 'picture': None}
}''',
     '''    'LAYOUT_sst_quiztime_page_02': {'title': None, 'question': None, 'options': [], 'picture': None},
    'LAYOUT_sst_discussion_page': {'question1': None}
}'''),

    # 4. Inject discussion data
    ('''            replace_text_preserve_format(shape, processed_content)
            print(f"Updated {layout.name} Body text with preserved formatting and spacing.")

if __name__ == "__main__":''',
     '''            replace_text_preserve_format(shape, processed_content)
            print(f"Updated {layout.name} Body text with preserved formatting and spacing.")

    elif layout.name == 'LAYOUT_sst_discussion_page':
        templates = sst_content_templates.get(layout.name, {})
        if templates.get('question1') is not None:
            # Parse 'question1: <text>' from docx
            q1_text = ""
            for entry in section['content']:
                line = get_text(entry)
                if ":" in line:
                    parts = line.split(":", 1)
                    if parts[0].strip().lower() == 'question1':
                        q1_text = parts[1].strip()
                        break
                elif line.strip() and not q1_text:
                    q1_text = line.strip()  # Fallback to first line
            
            q1_elem = copy.deepcopy(templates['question1'])
            slide.shapes._spTree.append(q1_elem)
            shape = slide.shapes[-1]
            if q1_text:
                replace_text_preserve_format(shape, q1_text)
                print(f"Updated {layout.name} question1 -> '{q1_text[:20]}...'")
            else:
                replace_text_preserve_format(shape, "question1")  # fallback
                print(f"Inserted {layout.name} with default text")

if __name__ == "__main__":''')
]

for old, new in changes:
    text = text.replace(old.lstrip('\n'), new.lstrip('\n'))

with open('docx_to_ppt.py', 'w', encoding='utf-8') as f:
    f.write(text)
