from docx import Document

doc = Document('content.docx')

in_learning_objective = False
paragraphs_to_remove = []

for para in doc.paragraphs:
    text = para.text.strip()
    if text == '[LearningObjective]':
        in_learning_objective = True
        continue
    elif text.startswith('['):
        in_learning_objective = False
        
    if in_learning_objective:
        # clear the text of the paragraph instead of completely deleting the paragraph
        para.text = ""

# Since we clear it, we might want to just find the [LearningObjective] paragraph and insert after it,
# but it's easier to just build a new docx or just clear the text and add after the header.

doc.save('content._temp.docx')
