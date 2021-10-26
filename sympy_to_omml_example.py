from docx import Document
from math_to_word import math_to_word
from sympy.abc import x


doc = Document()

p = doc.add_paragraph()
p._element.append(math_to_word(x+1, equal=10))
doc.save('test1.docx')