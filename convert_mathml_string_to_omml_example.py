from docx import Document
from lxml import etree


doc = Document()
mml2omml_stylesheet_path = 'MML2OMML.XSL'

# MathML string to parse
mathml_string = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mn>1</mn><mn>2</mn></mfrac></math>'

# Convert MathML into OfficeML using XSLT stylesheet
mathml_tree = etree.fromstring(mathml_string)
xslt = etree.parse(mml2omml_stylesheet_path)

xslt_transform = etree.XSLT(xslt)
equation_dom = xslt_transform(mathml_tree)

paragraph = doc.add_paragraph()
paragraph._element.append(equation_dom.getroot())

doc.save('test.docx')