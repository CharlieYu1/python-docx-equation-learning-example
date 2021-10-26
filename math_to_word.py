# https://github.com/python-openxml/python-docx/issues/320
from lxml import etree
from sympy.printing.mathml import mathml
from sympy import *


mml2omml_stylesheet_path = 'MML2OMML.XSL'


def math_to_word(eq, equal=None):
    """Transform a sympy equation to be printed in word document."""
    # Mathml uses sympy to transform equations to mathml format
    # eq = equation you want to add to Word using sympy format
    # The math_to_word functions also have an optional equal argument
    # that prints an = sign and the number passed to the argument

    math_ml = mathml(eq, printer='presentation')
    # Creates mathml string
    if equal is None:
        mathml_string = '''
            <math xmlns="http://www.w3.org/1998/Math/MathML">
                {}
            </math>
            '''.format(math_ml)
    else:
        mathml_string = '''
            <math xmlns="http://www.w3.org/1998/Math/MathML">
                {0}
                <mo>=</mo>
                <mn>{1}</mn>
             </math>
            '''.format(math_ml, equal)
    # Converts mathml string
    tree = etree.fromstring(mathml_string)
    xslt = etree.parse(
        mml2omml_stylesheet_path
        )
    transform = etree.XSLT(xslt)
    new_dom = transform(tree)
    return new_dom.getroot()