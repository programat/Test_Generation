import docx
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from docx.shared import Pt, Inches
from lxml import etree
import latex2mathml.converter

def printToMathml(paragraph,formula):
    stri = latex2mathml.converter.convert(formula)
    tree = etree.fromstring(stri)
    xslt = etree.parse('MML2OMML.XSL')
    transform = etree.XSLT(xslt)
    func = transform(tree)
    paragraph._element.append(func.getroot())

