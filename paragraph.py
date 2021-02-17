# -*- coding:utf-8 -*-

from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.oxml.shared import qn


def runs(self):
    """
    Sequence of |Run| instances corresponding to the <w:r> elements in
    this paragraph.
    """
    lst = [Run(r, self) for r in self._p.r_lst]
    for hl in self._p:
        if hl.tag == qn('w:hyperlink'):
            for r in hl:
                if r.tag == qn('w:r'):
                    lst.append(Run(r, self))
    return lst

Paragraph.runs = property(runs)