from docx import Document

import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
import tkinter
from tkinter import ttk
import threading
import tkinter.filedialog
import tkinter.messagebox


import pandas as pd
#from docx.enum.text import WD_ALIGN_PARAGRAPH

recs = pd.read_excel(r"c:\cnki.xlsx", 0)
i = 0
while i < recs.shape[0]:
    journal = recs['文献来源'][i]
    print(journal)