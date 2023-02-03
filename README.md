# Automate Microsoft Power Point Content with Python
## Introduction
This project can help to automate microsoft power point content, such as images, tables, chart with python. If you need more details customization that my source code don't provide, can check it in official [documentation](https://python-pptx.readthedocs.io/en/latest/).</br>
## Solution
#### Install python pptx library
```
pip install python-pptx
```
#### Import all the necessary libraries
```
import collections 
import collections.abc
from datetime import datetime
from pptx.enum.text import PP_ALIGN
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Pt
currencies_ppt = Presentation('oldppt.pptx')
```
#### Load your old PPT
```
currencies_ppt = Presentation('oldppt.pptx')
```
If you have error related to unixcode, try to put your file full path.
```
currencies_ppt = Presentation(r'C:\Users\oldppt.pptx')
```
#### Get your content ID inside PPT, this ID will be used to edit the selected content.
```
slide1 = currencies_ppt.slides[0]
for shape in slide1.shapes:
    print(shape.shape_id)
    print(shape.shape_type)
```
#### Edit your selected content and change shape_id with the ID you got from previous step.
```
 if shape.shape_id == 7 :
        text_frame = shape.text_frame
        text_frame.clear()
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = "Automating x\nWeekly Report\nWith Python\nStanley"
        font = run.font
        font.name = 'Segoe UI'
        font.size = Pt(32)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme
        font.color.rgb = RGBColor(0x10, 0x0c, 0x08)
```
