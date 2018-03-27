##
##  Copy & Paste Tool for images to PowerPoint(.pptx)
##
from pptx import Presentation
import pptx.util
import glob
import scipy.misc

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[8])
placeholder = slide.placeholders[1]  # idx key, not position
placeholder.placeholder_format.type
picture = placeholder.insert_picture('Capture.PNG')
prs.save('hello.pptx')