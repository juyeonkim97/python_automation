from pptx import Presentation
from pptx.util import Inches,Pt
import os

prs = Presentation()
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)
ch=input('input your chapter :')
week=input('input week: ')
left = Inches(0.7)
top = Inches(0.2)
width = height = Inches(1)
pleft=Inches(0.7)
ptop=Inches(1.5)
pwidth=Inches(10)
pheight=Inches(7)
for i in os.listdir('images'):
    full_path = os.path.join('images\\', i)
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    txBox = slide.shapes.add_textbox(left, top,width,height)
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = '예제 '+ os.path.splitext(i)[0]
    p.font.size = Pt(33)
    pic = slide.shapes.add_picture(full_path,pleft,ptop,pwidth,pheight)

filename=str(week)+'주차 김주연.pptx'
prs.save(filename)
