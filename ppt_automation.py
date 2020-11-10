from pptx import Presentation
from pptx.util import Inches,Pt
import os

prs = Presentation("template.pptx")
week=input('몇 주차 과제임? : ')
pleft=Inches(1)
ptop=Inches(1.6)
pwidth=Inches(7.5)
pheight=Inches(5.5)
j=1
for i in os.listdir('images'):
    full_path = os.path.join('images\\', i)
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    title=slide.shapes.title
    title.text='예제 '+ os.path.splitext(i)[0]
    slide_num=slide.placeholders[1]
    slide_num.text=str(j)
    pic = slide.shapes.add_picture(full_path,pleft,ptop,pwidth,pheight)
    j+=1

filename=str(week)+'주차 김주연.pptx'
prs.save(filename)
