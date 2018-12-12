from docx import Document
import docx
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn


document = Document()

document.styles["Normal"].font.name = "Times New Roman"
document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
document.styles["Normal"].paragraph_format.space_after = docx.shared.Pt(1)

document.styles["Default Paragraph Font"].font.name = "Times New Roman"
document.styles['Default Paragraph Font']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

#header = document.sections[0].header
#header.text = "论文收录引用检索证明报告         2720"
#print(header.style.name)

print(document.sections[0].start_type)

p = document.add_paragraph("")
print(p.style.name)
print(p.style.type)
print(p.style.font.name)
p.add_run("附件一：").bold = True


p = document.add_paragraph("")

run = p.add_run("美国《科学引文索引》（SCI-EXPANDED）收录情况")
print(run.style.name)
print(run.style.font.name)
run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
run.bold = True

#设置引文段落格式
style = document.styles.add_style("yinwen", WD_STYLE_TYPE.PARAGRAPH)
style.paragraph_format.left_indent  = docx.shared.Cm(1)
style.font.name = "Times New Roman"
style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

#设置标题样式
heading_style = document.styles.add_style("sci_heading", WD_STYLE_TYPE.CHARACTER)
heading_style.font.bold = True
heading_style.font.underline = True
heading_style.font.name = "Times New Roman"
heading_style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
for i in range(1,5):
    p = document.add_paragraph("")
    p.add_run(str(i)+". ")
    p.add_run("被引文献",style="sci_heading")
    p.add_run(":")



    p = document.add_paragraph("",style="yinwen")
    p.add_run("引用文献",style="sci_heading")
    p.add_run("  自引").bold = True
    p.add_run("    ",style="sci_heading")
    p.add_run("篇  ").bold = True

    p.add_run("他引").bold = True
    p.add_run("    ", style="sci_heading")
    p.add_run("篇  ").bold = True

    p = document.add_paragraph("Title: Observer-based sliding mode control for a class of discrete systems via delta operator approach",style="yinwen")
    print(p.style.name)

document.save(r"c:\users\wangxiaoshan\desktop\wxs_py\header_test.docx")

print("success")
