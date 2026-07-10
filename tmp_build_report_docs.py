from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT=Path('output')
OUT.mkdir(exist_ok=True)

YELLOW='FFFF00'

def font(run,size=12,bold=False,highlight=True):
    run.font.name='宋体'
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体')
    run.font.size=Pt(size)
    run.bold=bold
    if highlight:
        run.font.highlight_color='YELLOW'

def add(doc,text,size=12,bold=False):
    p=doc.add_paragraph()
    p.paragraph_format.line_spacing=1.3
    r=p.add_run(text)
    font(r,size,bold)
    return p

def heading(doc,text):
    p=doc.add_paragraph()
    r=p.add_run(text)
    font(r,14,True)
    return p

def table(doc,headers,rows):
    t=doc.add_table(rows=1,cols=len(headers))
    t.style='Table Grid'
    for c,h in zip(t.rows[0].cells,headers):
        c.text=h
    for row in rows:
        cells=t.add_row().cells
        for c,v in zip(cells,row): c.text=v
    return t

# document 1
D=Document()
D.add_paragraph('铁路单北斗差分定位算法优化与应用研究').alignment=WD_ALIGN_PARAGRAPH.CENTER
add(D,'结题报告企业验收优化修订稿（标黄修改版）',14,True)
add(D,'说明：本稿用于替换原结题报告中需要提升的摘要、逻辑衔接、成果总结、工程应用表达和格式规范部分；原有实验数据、公式和图表主体内容保持不变。',11)

heading(D,'一、内容摘要（优化版）')
add(D,'本课题面向铁路行业单北斗规模化应用需求，围绕自主可控高精度定位关键技术开展研究。针对现有铁路地基增强系统依赖多系统GNSS、复杂环境下单北斗定位可靠性不足等问题，开展单北斗时空基准评估、服务端单北斗改造、铁路电离层建模与闪烁监测、国产接收机信号畸变偏差建模改正以及工程示范应用研究。课题形成了覆盖时空基准、误差建模、质量监测、定位服务和工程验证的完整技术链。研究成果实现了单北斗差分定位服务能力提升，形成发明专利3项、论文3篇及铁路单北斗应用验证成果，为集团公司后续北斗规模应用和铁路高精度位置服务提供技术支撑。')

heading(D,'二、主要修改内容及验收逻辑优化')
table(D,['位置','原问题','修改方向'],[
['摘要','偏论文总结，工程价值不足','增加合同目标、工程应用和成果产出闭环'],
['第一章','背景描述较宏观','突出铁路单北斗规模应用需求'],
['第二章','综述较多','增加已有技术不足与本课题切入点'],
['第三至第五章','技术成果分散','强化形成的软件、算法和平台能力'],
['第六章','实验结果描述不足','增加工程示范意义和推广价值'],
['第七章','总结偏中期报告','改为合同指标完成和成果评价']])

heading(D,'三、第二章新增：现有技术不足与本课题突破点')
add(D,'现有铁路北斗增强服务主要关注基准站建设、差分定位算法和多源融合定位，但在单北斗自主运行、复杂铁路环境适应、国产终端质量控制和工程化服务体系方面仍存在不足。本课题围绕单北斗时空基准、准四维电离层模型、信号畸变偏差改正和铁路示范应用形成技术突破，实现从算法研究到工程服务能力构建。')

heading(D,'四、章节工程贡献强化')
for x in [
'第三章：形成基于北斗时空基准的单北斗地基增强服务能力，实现CORS服务端处理链路自主化。',
'第四章：形成铁路区域电离层精化建模和闪烁监测方法，为复杂电离层环境下质量控制提供支撑。',
'第五章：形成国产单北斗接收机信号畸变偏差提取与改正方法，提高高精度定位可靠性。',
'第六章：通过南昌局示范验证，证明单北斗铁路地基增强服务具备工程应用条件。']:
    add(D,x)

heading(D,'五、第七章结题总结（优化版）')
add(D,'本课题围绕铁路单北斗差分定位算法优化与应用开展研究，完成了单北斗时空基准评估、CORS服务端改造、电离层建模与监测、接收机信号畸变偏差改正以及工程示范验证等研究任务。课题形成了面向铁路单北斗应用的关键算法、软件模块和验证体系，实现了从理论研究到工程应用的技术转化。')
table(D,['合同任务','完成情况'],[
['单北斗时空基准与差分定位研究','完成'],['CORS服务端单北斗改造','完成'],['铁路电离层建模与闪烁监测','完成'],['信号畸变偏差建模与改正','完成'],['铁路工程示范应用','完成'],['专利论文成果','完成']])

# document 2
M=Document()
M.add_paragraph('铁路单北斗差分定位算法优化与应用研究').alignment=WD_ALIGN_PARAGRAPH.CENTER
add(M,'结题报告修改说明（企业验收版）',14,True)
heading(M,'一、总体修改原则')
add(M,'围绕企业科研课题验收要求，对原报告进行“合同指标对应化、成果工程化、语言规范化、格式统一化”修改。重点解决原报告存在的论文化表达明显、工程贡献突出不足、章节衔接弱、成果总结与合同目标对应不足等问题。')
heading(M,'二、重点修改清单')
table(M,['类别','修改内容','目标'],[
['内容','重写摘要、总结和成果评价','体现企业验收逻辑'],
['逻辑','增加问题-方法-成果-应用链条','增强章节关联'],
['成果','增加软件、平台、工程能力描述','突出转化价值'],
['格式','统一标题、图表、术语、引用','提升报告规范性'],
['应用','强化南昌局示范工程意义','突出推广价值']])
heading(M,'三、格式规范调整')
add(M,'统一正文中文字体、标题层级、图表标题格式、章节编号、公式格式和术语表达；检查重复编号、引用错误和章节名称一致性；对新增或修改文字全部采用黄色高亮标识。')
heading(M,'四、验收提升目标')
add(M,'修改后报告由“科研论文型总结”提升为“企业科研项目结题验收型报告”，突出合同完成度、技术创新性、工程应用价值和后续推广路径。')

D.save(OUT/'铁路单北斗差分定位算法优化与应用研究_结题报告企业验收优化修订稿_标黄版.docx')
M.save(OUT/'铁路单北斗差分定位算法优化与应用研究_修改说明.docx')
print('done')
