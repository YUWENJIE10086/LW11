# -*- coding: utf-8 -*-
"""V4: 原图、分层、逐字段、并列关联；基于V3及9.29讨论稿/11.17现场表单。"""
from pathlib import Path
import re,json,zipfile,hashlib
from PIL import Image
from docx import Document
from docx.shared import Cm,Pt,RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path("低代码1117"); SHOT=ROOT/"解压截图_11172053"
STAMP="11181139_V4"
OUT=ROOT/f"烤房四大部位信息采集操作指南_逐项填写原图完善版_{STAMP}.docx"
ref=Document(ROOT/"烤房项目基础信息录入指南11.13_1794919896877.docx")
doc=Document()
sec=doc.sections[0]; orig=ref.sections[0]
for a in ("page_width","page_height","top_margin","bottom_margin","left_margin","right_margin"):
    setattr(sec,a,getattr(orig,a))
sec.footer_distance=Cm(1.2)
CW=(sec.page_width-sec.left_margin-sec.right_margin)/360000
CH=(sec.page_height-sec.top_margin-sec.bottom_margin)/360000
manifest=json.loads((SHOT/"_截图明细_11172053.json").read_text(encoding="utf-8"))
imgs={int(x["filename"][:3]):SHOT/x["filename"] for x in manifest}
menu=ROOT/"6e408f052c2c854c097659b8a008a62a.jpg"
assert len(imgs)==41 and menu.exists()
normal=doc.styles["Normal"]
normal.font.name="仿宋_GB2312";normal.font.size=Pt(14)
normal._element.rPr.rFonts.set(qn("w:eastAsia"),"仿宋_GB2312")
normal.paragraph_format.line_spacing=1.16
normal.paragraph_format.space_after=Pt(4)
RED=(192,0,0); BLACK=(0,0,0)
def run(p,t,b=False,sz=14,font="仿宋_GB2312",color=None):
    r=p.add_run(t);r.font.name=font;r.font.size=Pt(sz);r.bold=b
    r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"),font)
    if color:r.font.color.rgb=RGBColor(*color)
    return r
def para(t="",sz=14,indent=False):
    p=doc.add_paragraph()
    if indent:p.paragraph_format.first_line_indent=Pt(28)
    run(p,t,sz=sz)
    return p
def line(name,txt="",red=False,sz=14):
    p=doc.add_paragraph()
    p.paragraph_format.space_after=Pt(4)
    run(p,name+"：",True,sz,color=RED if red else None)
    run(p,txt,red,sz,color=RED if red else None)
    return p
def head(t,level=1):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(0 if level==1 else 5)
    p.paragraph_format.space_after=Pt(6)
    p.paragraph_format.keep_with_next=True
    run(p,t,True,16 if level==1 else 14)
    return p
def shade(cell,fill):
    tcPr=cell._tc.get_or_add_tcPr()
    s=OxmlElement("w:shd");s.set(qn("w:fill"),fill);tcPr.append(s)
def row_keep(row):
    trPr=row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))
def set_cols(t,widths):
    t.autofit=False;t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for col,w in zip(t.columns,widths):col.width=Cm(w)
    for row in t.rows:
        for cell,w in zip(row.cells,widths):cell.width=Cm(w)
def summary(rows):
    t=doc.add_table(rows=1,cols=3);t.style="Table Grid"
    for cell,s in zip(t.rows[0].cells,["展开栏目","需填写字段","填写口径"]):
        q=cell.paragraphs[0];q.paragraph_format.space_after=Pt(1)
        run(q,s,True,11);shade(cell,"E7E6E6")
    for a,b,c in rows:
        row=t.add_row();row_keep(row)
        for cell,v in zip(row.cells,[a,b,c]):
            q=cell.paragraphs[0];q.paragraph_format.space_after=Pt(1)
            q.paragraph_format.line_spacing=1.09
            run(q,v,sz=10.5)
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cols(t,[2.62,4.55,CW-7.17])
    return t

page_count=0
fig=0; used=[]; figure_map={}
def page(t=None):
    global page_count
    if page_count:doc.add_page_break()
    page_count+=1
    if t:head(t)
def photo_run(p,path,h,wmax):
    with Image.open(path) as im:w,hpx=im.size
    ratio=w/hpx; ih=min(h,wmax/ratio)
    p.paragraph_format.space_after=Pt(1)
    p.paragraph_format.keep_with_next=True
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path),height=Cm(ih),width=Cm(ih*ratio))
def figures(pairs,height=14.6):
    """Each image is original JPEG embedded individually; never crop or merge."""
    global fig
    paths=[menu if i=="menu" else imgs[i] for i,_ in pairs]
    if len(pairs)==1:
        i,title=pairs[0]
        p=doc.add_paragraph()
        photo_run(p,paths[0],height,CW-0.4)
        q=doc.add_paragraph();q.alignment=WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.space_after=Pt(3)
        fig+=1; figure_map[str(i)]=fig
        run(q,f"图{fig}  {title}",sz=10.5)
        if i!="menu":used.append(i)
        return
    assert len(pairs)==2
    t=doc.add_table(rows=1,cols=2);set_cols(t,[CW/2,CW/2])
    row=t.rows[0];row_keep(row)
    for cell,(i,title),path in zip(row.cells,pairs,paths):
        cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
        photo_run(cell.paragraphs[0],path,height,CW/2-0.45)
        q=cell.add_paragraph();q.alignment=WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.space_after=Pt(0)
        q.paragraph_format.line_spacing=1.1
        fig+=1;figure_map[str(i)]=fig
        run(q,f"图{fig}  {title}",sz=10.5)
        if i!="menu":used.append(i)
    # No borders, side by side images remain editable independently.
    return t
def overview(title,pairs,rows,h=12.5,extra=None):
    page(title)
    figures(pairs,h)
    if extra:para(extra,12)
    summary(rows)
def fields(title,values,choice,photo=None,warning=None):
    head(title,2)
    line("填写字段",values)
    line("选择依据",choice)
    if photo:line("拍照要求",photo)
    if warning:line("注意",warning,True)
def done():
    line("填完后","核对本入口的红色*必填项，确认照片已上传，再点“提交”。成功后关闭表单，进入下一入口。",True,12)

# 1. 总界面：指定原图必须置于文首。
page()
t=doc.add_paragraph();t.alignment=WD_ALIGN_PARAGRAPH.CENTER
t.paragraph_format.space_after=Pt(6)
run(t,"烤房四大部位信息采集操作指南",True,22,"方正小标宋简体")
t=doc.add_paragraph();t.alignment=WD_ALIGN_PARAGRAPH.CENTER
run(t,"（逐项填写·原图完善版）",sz=14)
figures([("menu","低代码搜索数据页面及四个设备采集入口")],12.3)
para("按11.13版《烤房项目基础信息录入指南》找到烤房并完成“其他—信息完善”后，依次点击底部“主体门窗、加热设备、循环自控、附属设施”。")
line("必填与编号","红色*为必填。每次进入表单先核对“关联项目编号”，截图中的编号仅作示例，严禁选错烤房或串传照片。",True)
line("状态与年度","正常、损坏、缺失三类按现场实况判断；年度选该部件最近一次实际修复、改造的年份。未维修的按原建成/安装年度核实，不能一律选择2026年。")
line("照片标准","正常设备拍全景1张；损坏设备拍全景、瑕疵细节各1张；11.17页面要求缺失设备拍全景1张。无上传栏、只评价的部件不另拍。",True)

# 2. 主体门窗全表
overview("一、主体门窗",[(40,"首次进入主体门窗"),(19,"六个栏目全部收起")],[
 ("加热室外立面","状态、最新修复、改造年度、照片","左右清灰门、风机检修门及外立面。"),
 ("大门","状态、最新修复、改造年度、照片","含观察窗；检查整门及边框。"),
 ("装烟室","装烟室门状态、最新修复、改造年度、照片","检查地面、屋墙、挂烟梁。"),
 ("挂烟梁","材质","钢制挂烟梁／木制挂烟梁。"),
 ("综合评价","评价选择","优良／较好／较差。"),
 ("更新GPS信息","定位、经度、纬度","核对台账或现场实际坐标。")
],12.1,"进入后共有6个栏目。逐个点右侧箭头展开，按下表顺序填写。")

page("（一）加热室外立面")
fields("加热室外立面（包括左右清灰门和风机检修门）",
 "加热室外立面-状态*；加热室外立面-最新修复、改造年度*；加热室外立面-正常设备全景照片*（随所选状态显示照片栏）。",
 "墙体、清灰门、检修门完整可用选“正常”；有裂痕、破损、变形或不能正常使用选“损坏”；应有而无实物选“缺失”。点击年度栏滑动年份，选实有记录的年份后点“确定”。",
 "拍完整外立面，确保清晰看见左右清灰门及风机检修门；有损坏时另拍瑕疵细节。")
figures([(39,"展开外立面后选择最新修复、改造年度")],13.9)

page("照片采集方法（四个入口通用）")
line("选择相机","点击字段下方“拍照上传”，底部出现“系统相机、浏览器相机”。选择“系统相机”，点击右上角“确定”。")
line("实际拍照","系统打开手机相机，选择“拍照”，对准本烤房、本部件拍摄全景或瑕疵图；拍摄确认返回后，检查照片已显示在对应字段。")
line("重要","正常拍全景1张；损坏拍全景和瑕疵细节各1张；缺失按11.17新表单拍安装位置全景1张。9.29讨论稿原为缺失不拍照，以现行页面及最新通知为准。",True)
figures([(38,"选择系统相机并点击确定"),(37,"系统相机打开后进行拍照")],14.6)

page("（二）大门　（三）装烟室")
fields("（二）大门",
 "大门-状态*；大门-最新修复、改造年度*；大门-正常设备全景照片*。",
 "含观察窗。门扇、边框、观察窗、密封和开关正常选“正常”；破损、漏风或开关故障选“损坏”；整门缺失选“缺失”。",
 "拍摄大门及边框、观察窗全景；损坏另拍问题位置。")
fields("（三）装烟室",
 "装烟室门-状态*；装烟室-最新修复、改造年度*；装烟室-正常设备全景照片*。",
 "核查门、地面、屋墙及主体是否可正常烘烤；完整可用选正常，裂损漏雨选损坏，缺失按实际选。",
 "室内全景应体现挂烟梁及墙体、地面情况。")
figures([(36,"大门展开，装烟室栏目紧随其后"),(35,"装烟室字段和挂烟梁材质栏")],12.7)

page("（四）挂烟梁　（五）综合评价")
fields("（四）挂烟梁","材质*。","下拉显示“钢制挂烟梁、木制挂烟梁”，按现场材质选择。钢质选钢制，木质选木制；混合材质或缺失无法对应时先核实。",
 "挂烟梁随装烟室全景体现，无独立拍照栏。")
fields("（五）综合评价","评价选择*。","结合主体能否正常烘烤、墙体裂损、门窗和修复年份综合判断，点击评价选择框。")
figures([(34,"挂烟梁实际两项材质选项"),(33,"综合评价标准及GPS栏目入口")],13.5)

page("综合评价：按页面三级标准选择")
line("优良","2023—2026年建设维修，性能稳定，预计可使用5—8年。")
line("较好","2018—2022年建设维修，性能良好，仅需常规管护或少量部件更换，预计可使用3—5年。")
line("较差","2017年及以前建设维修，设备老化明显，需重点投入管护或部位更换，预计可使用1—3年。")
line("重要","主体门窗下拉为“优良、较好、较差”。9.29讨论稿曾采用四级评价，现行表单已改为三级，不照搬旧等级或仅凭年份选。",True,12)
figures([(32,"打开综合评价下拉选项"),(31,"确认主体门窗为优良、较好、较差")],14.4)

page("（六）更新GPS信息：获取当前位置")
line("填写字段","定位*、地址转换-获取经纬度、获取经纬度信息、经度*、纬度*。")
line("获取位置","在实际烤房现场打开手机定位权限，点“获取当前位置”。等待地址显示；正在获取时不要重复点。")
line("定位依据","行业基础信息表中的烤房，GPS以行业基础信息表为准；政府烤房按现场采集为准。发现台账与实地明显不符时先核实。")
line("限制","系统允许手工调整定位位置（500米内），仅可调整两次。不要在办公室或其他烤房定位。",True)
figures([(30,"初始定位页面及空白经纬度"),(29,"点击获取当前位置后等待定位")],14.0)

page("（六）更新GPS信息：地图定位及经纬度")
line("填写","定位地址出现后查看地图；放大地图、移动并点击可标注精确位置。确认地图位置后点击“获取经纬度信息”。")
line("查看结果","页面经度*、纬度*由系统回填，截图中的地址和坐标仅为演示，不得直接抄填。")
done()
figures([(28,"已获取地址，可重新定位并查看地图"),(27,"点击获取经纬度信息后回填经纬度")],15.0)

overview("二、加热设备",[(26,"首次进入加热设备"),(20,"另一烤房的三栏目收起页面")],[
 ("基础信息","工程造价、补贴总额、国家局补贴、产区补贴、年度、施工（供应）单位","灰底信息核对；带*单位填写全称。"),
 ("燃烧机","加热设备类型及选择后出现的字段","按实际热源；现场核实类型、品牌、状态及编号。"),
 ("散热器","状态、最新修复、改造年度","正常／损坏／缺失；只评价，不拍照。")
],14.1,"进入后依次展开“基础信息、燃烧机、散热器”3个栏目。")

page("（一）基础信息　（二）燃烧机")
fields("（一）基础信息",
 "工程造价（元）、补贴总额、国家局补贴、产区补贴、年度、施工（供应）单位*。",
 "灰底展示项按项目档案核对，“暂无内容”不能随意当作0；施工（供应）单位按真实合同或验收资料填写全称。行业投入和地方政府投入应核清来源。")
fields("（二）燃烧机（先选择类型）",
 "加热设备类型*。",
 "点击选择框，按现场实际加热方式选型，选中后再看是否出现新的状态、年度、编号或图片字段。")
figures([(25,"基础信息的六项字段"),(24,"展开燃烧机后选择加热设备类型")],13.6)

page("（二）燃烧机类型　（三）散热器")
fields("（二）燃烧机类型",
 "下拉截图为“内置燃烧机、电热泵、普通燃煤（柴）”。",
 "9.29讨论稿列4类：内置、外置燃烧机、电热泵、燃煤烤房。内置/外置按实物安装方式判断；热泵选电热泵，煤或柴供热选普通燃煤（柴）。外置项只有实时菜单出现时才选。",
 "选型后如出现照片栏，拍清燃烧机整体、类型、品牌、使用状态；内置燃烧机编号以照片留存。")
fields("（三）散热器",
 "散热器状态*；散热器-最新修复、改造年度*。",
 "可正常散热选正常；腐蚀破损、泄漏、故障选损坏；无散热器选缺失。按真实年份填写。此项只评价、不拍照片。")
figures([(23,"加热设备类型实际下拉选项"),(22,"散热器状态及最新修复、改造年度")],12.6)
# no additional line here, to keep the page intact

overview("三、循环自控",[(21,"循环自控设备五个栏目全部收起")],[
 ("自控仪表","状态、最新修复、改造年度、照片","拍清仪表品牌、编号及实际状态。"),
 ("循环风机","状态、最新修复、改造年度","核实风机是否能正常运行；只评价。"),
 ("温湿度传感器","状态、最新修复、改造年度","核实测温、测湿功能；只评价。"),
 ("冷风进风设备","状态、最新修复、改造年度","核实风门、进风设施；只评价。"),
 ("综合评价","评价选择","本模块实际下拉：优良／较好／差。")
],13.0,"进入后依次展开5个栏目。注意循环风机与冷风进风设备在截图中的字段名称存在交叉显示。")

page("（一）自控仪表")
fields("自控仪表",
 "自控仪表状态*；自控仪表-最新修复、改造年度*；自控仪表-正常设备拍摄全景图*。",
 "显示、按键和自动控制功能正常选正常；显示、控制异常或无法使用选损坏；没有仪表选缺失。年度选择本仪表最近修复、改造年份。",
 "正面全景要反映厂家/品牌、编号和当前使用状态；损坏另拍瑕疵图。内置燃烧机、自控仪编号应以照片保存。")
figures([(18,"自控仪表展开后的三项字段"),(17,"点年度字段，滚动年份并确定")],14.0)

page("（二）循环风机　（三）温湿度传感器年度")
fields("（二）循环风机",
 "理论对应循环风机状态、最新修复、改造年度；截图实际显示“冷风门状态*”“冷风门-最新修复、改造年度*”。",
 "现场查看风机、电机、叶片及安装状态，正常运行选正常；故障或不能运转选损坏；缺失选缺失。")
line("注意","图中栏目名为“循环风机”，字段名却为“冷风门”。请先核实系统字段对应关系再填，不可将两类设备的数据混填。",True,12)
fields("（三）温湿度传感器年度","温湿度传感器-最新修复、改造年度*。","每个部件采用自己的实际修复或改造年份。")
figures([(16,"循环风机栏目内显示冷风门字段"),(15,"温湿度传感器年度选择")],13.0)

page("（三）温湿度传感器　（四）冷风进风设备")
fields("（三）温湿度传感器",
 "温湿度传感器状态*；温湿度传感器-最新修复、改造年度*。",
 "探头、线缆正常且温湿度检测有效选正常；异常、失灵选损坏；无探头选缺失。")
fields("（四）冷风进风设备",
 "该栏目截图内显示“循环风机状态*”“循环风机-最新修复、改造年度*”。",
 "核实冷风进风口、风门及控制设施是否正常；待确认字段对应关系后按实填。",
 warning="循环风机、温湿度传感器、冷风进风设备按9.29说明只评价，不拍部件照片；字段名异常应向管理员核实。")
figures([(14,"温湿度传感器及冷风进风设备展开页面")],12.6)

page("（五）循环自控综合评价")
line("填写字段","评价选择*；页面标准说明为“优良、较好、较差”。")
line("实际选项","下拉实际为“优良、较好、差”，应按本页面选项填写。优良预计可用5—8年；较好3—5年；差1—3年。")
line("选择依据","结合自控仪表、循环风机、温湿度传感器、进风设备的实际性能及维修年代判断，故障明显影响烘烤时不能选优良。")
done()
figures([(13,"循环自控综合评价说明"),(12,"实际下拉：优良、较好、差")],14.3)

overview("四、附属设施",[(10,"七栏目总览（其他设施已展开）"),(11,"同一表单中的烤房加盖展开示例")],[
 ("装烟棚","状态、年度、照片","棚架、顶棚和实际使用状态。"),
 ("烤房加盖","状态、年度、照片","屋顶是否加盖、是否漏雨。"),
 ("烟夹","烟夹是否有","是／否。"),
 ("备用电源","备用电源是否有","是／否。"),
 ("电线设施","电线设施是否有","是／否。"),
 ("其他设施","编烟机、操作台、监控","按存在情况多选。"),
 ("综合评价","评价选择","按页面三级标准选择。")
],12.0,"进入后共有7个栏目。总览原图中“其他设施”已展开，未提供所有栏目同时收起的截图。")

page("（一）装烟棚　（二）烤房加盖")
fields("（一）装烟棚",
 "装烟棚状态*；装烟棚-最新修复、改造年度*；装烟棚-正常设备全景照片*。",
 "棚架、顶棚完整可用选正常；破损或漏雨选损坏；无棚选缺失。",
 "拍完整装烟棚，包括支撑及顶棚，损坏另拍细节。")
fields("（二）烤房加盖",
 "烤房加盖状态*；烤房加盖-最新修复、改造年度*；烤房加盖-正常设备拍摄全景图*。",
 "加盖完整、不漏雨选正常；破损漏雨选损坏；没有加盖选缺失。",
 "照片应清楚反映屋顶是否加盖、是否漏雨或破损。")
figures([(9,"装烟棚展开后的状态、年度、照片"),(8,"烤房加盖展开后的三项字段")],12.9)

page("（三）烟夹　（四）备用电源")
fields("（三）烟夹","烟夹是否有*。","下拉“是、否”。现场有配套烟夹选是，无则选否。该项只判断有无，不选设备状态。")
fields("（四）备用电源","备用电源是否有*。","实际有配套备用供电设备选是，无选否。可核对烤房或烤房群配套设施，普通线路不等于备用电源。")
figures([(7,"烟夹是否有：是、否两个选项"),(6,"备用电源是否有及下方栏目")],14.5)

page("（五）电线设施　（六）其他设施")
fields("（五）电线设施","电线设施是否有*。","现场有配套线路、配电设施选是，无选否；如有裸线、老化等安全问题另行报告。")
fields("（六）其他设施",
 "是否存在以下设备*：编烟机、操作台、监控。",
 "逐项查看实物，有哪项勾哪项，可多选；未安装的不要勾选。截图中全勾选和部分勾选只是演示。实际三项都没有但表单要求必填时，先核实录入方式，不得为通过校验虚假勾选。")
figures([(5,"电线设施选是，其他设施尚未勾选"),(4,"编烟机、操作台、监控全部勾选示例")],14.2)

page("（六）其他设施　（七）综合评价")
line("其他设施","图中仅勾选“编烟机、监控”，未勾选“操作台”。实际填写按现场有无选择，不能照抄示例。")
fields("（七）综合评价","评价选择*。","页面标准为优良、较好、较差，按装烟棚、加盖及配套设施整体工况、修复年份和预计可用年限选择（5—8年、3—5年、1—3年）。")
done()
figures([(3,"电线设施、其他设施部分勾选示例"),(2,"其他设施及综合评价展开页面")],14.0)

page("五、补充：基础信息使用状态及采集注意事项")
line("与11.13版衔接","图中“其他—信息完善”属于基础信息指南；填写四大设备前核对烤房使用状态。")
line("使用状态","下拉：在用、闲置、转作他用、不可再用（需完全拆除）、维修中。当前年度正常使用选在用；闲置、改作人用/畜用/杂用、拆除或正在维修按实选。")
line("采集范围","9.29讨论稿提出2009年起建成的烤房纳入采集；每座烤房全景带编号，正常烤房原口径为全景及指定部件共8张，实际照片数按11.17表单及状态要求确定。")
line("弱网与权限","信号差先保存与编号对应的照片、记录地点及状态，按11.13指南的异地上传入口和授权流程补录。技术员提交后不得自行覆盖他人的采集内容。",True,12)
figures([(1,"基础信息完善：关联项目编号和使用状态"),(0,"使用状态的五项实际下拉选项")],12.8)

# Add a compact time-to-figure index to the end, not a duplicate screenshot album.
page("附：原图时间与正文图号索引")
para("压缩包内41张截图的ZIP修改时间均为2026-11-17 20:42:46；无可用EXIF时间。以下按截图文件名的时间从早到晚列出，正文按页面和栏目关系重新排序。图1为单独提供的四按钮总览。",12)
t=doc.add_table(rows=1,cols=4);t.style="Table Grid"
for c,v in zip(t.rows[0].cells,["截图时刻","正文图号","截图时刻","正文图号"]):
    run(c.paragraphs[0],v,True,11);shade(c,"E7E6E6")
def ts(i):
    m=re.search(r"_(\d{6})\.jpg$",imgs[i].name)
    if not m:raise ValueError(imgs[i].name)
    a=m.group(1);return a[:2]+":"+a[2:4]+":"+a[4:]
ordered=sorted(imgs,key=ts)
half=(len(ordered)+1)//2
for j in range(half):
    ids=[ordered[j],ordered[j+half] if j+half<len(ordered) else None]
    vals=[]
    for i in ids:vals.extend([ts(i),str(figure_map[str(i)])] if i is not None else ["",""])
    row=t.add_row();row_keep(row)
    for c,v in zip(row.cells,vals):
        p=c.paragraphs[0];p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.0
        run(p,v,sz=10.5)
set_cols(t,[CW/4]*4)
para("依据：《烤房项目基础信息录入指南11.13》《烤房信息采集说明（9.29）》（讨论稿）、11.17现场表单和截图。9.29与现行表单不一致处，按最新正式业务要求执行。",11)
footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
run(footer,"烤房四大部位信息采集操作指南　"+STAMP+"　｜　第",sz=9)
fld=OxmlElement("w:fldSimple");fld.set(qn("w:instr"),"PAGE")
footer._p.append(fld);run(footer,"页",sz=9)
assert sorted(used)==list(range(41)),(len(used),set(range(41))-set(used))
doc.core_properties.title="烤房四大部位信息采集操作指南（逐项填写·原图完善版）"
doc.save(OUT)
expected={hashlib.sha256(p.read_bytes()).hexdigest() for p in [menu]+list(imgs.values())}
with zipfile.ZipFile(OUT) as z:
    assert z.testzip() is None
    actual={hashlib.sha256(z.read(x)).hexdigest() for x in z.namelist() if x.startswith("word/media/")}
    assert expected==actual,(len(expected),len(actual),len(expected-actual))
    xml=z.read("word/document.xml").decode()
    assert xml.count("<w:color w:val=\"C00000\"")>=10
    assert "<a:srcRect" not in xml
print("DOCX_OK",OUT,OUT.stat().st_size,"screenshots",len(used),"figures",fig,"planned_pages",page_count)
(ROOT/f"_V4图文核验_{STAMP}.json").write_text(json.dumps(dict(planned_pages=page_count,original_images=len(expected),figures=fig,word_bytes=OUT.stat().st_size,figure_map=figure_map,red_marking=True,sha256_original_match=True),ensure_ascii=False,indent=2),encoding="utf-8")
