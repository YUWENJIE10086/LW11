# -*- coding: utf-8 -*-
"""四大部位逐项填写简明图解 V3：原图不裁剪，不改写截图文字。"""
from pathlib import Path
import json,re,zipfile,hashlib
from PIL import Image, ImageOps
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT=Path("低代码1117")
SHOT=ROOT/"解压截图_11172053"
OUT=ROOT/"烤房四大部位信息采集操作指南_逐项填写原图版_11172210_V3.docx"
d0=Document(ROOT/"烤房项目基础信息录入指南11.13_1794919896877.docx")
doc=Document()
for a in ("page_width","page_height","top_margin","bottom_margin","left_margin","right_margin"):
    setattr(doc.sections[0],a,getattr(d0.sections[0],a))
SEC=doc.sections[0]
usable_w=(SEC.page_width-SEC.left_margin-SEC.right_margin)/360000
imgs={}
for x in json.loads((SHOT/"_截图明细_11172053.json").read_text(encoding="utf-8")):
    imgs[int(x["filename"][:3])]=SHOT/x["filename"]
assert len(imgs)==41
normal=doc.styles["Normal"]
normal.font.name="仿宋_GB2312";normal.font.size=Pt(14)
normal._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"),"仿宋_GB2312")
normal.paragraph_format.line_spacing=1.25
normal.paragraph_format.space_after=Pt(5)
def f(p,t,b=False,sz=14,name="仿宋_GB2312"):
    r=p.add_run(t);r.font.name=name;r.font.size=Pt(sz);r.bold=b
    r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"),name)
    return r
def p(t="",indent=True):
    a=doc.add_paragraph()
    if indent:a.paragraph_format.first_line_indent=Pt(28)
    f(a,t)
    return a
def h(t,level=1):
    a=doc.add_paragraph()
    a.paragraph_format.space_before=Pt(10)
    a.paragraph_format.space_after=Pt(7 if level==1 else 4)
    a.paragraph_format.keep_with_next=True
    f(a,t,True,16 if level==1 else 14)
    return a
def detail(label,text):
    a=doc.add_paragraph()
    a.paragraph_format.first_line_indent=Pt(0)
    a.paragraph_format.space_after=Pt(4)
    f(a,label+"：",True);f(a,text)
def table(rows):
    tb=doc.add_table(rows=1,cols=3)
    tb.style="Table Grid"
    tb.alignment=WD_TABLE_ALIGNMENT.CENTER
    tb.autofit=False
    w=[Cm(2.65),Cm(4.6),Cm(7.3)]
    for cell,val,width in zip(tb.rows[0].cells,("展开栏目","看到哪些字段","按什么情况填写"),w):
        cell.width=width;f(cell.paragraphs[0],val,True,11)
    for a,b,c in rows:
        cells=tb.add_row().cells
        for cell,val,width in zip(cells,(a,b,c),w):
            cell.width=width;cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            q=cell.paragraphs[0];q.paragraph_format.space_after=Pt(1)
            f(q,val,sz=10.5)
    doc.add_paragraph()
fig=0
used=[]
def picture(path,title,max_h=16.2,max_w=None):
    global fig
    with Image.open(path) as im:
        width,height=im.size
    max_w=max_w or usable_w-0.5
    set_h=min(max_h,max_w*height/width)
    set_w=set_h*width/height
    pr=doc.add_paragraph()
    pr.alignment=WD_ALIGN_PARAGRAPH.CENTER
    pr.paragraph_format.keep_with_next=True
    pr.paragraph_format.space_after=Pt(1)
    pr.add_run().add_picture(str(path),height=Cm(set_h),width=Cm(set_w))
    fig+=1
    ca=doc.add_paragraph()
    ca.alignment=WD_ALIGN_PARAGRAPH.CENTER
    ca.paragraph_format.space_after=Pt(9)
    f(ca,f"图{fig}　{title}",sz=10)
def shot(i,title,max_h=16.3):
    used.append(i)
    picture(imgs[i],title+"（原始完整截图）",max_h)
def part(title,overview_id,sections,rows):
    doc.add_page_break();h(title,1)
    shot(overview_id,"本入口全部栏目，未逐项展开",16.8)
    detail("栏目总览"," → ".join(sections))
    table(rows)
def item(title,fields,select,photo=None,ids=()):
    h(title,2)
    detail("展开后",fields)
    detail("填写",select)
    if photo:detail("照片",photo)
    for i in ids:shot(i,title,15.6)

top=doc.add_paragraph();top.alignment=WD_ALIGN_PARAGRAPH.CENTER
top.paragraph_format.space_after=Pt(8)
f(top,"烤房四大部位信息采集操作指南",True,22,"方正小标宋简体")
sub=doc.add_paragraph();sub.alignment=WD_ALIGN_PARAGRAPH.CENTER
f(sub,"（逐项填写·原图版）",sz=14)
picture(ROOT/"6e408f052c2c854c097659b8a008a62a.jpg","四个入口总页面",15.7)
p("按《烤房项目基础信息录入指南11.13》找到烤房后，依次进入“主体门窗—加热设备—循环自控—附属设施”。先核对“关联项目编号”，再填写每个栏目。")
detail("共同规则","带*的必填；状态按“正常、损坏、缺失”选；年份按该部件最近实际修复、改造年份选，不照抄截图示例。")
detail("照片规则","正常拍全景1张；损坏拍全景和损坏细节各1张；11.17页面要求缺失拍安装位置全景1张。没有上传栏的部件不另传。")
detail("评价等级","按设备实际工况选：“优良”性能稳定；“较好”基本正常、少量维护；“较差/差”老化明显、需重点维修。页面的维修年份（2023—2026、2018—2022、2017及以前）是参考依据。")

part("一、主体门窗",19,
    ["加热室外立面","大门","装烟室","挂烟梁","综合评价","更新GPS信息"],
    [
      ("加热室外立面","状态、最新修复改造年度、照片","核查墙体、清灰门、风机检修门；按实际状态选。"),
      ("大门","状态、最新修复改造年度、照片","含观察窗；完整可用选正常，有故障选损坏。"),
      ("装烟室","装烟室门状态、最新修复改造年度、照片","检查门、地面、屋墙；裂损漏雨如实选损坏。"),
      ("挂烟梁","材质*","钢制挂烟梁／木制挂烟梁，按实物选。"),
      ("综合评价","评价选择*","按整体状况选优良／较好／较差。"),
      ("更新GPS信息","定位、经度、纬度","在现场获取位置；手动调整500米内、最多两次。")
    ])
item("（一）加热室外立面",
     "外立面状态、最新修复改造年度及照片。",
     "墙体和左右清灰门、风机检修门完好选“正常”；有裂缝、门损坏或无法使用选“损坏”；缺失选“缺失”。年份按实际维修记录选。",
     "拍加热室外立面，能看清清灰门和检修门。",ids=[39])
item("（二）大门",
     "大门状态*、大门最新修复改造年度*、大门全景照片*。",
     "含观察窗。门、窗完整且开关正常选“正常”；破损或不能正常开关选“损坏”；无大门选“缺失”。",
     "拍整扇门及观察窗。",ids=[36])
item("（三）装烟室",
     "装烟室门状态*、装烟室最新修复改造年度*、装烟室全景照片*。",
     "检查门、地面、屋墙；可正常使用选“正常”，破损、明显裂痕或漏雨选“损坏”。",
     "拍门和室内整体情况。",ids=[35])
item("（四）挂烟梁",
     "材质*。",
     "下拉只有“钢制挂烟梁”“木制挂烟梁”，按实际材料选择。",ids=[34])
item("（五）综合评价",
     "评价选择*。",
     "可选“优良、较好、较差”；对照页面的维修年代说明，并结合主体结构是否完好选择。",ids=[33,31])
item("（六）更新GPS信息",
     "定位*、获取经纬度信息、经度*、纬度*。",
     "在烤房现场点“获取当前位置”；如有偏差，可在500米内手工调整，最多两次。确认位置后点“获取经纬度信息”。",
     ids=[28,27])

part("二、加热设备",26,
    ["基础信息","燃烧机","散热器"],
    [
      ("基础信息","工程造价、补贴、年度、施工（供应）单位*","灰底信息核对即可；施工单位按合同填全称。"),
      ("燃烧机","加热设备类型*","实际看到内置燃烧机／电热泵／普通燃煤（柴）等，按热源选。"),
      ("散热器","状态*、最新修复改造年度*","正常／损坏／缺失，按是否完整可用选。")
    ])
item("（一）基础信息",
     "工程造价（元）、补贴总额、国家局补贴、产区补贴、年度、施工（供应）单位*。",
     "前面灰底内容核对原档案；“施工（供应）单位*”输入合同或档案中的单位全称。“暂无内容”不能随便填0。",ids=[25])
item("（二）燃烧机",
     "加热设备类型*。",
     "按现场热源选。截图能看到“内置燃烧机”“电热泵”“普通燃煤（柴）”；9.29说明还列有外置燃烧机，以实时下拉选项为准。",ids=[23])
item("（三）散热器",
     "散热器状态*、散热器最新修复改造年度*。",
     "能正常散热、无明显损坏选“正常”；损坏或不能正常使用选“损坏”；没有设备选“缺失”。此项按旧说明只评价，无照片栏。",ids=[22])

part("三、循环自控",21,
    ["自控仪表","循环风机","温湿度传感器","冷风进风设备","综合评价"],
    [
      ("自控仪表","状态、最新修复改造年度、照片","显示和控制正常选正常，失灵选损坏。"),
      ("循环风机","状态、最新修复改造年度","可正常运转选正常，故障选损坏。"),
      ("温湿度传感器","状态、最新修复改造年度","测温湿正常选正常，失灵选损坏。"),
      ("冷风进风设备","状态、最新修复改造年度","按冷风进风实际工况填写，注意截图字段名异常。"),
      ("综合评价","评价选择*","本模块截图下拉为优良／较好／差。")
    ])
item("（一）自控仪表",
     "自控仪表状态*、最新修复改造年度*、自控仪表全景照片*。",
     "显示、按键和控制正常选“正常”；失灵选“损坏”；未安装选“缺失”。",
     "正面全景拍清仪表和编号。",ids=[18])
item("（二）循环风机",
     "循环风机状态*、最新修复改造年度*。",
     "运转正常选“正常”；电机、叶片故障或无法运转选“损坏”；无风机选“缺失”。9.29说明此项只评价。",ids=[16])
item("（三）温湿度传感器",
     "温湿度传感器状态*、最新修复改造年度*。",
     "探头能正常测温湿选“正常”；失灵或损坏选“损坏”；无探头选“缺失”。只评价、不拍照片。",ids=[14])
item("（四）冷风进风设备",
     "展开后截图错误显示“循环风机状态”“循环风机最新修复改造年度”。",
     "先核实表单字段关联，不要把循环风机状态当成冷风进风设备状态；核实后按进风口、风阀是否正常选择。")
item("（五）综合评价",
     "评价选择*。",
     "下拉显示“优良、较好、差”（本模块为“差”字），按自控、风机、传感器整体状况选。",ids=[12])

part("四、附属设施",10,
    ["装烟棚","烤房加盖","烟夹","备用电源","电线设施","其他设施","综合评价"],
    [
      ("装烟棚","状态、最新修复改造年度、照片","完好可用正常；破损选损坏。"),
      ("烤房加盖","状态、最新修复改造年度、照片","加盖完整正常；破损漏雨选损坏。"),
      ("烟夹","烟夹是否有*","有选“是”，无选“否”。"),
      ("备用电源","备用电源是否有*","现场有选“是”，无选“否”。"),
      ("电线设施","电线设施是否有*","实际有选“是”，无选“否”。"),
      ("其他设施","编烟机、操作台、监控","逐项勾选实际存在的设备。"),
      ("综合评价","评价选择*","选优良／较好／较差。")
    ])
item("（一）装烟棚",
     "装烟棚状态*、最新修复改造年度*、装烟棚全景照片*。",
     "棚架、顶棚完好可用选“正常”；破损漏雨选“损坏”；没有装烟棚选“缺失”。",
     "拍整个装烟棚。",ids=[9])
item("（二）烤房加盖",
     "烤房加盖状态*、最新修复改造年度*、加盖全景照片*。",
     "屋顶加盖完好选“正常”；破损或漏雨选“损坏”；未加盖选“缺失”。",
     "拍清屋顶加盖整体情况。",ids=[8])
item("（三）烟夹",
     "烟夹是否有*。",
     "下拉“是、否”；现场有烟夹选“是”，没有选“否”。",ids=[7])
item("（四）备用电源",
     "备用电源是否有*。",
     "有备用电源选“是”，没有选“否”；不能把普通供电线路当备用电源。",ids=[6])
item("（五）电线设施",
     "电线设施是否有*。",
     "实际有配套电线设施选“是”，没有选“否”；有故障另行报告。",ids=[5])
item("（六）其他设施",
     "是否存在以下设备*：编烟机、操作台、监控。",
     "哪个设备实际存在就勾选哪个，可多选；不存在的不勾选。",ids=[2])
item("（七）综合评价",
     "评价选择*。",
     "根据棚、加盖、配套设施整体情况选“优良、较好、较差”。",ids=[3])

doc.add_page_break()
h("五、共用填写方法")
detail("年份","点“最新修复、改造年度*”，上下滑动到真实年份，再点“确定”。不是所有部件都统一选2026年。")
detail("拍照","点“拍照上传”，选“系统相机”或“浏览器相机”，拍摄后确认页面已显示照片。")
shot(38,"拍照上传后选择相机",16)
detail("提交","每个入口填写结束，核对带*字段和照片后点“提交”，成功后关闭，再进下一个按钮。")
detail("提醒","9.29说明曾写“缺失无需拍照”，11.17表单顶端已写“缺失设备拍全景1张”，以当前页面和最新通知为准。")
detail("编号","截图有00138、00144、00147等不同烤房，仅用于演示；实际每一页都要核对“关联项目编号*”。")

# Include complete unmodified bytes as media. No crops, transformations or recompression.
all_hashes={hashlib.sha256((ROOT/"6e408f052c2c854c097659b8a008a62a.jpg").read_bytes()).hexdigest()}
all_hashes|={hashlib.sha256(imgs[i].read_bytes()).hexdigest() for i in set(used)}
doc.save(OUT)
with zipfile.ZipFile(OUT) as z:
    assert z.testzip() is None
    media=[z.read(x) for x in z.namelist() if x.startswith("word/media/") and not x.endswith("/")]
    actual={hashlib.sha256(x).hexdigest() for x in media}
    assert all_hashes<=actual, ("Original JPEG bytes changed",len(all_hashes-actual))
print("WORD_OK",str(OUT),"figures",fig,"original_files",len(all_hashes),"size",OUT.stat().st_size)
