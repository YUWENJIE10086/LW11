# -*- coding: utf-8 -*-
"""四大部位四大部位信息采集操作指南：原图不裁剪，不改写截图文字。"""
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
OUT=ROOT/"烤房四大部位信息采集操作指南_逐项填写原图版_11181455_V5.docx"
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
def required_runs(p,text,sz=14):
    # 星号字段整体红色，保留字段原有标点及“修复、改造年度”中的顿号。
    for v in re.split(r"(、(?!改造年度)|；|。)", text):
        if not v: continue
        rr=f(p,v,sz=sz)
        if "*" in v:
            rr.font.color.rgb=RGBColor(192,0,0);rr.bold=True
def detail(label,text):
    a=doc.add_paragraph()
    a.paragraph_format.first_line_indent=Pt(0)
    a.paragraph_format.space_after=Pt(4)
    f(a,label+"：",True)
    if label=="需填写内容":
        required_runs(a,text);return
    # 强调现场填写的关键要求。
    marks=r"(带\*的字段必须填写|必填字段|正常拍全景1张|全景和损坏细节各1张|缺失拍安装位置全景1张|带\*的必填|500米内|最多两次|只评价、不拍照片|只评价|不得|不能|系统相机|关联项目编号|2023—2026年|2018—2022年|2017年及以前|5—8年|3—5年|1—3年|先核实表单字段关联|编号以照片保存)"
    for v in re.split(marks,text):
        if not v:continue
        rr=f(a,v)
        if re.fullmatch(marks,v):
            rr.font.color.rgb=RGBColor(192,0,0);rr.bold=True
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
            required_runs(q,val,10.5)
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
    picture(imgs[i],title,max_h)
def pair(ids,titles,height=12.5):
    global fig
    assert len(ids)==len(titles)==2
    tb=doc.add_table(rows=1,cols=2);tb.autofit=False;tb.alignment=WD_TABLE_ALIGNMENT.CENTER
    row=tb.rows[0]
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    for cell,i,title in zip(row.cells,ids,titles):
        used.append(i);cell.width=Cm(usable_w/2)
        with Image.open(imgs[i]) as im: w,h=im.size
        hh=min(height,(usable_w/2-.2)*h/w)
        pp=cell.paragraphs[0];pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.space_after=Pt(2)
        pp.add_run().add_picture(str(imgs[i]),height=Cm(hh),width=Cm(hh*w/h))
        fig+=1
        cc=cell.add_paragraph();cc.alignment=WD_ALIGN_PARAGRAPH.CENTER
        f(cc,f"图{fig}　{title}",sz=10)
    doc.add_paragraph().paragraph_format.space_after=Pt(1)
def part(title,overview_id,sections,rows):
    doc.add_page_break();h(title,1)
    overview_extra={19:40,26:20}
    if overview_id in overview_extra:
        pair([overview_id,overview_extra[overview_id]],["采集项目完整页面","采集页面首屏"],13.2)
    else:
        shot(overview_id,"采集项目页面",14.2)
    detail("采集项目","、".join(sections)+"。")
    doc.add_page_break()
    h("采集项目及填写要求",2)
    table(rows)
CAPTIONS={
0:"使用状态选项",1:"基本情况及使用状态",2:"其他设施与综合评价",
3:"电线设施及其他设施选择结果",4:"其他设施全部勾选",5:"电线设施及其他设施选项",
6:"备用电源是否有",7:"烟夹是否有选项",8:"烤房加盖状态、年度及照片",9:"装烟棚状态、年度及照片",
10:"附属设施采集项目",12:"循环自控评价选项",13:"循环自控评价标准",
14:"冷风进风设备状态及年度",15:"温湿度传感器修复改造年度",16:"循环风机栏目及年度选择",
17:"自控仪表修复改造年度",18:"自控仪表状态、年度及照片",19:"主体门窗采集项目",20:"加热设备采集项目",
21:"循环自控采集项目",22:"散热器状态及年度",23:"加热设备类型选项",24:"燃烧机及散热器填写页面",
25:"加热设备基础信息",26:"加热设备采集项目",27:"获取经纬度信息及坐标",28:"定位结果及地图调整",
29:"获取当前位置",30:"定位及经纬度字段",31:"主体门窗评价选项",32:"综合评价下拉选项",33:"综合评价标准",
34:"挂烟梁材质选项",35:"装烟室状态、年度、照片及挂烟梁材质",36:"大门及装烟室填写页面",37:"手机相机拍摄页面",
38:"选择系统相机",39:"加热室外立面及年份选择",40:"主体门窗采集页面"
}
def item(title,fields,select,photo=None,ids=()):
    h(title,2)
    detail("需填写内容",fields)
    detail("如何填写",select)
    if photo:detail("图片注意事项",photo)
    # 同一栏目连续的两幅原图并排，标题直接标示展示内容。
    for j in range(0,len(ids),2):
        a=ids[j:j+2]
        if len(a)==2:pair(a,[CAPTIONS[i] for i in a],12.5)
        else:shot(a[0],CAPTIONS[a[0]],13.0)

top=doc.add_paragraph();top.alignment=WD_ALIGN_PARAGRAPH.CENTER
top.paragraph_format.space_after=Pt(8)
f(top,"烤房四大部位信息采集操作指南",True,22,"方正小标宋简体")
picture(ROOT/"6e408f052c2c854c097659b8a008a62a.jpg","四个入口总页面",11.0)
p("搜索烤房编码，找到对应烤房后，依次点击页面下方“主体门窗、加热设备、循环自控、附属设施”四个按钮。进入每个采集页面后，先核对关联项目编号，再逐项填写。")
detail("共同规则","带*的字段必须填写；状态按“正常、损坏、缺失”选；年份按该部件最近实际修复、改造年份选；从未维修改造的，核对原建设或安装年度，不能凭空填写。")
detail("照片规则","正常拍全景1张；损坏拍全景和损坏细节各1张；缺失拍安装位置全景1张。没有上传栏的部件不另传。")
detail("评价等级","优良：2023—2026年建设维修，性能稳定，预计可使用5—8年；较好：2018—2022年建设维修，性能良好，仅需常规管护或少量部件更换，预计可使用3—5年；较差：2017年及以前建设维修，设备老化明显，需重点投入管护或更换部件，预计可使用1—3年。")

doc.add_page_break()
h("现场采集及照片上传要求")
detail("必填字段","页面中带*的必填字段必须填写完整；同一烤房各入口的关联项目编号必须一致。")
detail("修复改造年度","点击“最新修复、改造年度”，上下滑动选择实际年份，再点击“确定”。未维修改造的，核对原建设或安装年度。不得将所有部件统一填作2026年。")
detail("照片采集","正常拍全景1张；损坏拍全景和损坏细节各1张；缺失拍安装位置全景1张。整房全景须包含烤房编号；损坏或改作他用可补充现场照片1—2张。散热器、循环风机、温湿度传感器、冷风进风设备只评价、不拍照片。")
detail("拍照上传","点击“拍照上传”→选择“系统相机”→点击“确定”→打开手机相机→对准部件拍摄→返回表单确认照片上传成功。系统相机无法调用时，可尝试浏览器相机。")
pair([38,37],["选择系统相机并点击确定","打开手机相机并拍照"],12.8)
detail("图片注意事项","照片应清晰、真实，能辨认部件及安装位置。全景拍设备整体，瑕疵图对准损坏部位；内置燃烧机及自控仪编号以照片保存。不得使用其他烤房的照片。")
detail("提交要求","每个入口填写完毕，核对必填字段及照片，点击“提交”；信号较弱时，先记录烤房编号和设备信息、拍好照片，待网络恢复后按异地上传流程补录。")

part("一、主体门窗",19,
    ["加热室外立面","大门","装烟室","挂烟梁","综合评价","更新GPS信息"],
    [
      ("加热室外立面","状态*、最新修复、改造年度*、照片*","核查墙体、清灰门、风机检修门；按实际状态选。"),
      ("大门","状态*、最新修复、改造年度*、照片*","含观察窗；完整可用选正常，有故障选损坏。"),
      ("装烟室","装烟室门状态*、最新修复、改造年度*、照片*","检查门、地面、屋墙；裂损漏雨如实选损坏。"),
      ("挂烟梁","材质*","钢制挂烟梁／木制挂烟梁，按实物选。"),
      ("综合评价","评价选择*","按整体状况选优良／较好／较差。"),
      ("更新GPS信息","定位*、经度*、纬度*","在现场获取位置；手动调整500米内、最多两次。")
    ])
item("（一）加热室外立面",
     "加热室外立面-状态*、加热室外立面-最新修复、改造年度*、加热室外立面-正常设备全景照片*。",
     "墙体和左右清灰门、风机检修门完好选“正常”；有裂缝、门损坏或无法使用选“损坏”；缺失选“缺失”。年份按实际维修记录选。",
     "拍加热室外立面，能看清左右清灰门和风机检修门。正常1张；损坏另拍瑕疵细节1张。点击“拍照上传”，选择系统相机并确定，拍摄后核对照片。",ids=[39])
item("（二）大门",
     "大门-状态*、大门-最新修复、改造年度*、大门-正常设备全景照片*。",
     "含观察窗。门、窗完整且开关正常选“正常”；破损或不能正常开关选“损坏”；无大门选“缺失”。",
     "拍整扇门及观察窗。",ids=[36])
item("（三）装烟室",
     "装烟室门-状态*、装烟室-最新修复、改造年度*、装烟室-正常设备全景照片*。",
     "检查门、地面、屋墙；可正常使用选“正常”，破损、明显裂痕或漏雨选“损坏”。",
     "拍门和室内整体情况。",ids=[35])
item("（四）挂烟梁",
     "材质*。",
     "下拉只有“钢制挂烟梁”“木制挂烟梁”，按实际材料选择。",ids=[34])
item("（五）综合评价",
     "评价选择*。",
     "可选“优良、较好、较差”。优良：2023—2026年建设维修、性能稳定、预计可用5—8年；较好：2018—2022年建设维修、性能良好、常规管护或少量更换、预计3—5年；较差：2017年及以前建设维修、老化明显、需重点管护或更换、预计1—3年。",ids=[33,31,32])
item("（六）更新GPS信息",
     "定位*、获取经纬度信息、经度*、纬度*。",
     "在烤房现场点“获取当前位置”；如有偏差，可在500米内手工调整，最多两次。确认位置后点“获取经纬度信息”，查看经度、纬度是否回填。行业烤房原基础信息表GPS为核对依据；政府烤房以现场采集为准。",
     ids=[30,29,28,27])

part("二、加热设备",26,
    ["基础信息","燃烧机","散热器"],
    [
      ("基础信息","工程造价、补贴、年度、施工（供应）单位*","灰底信息核对即可；施工单位按合同填全称。"),
      ("燃烧机","加热设备类型*","可选内置燃烧机／电热泵／普通燃煤（柴）等，按热源选。"),
      ("散热器","状态*、最新修复改造年度*","正常／损坏／缺失，按是否完整可用选。")
    ])
item("（一）基础信息",
     "工程造价（元）、补贴总额、国家局补贴、产区补贴、年度、施工（供应）单位*。",
     "前面灰底内容核对原档案；“施工（供应）单位*”输入合同或档案中的单位全称。“暂无内容”不能随便填0。",ids=[25])
item("（二）燃烧机",
     "加热设备类型*。",
     "选择实际热源：内置燃烧机（燃烧机内置）、电热泵（热泵供热）、普通燃煤（柴）（传统炉具）。如实际菜单提供“外置燃烧机”，按现场设备选择。选型后如出现品牌、状态、年度、照片字段，按实际填写；燃烧机类型、品牌和状态须真实，内置燃烧机编号以照片保存。",ids=[24,23])
item("（三）散热器",
     "散热器状态*、散热器最新修复改造年度*。",
     "能正常散热、无明显损坏选“正常”；损坏或不能正常使用选“损坏”；没有设备选“缺失”。此项只评价、不拍照片。",ids=[22])

part("三、循环自控",21,
    ["自控仪表","循环风机","温湿度传感器","冷风进风设备","综合评价"],
    [
      ("自控仪表","状态*、最新修复、改造年度*、照片*","显示和控制正常选正常，失灵选损坏。"),
      ("循环风机","状态*、最新修复、改造年度*","可正常运转选正常，故障选损坏。"),
      ("温湿度传感器","状态*、最新修复、改造年度*","测温湿正常选正常，失灵选损坏。"),
      ("冷风进风设备","状态*、最新修复、改造年度*","按冷风进风实际工况填写，注意截图字段名异常。"),
      ("综合评价","评价选择*","本模块下拉为优良／较好／差。")
    ])
item("（一）自控仪表",
     "自控仪表状态*、自控仪表-最新修复、改造年度*、自控仪表-正常设备拍摄全景图*。",
     "显示、按键和控制正常选“正常”；失灵选“损坏”；未安装选“缺失”。",
     "正面全景拍清仪表品牌、显示屏和编号；自控仪编号以照片保存。",ids=[18,17])
item("（二）循环风机",
     "页面“循环风机”栏目下显示“冷风门状态*”“冷风门-最新修复、改造年度*”。",
     "按实际循环风机运行情况判断正常、损坏、缺失，并选择最新修复年度。先核实表单字段关联，勿将冷风门与风机混填。此项只评价、不拍照片。",ids=[16])
item("（三）温湿度传感器",
     "温湿度传感器状态*、温湿度传感器-最新修复、改造年度*。",
     "探头能正常测温湿选“正常”；失灵或损坏选“损坏”；无探头选“缺失”。只评价、不拍照片。",ids=[15])
item("（四）冷风进风设备",
     "页面“冷风进风设备”栏目下显示“循环风机状态*”“循环风机-最新修复、改造年度*”。",
     "先核实表单字段关联，勿将循环风机状态当成冷风进风设备状态；核实后按进风口、风阀实际情况选正常、损坏或缺失。只评价、不拍照片。",ids=[14])
item("（五）综合评价",
     "评价选择*。",
     "下拉选项为“优良、较好、差”，按自控、风机、传感器整体状况选。",ids=[13,12])

part("四、附属设施",10,
    ["装烟棚","烤房加盖","烟夹","备用电源","电线设施","其他设施","综合评价"],
    [
      ("装烟棚","状态*、最新修复、改造年度*、照片*","完好可用正常；破损选损坏。"),
      ("烤房加盖","状态*、最新修复、改造年度*、照片*","加盖完整正常；破损漏雨选损坏。"),
      ("烟夹","烟夹是否有*","有选“是”，无选“否”。"),
      ("备用电源","备用电源是否有*","现场有选“是”，无选“否”。"),
      ("电线设施","电线设施是否有*","实际有选“是”，无选“否”。"),
      ("其他设施","是否存在以下设备*（编烟机、操作台、监控）","逐项勾选实际存在的设备。"),
      ("综合评价","评价选择*","选优良／较好／较差。")
    ])
item("（一）装烟棚",
     "装烟棚状态*、装烟棚-最新修复、改造年度*、装烟棚-正常设备全景照片*。",
     "棚架、顶棚完好可用选“正常”；破损漏雨选“损坏”；没有装烟棚选“缺失”。",
     "拍整个装烟棚。",ids=[9])
item("（二）烤房加盖",
     "烤房加盖状态*、烤房加盖-最新修复、改造年度*、烤房加盖-正常设备拍摄全景图*。",
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
     "编烟机、操作台、监控各自核实，有就勾选、无不勾选，可多选。未勾选表示未配置，黄色勾选表示已配置；不能套用图中选择结果。",ids=[4,3])
item("（七）综合评价",
     "评价选择*。",
     "根据棚、加盖、配套设施整体情况选“优良、较好、较差”。",ids=[2])

doc.add_page_break()
h("五、使用状态核对")
detail("需填写内容","使用状态*。")
detail("如何填写","点击“其他—信息完善”，核对“关联项目编号*”，展开“基本情况”并选择使用状态。实际用于烘烤选“在用”；暂未使用选“闲置”；已经改作其他用途选“转作他用”；完全失去使用能力且需拆除选“不可再用（需完全拆除）”；正在维修选“维修中”。")
pair([1,0],["基本情况及使用状态","使用状态选项"],12.8)
assert set(used)==set(imgs)-{11}, ("Image coverage",set(imgs)-set(used), set(used)&{11})
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
