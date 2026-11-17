# -*- coding: utf-8 -*-
"""按2026-11-17现场截图编写的烤房四大部位逐项填写指南（独立续篇）。"""
from pathlib import Path
import re,json,zipfile,math
from PIL import Image,ImageOps,ImageDraw,ImageFont
from docx import Document
from docx.shared import Cm,Pt,RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path("低代码1117"); IMG=ROOT/"解压截图_11172053"
OUTPUT=ROOT/"烤房四大部位信息采集_逐项填写图解指南_11172125_V2.docx"
CROPS=ROOT/"_guide_v2_crops"
CROPS.mkdir(exist_ok=True)
manifest=json.loads((IMG/"_截图明细_11172053.json").read_text(encoding="utf-8"))
PIC={int(x["filename"][:3]):IMG/x["filename"] for x in manifest}
assert len(PIC)==41
ref=Document(ROOT/"烤房项目基础信息录入指南11.13_1794919896877.docx")
doc=Document()
s=doc.sections[0]; rs=ref.sections[0]
for a in ("page_width","page_height","top_margin","bottom_margin","left_margin","right_margin"):
    setattr(s,a,getattr(rs,a))
s.header_distance=Cm(1.2);s.footer_distance=Cm(1.2)
normal=doc.styles["Normal"]
normal.font.name="仿宋_GB2312";normal.font.size=Pt(14)
normal._element.rPr.rFonts.set(qn("w:eastAsia"),"仿宋_GB2312")
normal.paragraph_format.line_spacing=1.25
normal.paragraph_format.space_after=Pt(6)
def run(p,t,b=False,size=14,font="仿宋_GB2312",c=None):
    r=p.add_run(t);r.font.name=font;r.font.size=Pt(size);r.bold=b
    r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"),font)
    if c:r.font.color.rgb=RGBColor(*c)
    return r
def para(txt="",indent=True,size=14,bold=False):
    p=doc.add_paragraph();p.paragraph_format.space_after=Pt(5)
    if indent:p.paragraph_format.first_line_indent=Pt(28)
    run(p,txt,bold,size)
    return p
def label(title,text):
    p=doc.add_paragraph();p.paragraph_format.left_indent=Pt(0)
    p.paragraph_format.first_line_indent=Pt(0)
    p.paragraph_format.space_after=Pt(4)
    run(p,title+"：",True);run(p,text)
    return p
def h(txt,level=1):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(13 if level==1 else 7)
    p.paragraph_format.space_after=Pt(6)
    p.paragraph_format.keep_with_next=True
    run(p,txt,b=True,size=16 if level==1 else 14)
    return p
def meta(msg):
    p=doc.add_paragraph()
    p.paragraph_format.space_after=Pt(6)
    run(p,msg,size=11,c=(90,90,90))
    return p
def font_path():
    for x in ["/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
              "/usr/share/fonts/truetype/arphic/ukai.ttc",
              "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]:
        if Path(x).exists():return x
    return None
FONT=font_path()
def crop(parts,width_limit=1260):
    """parts list (index,start_frac,end_frac); combine narrow parts horizontally."""
    src_id=".".join(f"{i}-{int(y0*100)}-{int(y1*100)}" for i,y0,y1 in parts)
    target=CROPS/(src_id+".jpg")
    if target.exists():return target
    pieces=[]
    for i,a,b in parts:
        im=Image.open(PIC[i]);im=ImageOps.exif_transpose(im).convert("RGB")
        W,H=im.size
        y0=max(0,min(H-50,int(H*a)));y1=min(H,max(y0+50,int(H*b)))
        im=im.crop((int(W*.028),y0,int(W*.972),y1))
        pieces.append(im)
    if len(pieces)==1:
        result=pieces[0]
    else:
        # equal height panels, making tall scroll views legible as 2-column composite
        H=max(p.height for p in pieces)
        W=max(p.width for p in pieces)
        result=Image.new("RGB",(W*len(pieces)+24*(len(pieces)-1),H),(255,255,255))
        for k,im in enumerate(pieces):
            result.paste(im,(k*(W+24),0))
    if result.width>2600:
        sc=2600/result.width
        result=result.resize((int(result.width*sc),int(result.height*sc)),Image.Resampling.LANCZOS)
    result.save(target,quality=87,optimize=True)
    return target
fig=0
def photo(parts,caption,max_h=11.6):
    global fig
    path=crop(parts)
    with Image.open(path) as im:
        W,H=im.size
    width=min(14.35,max_h*W/H)
    # prevent rendering too tiny; readable crops are preferred over full 2800px screenshots
    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after=Pt(1)
    p.paragraph_format.keep_with_next=True
    p.add_run().add_picture(str(path),width=Cm(width))
    fig+=1
    c=doc.add_paragraph();c.alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_after=Pt(8)
    run(c,f"图{fig}　{caption}",size=10)
def module(title,sequence,overview_img,intro):
    doc.add_page_break()
    h(title)
    label("点开后首先看到", "、".join(sequence)+"。")
    photo(overview_img,title+"——未展开栏目（先认栏目名称，再逐项展开）",max_h=11.6)
    para(intro)
def item(title,fields,choice,rule,pics=None,comment=None):
    h(title,2)
    label("展开后看到",fields)
    label("怎么选、怎么填",choice)
    if rule:label("照片或其他",rule)
    if pics:photo(pics,title+"对应的页面字段或选项",max_h=8.9)
    if comment:meta(comment)

top=doc.add_paragraph()
top.alignment=WD_ALIGN_PARAGRAPH.CENTER
top.paragraph_format.space_after=Pt(9)
run(top,"烤房四大部位信息采集逐项填写指南",True,22,"方正小标宋简体")
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
run(p,"（主体门窗、加热设备、循环自控、附属设施）",size=14)
meta("本指南接续《烤房项目基础信息录入指南11.13》，从已搜索到烤房并核对项目编号后开始。")
photo([(0,.0,1.0)],"截图序号000并非此图：正式首页采用仓库单独提供的四按钮总览截图。") if False else None
# The user-supplied menu image has a unique filename and is deliberately placed first.
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.keep_with_next=True;p.paragraph_format.space_after=Pt(1)
p.add_run().add_picture(str(ROOT/"6e408f052c2c854c097659b8a008a62a.jpg"),height=Cm(16.9))
fig+=1
c=doc.add_paragraph();c.alignment=WD_ALIGN_PARAGRAPH.CENTER
run(c,"图1　四个按钮的总页面：主体门窗／加热设备／循环自控／附属设施",size=10)
label("填写顺序","先按11.13版指南完成“其他—信息完善”；回到这个页面，依次点击底部“主体门窗”“加热设备”“循环自控”“附属设施”。每次进去先核对“关联项目编号”，填完本入口再点击“提交”，关闭后进入下一个入口。")
label("所有页面共用","带红色“*”是必填；灰色标题栏右边的小箭头用于展开或收起；点“最新修复、改造年度”会弹出年份滚轮，选好点“确定”；选择完成后检查字段确实显示所选内容。")

# Source standard after newer version inspection
h("填写前先认清3类常用选项")
label("部件状态","采集说明统一为“正常、损坏、缺失”。能正常使用、外观及功能无明显异常，选“正常”；部件仍在但破损、故障或不能正常使用，选“损坏”；原本应有部件而现场没有，选“缺失”。不要因为某项是截图示例中的绿色“正常”就照选。")
label("修复、改造年度","按本部件最近一次实际修复或更新的年份选择，参考维修记录、验收资料和现场核对结果。截图滚轮中出现“2026”只是示例。没有修复记录或者建设年份与维修年份关系不清时，先核实，不要直接填当年。")
label("拍摄照片","最新表单顶端明确写：正常拍全景1张；损坏拍全景1张和瑕疵细节1张；缺失拍安装位置全景1张。与9.29采集说明中“缺失不拍照”不一致，新稿按11.17表单来讲。散热器、循环风机、温湿度传感器、冷风进风设备等若实际页只有状态与年份、没有上传框，就不自行增加照片。")
photo([(38,.30,.63),(37,.32,.88)],"“拍照上传”选择系统相机/浏览器相机，进入相机后拍照",max_h=6.4)

module("一、主体门窗",
       ["加热室外立面","大门","装烟室","挂烟梁","综合评价","更新GPS信息"],
       [(19,.05,.54),(19,.47,.97)],
       "以下按页面出现的先后顺序展开。对应图片中有些项目编号是23-420527KF00144，有些为00147，都是演示画面，实际填写时只能使用当前烤房自己的编号。")
item("（一）加热室外立面",
     "“加热室外立面-状态*”“加热室外立面-最新修复、改造年度*”以及设备照片上传栏。页面说明包含左右清灰门和风机检修门。",
     "检查加热室墙体、左右清灰门、风机检修门是否完整，门是否能正常开启关闭；完整且正常使用选“正常”，出现破损、变形、开关失灵选“损坏”，整项不存在选“缺失”。年份选该处最近一次修复或改造的实际年份。",
     "正常拍下含左右清灰门和风机检修门的加热室外立面全景；损坏时再拍清晰的裂缝、变形或损坏部位；缺失按现行表单拍摄原位置。",
     [(39,.22,.74)],"截图是打开年份选择时的画面，背景稍暗，但可对应加热室外立面的状态和年度字段。")
item("（二）大门",
     "“大门-状态*”“大门-最新修复、改造年度*”“大门-正常设备全景照片*”。页面备注“含观察窗”。",
     "门扇、合页、门锁、观察窗、密封状况正常且可用选“正常”；门扇破损、观察窗损坏、无法正常开关或漏风明显选“损坏”；整扇门缺失选“缺失”。维修年份以实际维修或更换记录为准。",
     "照片应能完整看见大门和观察窗。损坏时另拍破损、漏风、松脱等细节。",
     [(36,.17,.76)])
item("（三）装烟室",
     "“装烟室门-状态*”“装烟室-最新修复、改造年度*”“装烟室-正常设备全景照片*”。页面提示检查地面、屋墙等。",
     "检查装烟室门、地面、墙面和屋顶。无破损且能正常装烟使用选“正常”；墙面开裂、屋顶漏雨、地面严重损坏或门体故障选“损坏”；有应配但不存在的主体部件按实际选“缺失”。",
     "拍摄装烟室全景，能够说明地面、墙体和装烟室门的实际状况；发现破损再补拍问题点。",
     [(35,.14,.77)])
item("（四）挂烟梁",
     "该截图展开后只有一个必填字段：“材质*”。",
     "点击“材质*”，下拉显示“钢制挂烟梁”“木制挂烟梁”两项。现场看挂烟梁实际材料：钢质选“钢制挂烟梁”，木质选“木制挂烟梁”，不得按旧台账猜测。",
     "挂烟梁如果有断裂、缺失或无法确认材质，不能仅为通过校验随意选择；先核实并按管理人员要求处理。",
     [(34,.20,.73)])
item("（五）综合评价",
     "“评价选择*”，上方有“优良、较好、较差”三档文字说明。",
     "页面写明：“优良”为2023—2026年建设维修、性能稳定、预计可用5—8年；“较好”为2018—2022年建设维修、性能良好，仅需常规管护或少量部件更换、预计可用3—5年；“较差”为2017年及以前建设维修、老化明显、需重点管护或更换、预计可用1—3年。按实际设备性能和维修情况综合选，不只按年份机械判断。",
     "点“评价选择”，弹出“优良、较好、较差”；选择后确认已回填。",
     [(33,.16,.65),(31,.44,.70)])
item("（六）更新GPS信息",
     "“定位*”“获取当前位置”“重新定位”“地址转换-获取经纬度”“经度*”“纬度*”。页面红字提醒：500米内手工调整，只能调整两次。",
     "站在对应烤房现场开启定位权限，点击“获取当前位置”。地图上的点要与烤房实际位置一致；如果需要修改，仅在允许范围内调整，不要随意移动两次。然后点击“获取经纬度信息”，确认经度和纬度已经显示、地点准确。",
     "9.29说明要求行业烤房经纬度优先与基础信息表核对；外部烤房按实地采集。截图上的坐标是演示数值，不要照搬。",
     [(30,.26,.83),(27,.14,.75)])
label("本入口完成后","回看6个分组，带“*”的必填字段均有值，照片已成功显示，再点页面下方“提交”。")

module("二、加热设备",
       ["基础信息","燃烧机","散热器"],
       [(26,.10,.87)],
       "先核对关联项目编号，然后从“基础信息”开始向下展开。这里重点采集实际加热方式和散热器情况，不要把“燃烧机”和“散热器”当作同一个设备。")
item("（一）基础信息",
     "“工程造价（元）”“补贴总额”“国家局补贴”“产区补贴”“年度”，以及“施工（供应）单位*”。",
     "截图中的造价、补贴、年度显示为灰底“暂无内容”，这部分看起来是系统已有信息的展示，不应随意编造金额。下方白色输入框“施工（供应）单位*”需要输入对应施工单位或设备供应商的准确全称，以项目合同、验收资料或档案为准。",
     "如果前面系统带出的金额或年度有误，应按管理流程核实基础档案，不要随意写一个数字凑齐。",
     [(25,.10,.51),(25,.48,.92)])
item("（二）燃烧机",
     "先看到“加热设备类型*”下拉框。",
     "按现场实际热源选：截图能看到“内置燃烧机”“电热泵”“普通燃煤（柴）”；9.29采集说明还提到“外置燃烧机”，该项在所拍下拉画面没有显示，只有实际菜单出现时才能选择。内置燃烧机选内置，热泵供热选电热泵，传统燃煤/柴火选普通燃煤（柴）。",
     "选型后若系统继续展开燃烧机状态、修复年份或照片栏，就按实际显示继续填写；应核对设备品牌、编号、当前使用状态，内置燃烧机的编号要拍清楚。",
     [(24,.22,.58),(23,.41,.79)])
item("（三）散热器",
     "“散热器状态*”“散热器-最新修复、改造年度*”。截图没有出现散热器照片栏。",
     "散热器无明显腐蚀、破损、泄漏且可正常散热选“正常”；存在漏风漏气、严重锈蚀、破损或不能正常工作的选“损坏”；现场应有却没有散热器选“缺失”。年度选择实际维修或更新年份。",
     "9.29采集说明明确散热器只评价，不要求单独拍部件照片；若现在的表单没有上传框，不需要另行上传。",
     [(22,.43,.86)])
label("本入口完成后","检查热源类型、散热器状态、修复年度及施工单位填写正确，再点击“提交”。")

module("三、循环自控",
       ["自控仪表","循环风机","温湿度传感器","冷风进风设备","综合评价"],
       [(21,.06,.51),(21,.48,.95)],
       "这一页与前两页相同，先检查关联项目编号，再依次展开5个栏目。其中自控仪表需要拍照，其余设备在9.29说明中按只评价处理。")
item("（一）自控仪表",
     "“自控仪表状态*”“自控仪表-最新修复、改造年度*”“自控仪表-正常设备拍摄全景图*”。",
     "显示、按键、控制功能和连接正常选“正常”；显示异常、失控、按键失灵等选“损坏”；应装而未装选“缺失”。年度以最近一次维修或更新为准。",
     "拍自控仪表的正面全景，尽量能辨认品牌、设备编号、显示界面；损坏再补拍问题部位。9.29说明要求自控仪编号通过照片保留。",
     [(18,.18,.74)])
item("（二）循环风机",
     "“循环风机状态*”“循环风机-最新修复、改造年度*”。",
     "风机能正常运行、无明显异响和损坏选“正常”；电机故障、叶片损坏、不能正常运转选“损坏”；未配置应有的风机选“缺失”。最近维修年度单独选择，不能直接套用自控仪表的年份。",
     "9.29说明规定循环风机只评价、不拍部件照片；如果最新表单没有上传框，填状态和年度即可。",
     [(16,.17,.61)])
item("（三）温湿度传感器",
     "“温湿度传感器状态*”“温湿度传感器-最新修复、改造年度*”。",
     "传感器完整、测温测湿能正常使用选“正常”；探头损坏、显示异常或失灵选“损坏”；探头缺失选“缺失”。维修年度以实际记录为准。",
     "9.29说明规定该设备只评价，不另拍部件照。",
     [(14,.18,.53)])
item("（四）冷风进风设备",
     "栏目名称为“冷风进风设备”，但截屏展开后，下面出现的是“循环风机状态*”“循环风机-最新修复、改造年度*”文字。",
     "应检查冷风进风口、风门及相关设备是否存在、是否可正常进风。特别注意：截图中的字段名与栏目名不一致，可能存在表单配置问题；不要把循环风机状态直接冒充冷风进风设备状态，先向烟站或系统管理员核实表单对应关系，确认后按系统实际字段填。",
     "9.29说明对此项同样要求只评价、不拍部件照。",
     [(14,.48,.91)])
item("（五）综合评价",
     "“评价选择*”，页面列出“优良、较好、较差”的判定说明，但本模块所拍下拉选项实际是“优良、较好、差”。",
     "综合仪表、风机、传感器和进风设备使用情况。系统说明仍按2023—2026年性能稳定、2018—2022年性能良好、2017年及以前老化较重区分；实际下拉出现“差”就按“差”选，不要照着其他模块写“较差”。",
     "评价是对这一类设备整体情况的判断；故障明显影响烘烤时不能因个别零件能用就直接选“优良”。",
     [(13,.30,.83),(12,.39,.70)])
label("本入口完成后","确认四项设备状态与各自年份对应，仪表照片上传完成；对“冷风进风设备”字段名异常已核实处理，最后提交。")

module("四、附属设施",
       ["装烟棚","烤房加盖","烟夹","备用电源","电线设施","其他设施","综合评价"],
       [(10,.04,.49),(10,.47,.96)],
       "这一页共有7个栏目。前两个主要填状态、年度和拍照；中间“烟夹、备用电源、电线设施”是有无判断；“其他设施”是多项勾选；最后做综合评价。不要把所有项目都当成同一种状态填写。")
item("（一）装烟棚",
     "“装烟棚状态*”“装烟棚-最新修复、改造年度*”“装烟棚-正常设备全景照片*”。",
     "有装烟棚，棚架和顶棚完好、能够使用选“正常”；支架断裂、顶棚漏雨等选“损坏”；应有装烟棚却没有选“缺失”。年份选择实际修建或更新、维修记录对应的年份。",
     "拍到装烟棚整体，包括棚架和顶棚；损坏另拍具体破损处。",
     [(9,.17,.73)])
item("（二）烤房加盖",
     "“烤房加盖状态*”“烤房加盖-最新修复、改造年度*”“烤房加盖-正常设备拍摄全景图*”。",
     "屋顶加盖设施完整、不漏雨且正常使用选“正常”；破损、漏雨、脱落选“损坏”；应有加盖却没有选“缺失”。先看屋顶实际情况，不单凭历史登记判断。",
     "从能看清屋顶加盖情况的位置拍全景；如漏雨或破损再补拍细节。",
     [(8,.18,.70)])
item("（三）烟夹",
     "“烟夹是否有*”，下拉候选项为“是、否”。",
     "现场存在配套烟夹，选“是”；没有配套烟夹，选“否”。只判断是否真实存在，不要因为以前购买过就选“是”。",
     "这个栏目不是“正常/损坏/缺失”三级状态，也不是照片上传项。",
     [(7,.19,.68)])
item("（四）备用电源",
     "“备用电源是否有*”，下拉为“是、否”。",
     "现场有能够为烤房配套使用的备用供电设备，选“是”；没有选“否”。要核实实物，不能把普通供电线路当作备用电源。",
     "若确有备用电源但无法使用，应如实向管理人员说明故障，不能把“有”理解成“功能正常”。",
     [(6,.13,.39)])
item("（五）电线设施",
     "“电线设施是否有*”，下拉为“是、否”。",
     "实际有配套电线或供电设施，选“是”；没有选“否”。线路若存在裸露、严重老化等安全问题，应另行记录并报告，不要把存在情况当作安全合格结论。",
     "画面中的“是”仅是该演示烤房所选，不能复制到所有烤房。",
     [(4,.21,.57)])
item("（六）其他设施",
     "“是否存在以下设备*”，多选框列出“编烟机、操作台、监控”。",
     "现场有编烟机，就在“编烟机”前打勾；有操作台就勾“操作台”；安装了监控就勾“监控”。三个都存在可勾三个，只存在一项就只勾一项。没有的项目不要勾选。",
     "若三个都没有而表单又提示必填，先反馈管理员确认“无”的录入方式，不要为提交成功虚假勾选。",
     [(2,.19,.68)])
item("（七）综合评价",
     "“评价选择*”，截图中的评价说明为“优良、较好、较差”。",
     "根据装烟棚、屋顶加盖、烟夹、备用电源、线路和配套设备的整体完好程度选择；优良为性能稳定、预计可用5—8年；较好为仅需常规管护、预计可用3—5年；较差为老化明显、需重点维修或更新、预计可用1—3年。仍以现场实际使用情况和维修情况为准。",
     "评价前先核对“烟夹、备用电源、电线设施”的是/否及三项多选是否真实。",
     [(3,.33,.83)])
label("本入口完成后","逐项核对7个栏目，尤其“烟夹是否有”“备用电源是否有”“电线设施是否有”和“其他设施”多选，确认无误后提交。")

h("五、照片上传、年份选择和提交")
label("照片上传","在需要照片的分组点击“拍照上传”，会出现“系统相机、浏览器相机”，选能正常拍摄的一项，拍完返回表单看图片是否已显示。不能只拍照不上传，也不能拿其他烤房照片代替本烤房。")
photo([(38,.35,.64),(37,.47,.88)],"通用：选择相机并拍摄设备",max_h=6.7)
label("年份选择","点击“最新修复、改造年度*”，在年份滚轮中上下滑动到实际年份，点“确定”。不同部件可能有不同修复年份，不要全部选同一年。")
photo([(39,.37,.77)],"通用：维修年份滚轮中的“取消”和“确定”",max_h=6.7)
label("保存与提交","页面下方有“保存”“提交”两个按钮。采集中可按页面功能保存当前内容；各项必填和照片全部核对完成后点击“提交”，并确认显示成功提示。提交后通过右上角“×”关闭本表单，进入下一个按钮。")
label("信号差的情况","11.13版指南提供“异地上传”入口；现场先记录正确烤房编号和设备状态，保存清晰照片，待网络恢复后按授权流程补录，并由烟叶办审核。")
label("项目编号","原图出现23-420527KF00144、23-420527KF00147、23-420527KF00138等不同烤房。每进入一个新入口一定核对“关联项目编号*”，不要把不同编号的数据或照片混在一起。")

h("六、截图时间和页面关系")
para("解压后41张图片的ZIP记录修改时间均为2026-11-17 20:42:46，且没有可用的EXIF拍摄时间；因此依据文件名中20:24—20:31的截图时刻恢复实际查看过程，再按四个入口及其折叠分组重新排列正文。下表用于核对截图来源，不需要采集人员按拍摄时间填写。")
groups=[
("20:24—20:26","主体门窗","截图040、039—027、036—031","门窗字段、材质、综合评价、定位、拍照"),
("20:27—20:28","加热设备","截图026—022、020","基础信息、加热设备类型和散热器"),
("20:28—20:30","循环自控","截图021、018—012","仪表、风机、传感器、进风设备和评价"),
("20:30—20:31","附属设施","截图011—002","装烟棚、加盖、烟夹、备用电源、线路、其他设施"),
("20:31","其他—基础信息完善","截图001—000","使用状态选择，属于11.13版基础信息指南内容")
]
tbl=doc.add_table(rows=1,cols=3);tbl.style="Table Grid"
for i,v in enumerate(["截图时间","对应页面","在本指南中的用途"]):run(tbl.rows[0].cells[i].paragraphs[0],v,True,11)
for time,mod,ids,desc in groups:
    rr=tbl.add_row().cells
    for i,v in enumerate([time,mod,ids+"；"+desc]):run(rr[i].paragraphs[0],v,size=10)
para("本次已核对全部41张截图。重复的年份滚轮、相机弹层和勾选变化不再每张铺满一页；正文只保留与每个字段直接对应、最容易看懂的截图。每张配图下方标有具体用途。",size=12)
footer=s.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
run(footer,"烤房四大部位信息采集逐项填写指南　11172125 V2",size=9)

doc.save(OUTPUT)
with zipfile.ZipFile(OUTPUT) as z:
    assert z.testzip() is None
    cnt=sum(n.startswith("word/media/") for n in z.namelist())
print(f"DOC_READY path={OUTPUT} bytes={OUTPUT.stat().st_size} images={cnt} figures={fig}")
