"""
深入财务研报分析与生成脚本（简洁版）
基于自动化采集与分析的财务研报汇总，结合大模型能力，生成详细的公司财务、股权、行业、估值、治理结构等多维度深度分析与投资建议。
"""

import os
import yaml
from datetime import datetime
from data_analysis_agent.config.llm_config import LLMConfig
from data_analysis_agent.utils.llm_helper import LLMHelper
import re
import shutil
import requests
from urllib.parse import urlparse
from utils import save_markdown, format_markdown, convert_to_docx, extract_images_from_markdown, create_markdown_toc

def load_report_content(md_path):
    with open(md_path, "r", encoding="utf-8") as f:
        return f.read()

def get_background():
    return '''
本报告基于自动化采集与分析流程，涵盖如下环节：
- 公司基础信息等数据均通过akshare、公开年报、主流财经数据源自动采集。
- 财务三大报表数据来源：东方财富-港股-财务报表-三大报表 (https://emweb.securities.eastmoney.com/PC_HKF10/FinancialAnalysis/index)
- 主营业务信息来源：同花顺-主营介绍 (https://basic.10jqka.com.cn/new/000066/operate.html)
- 股东结构信息来源：同花顺-股东信息 (https://basic.10jqka.com.cn/HK0020/holder.html) 通过网页爬虫技术自动采集
- 行业信息通过DuckDuckGo等公开搜索引擎自动抓取，引用了权威新闻、研报、公司公告等。
- 财务分析、对比分析、估值与预测均由大模型（如GPT-4）自动生成，结合了行业对标、财务比率、治理结构等多维度内容。
- 相关数据与分析均在脚本自动化流程下完成，确保数据来源可追溯、分析逻辑透明。
- 详细引用与外部链接已在正文中标注。
- 数据接口说明与免责声明见文末。
'''


def get_requirements():
    return'''
- 一、「表」类内容规范  
  - 纯文字表（如发展历程、股权结构、政策总结、产品性能参数等）  
    - 报告正文：需明确标注资料来源（如“数据来源：公司年报”）。  
  - 涉及数字但不涉及计算（如产能规划列示、募投项目及金额、股权激励条件等）  
    - 报告正文：需明确标注资料来源。  
  - 涉及数字且涉及计算（如市场空间测算、盈利预测分拆、目标空间计算等）  
    - 报告正文：需明确标注资料来源。  

- 二、「图」类内容规范  
  - 纯文字图（如业务流程图、业务分类、发展曲线等）  
    - 报告正文：需明确标注资料来源。  
  - 涉及数字的截图（如外部预测的市场空间、经营数据等）  
    - 报告正文：需明确标注资料来源。  
  - 涉及数字且带有原始数据（如各项财务指标、行业竞争结构、行业关键数据等）  
    - 报告正文：需明确标注资料来源。  

- 三、「正文」类内容规范  
  - 涉及比较的表述（如龙头、领军、前列、第一梯队、优于、强于等）  
    - 要求：正文需说明判断依据，避免使用夸大性表述。  
  - 涉及文字中数字表述（如市场空间、产能储量、统计汇总等）  
    - 要求：正文需说明资料来源；如有对应截图，建议放入底稿。  
  - 涉及客观事实（如资本运作、企业战略、产能规划等）  
    - 要求：正文需说明资料来源；如有对应截图，建议放入底稿。  
  - 涉及主观观点及判断（如逻辑推论、主观判断、预测分析等）  
    - 要求：正文需增加提示性文字（如“我们认为”、“我们预测”），并补充判断依据。  

- 四、通用注释规则  
  - 资料来源需使用中文全称（如“FED”需改为“美联储”或“美国联邦储备委员会”）。
'''

def get_section_prompt():
     # 专属一级标题的提示
     special_title_prompts = {
         '盈利预测与估值': '本节请重点输出未来 3-5 年的盈利预测模型与估值区间，给出DCF表格和敏感性分析；',
         '风险提示': '''
         本节请突出主要风险因素，一般仅有4-5个，并对每一条给出简要缓释策略;
         示例：
         1.国补以旧换新不及预期 
         国补带来的以旧换新增量，是内需的重要因素。如果以旧换新带来的需求不及预期，那么对公司收入将产生影响。
         2.海外订单退坡速度超预期
         欧美和第三世界的需求支撑了公司外销的增长，若全球关税阴霾延续，公司海外订单退坡可能超预期，将影响公司收入水平。
         3.原材料价格暴涨
         白电的成本大头是原材料，如果原材料如铜、钢出现暴涨，那么公司短期无法向下游传导，将影响公司盈利能力。
         4.行业竞争加剧
         白电优秀的格局是盈利能力源头之一，如果行业竞争加剧，龙头追求收入规模增长掀起价格战，将影响公司的盈利能力。
         ''',
     }
     # 根据 desc 里出现的关键词来附加提示
     desc_keyword_prompts = {
    # 第一部分：公司概况
    '公司发展历程':
        "结构：按“创立阶段”“成长阶段”“成熟阶段”三个阶段撰写，每段列出关键时间节点与重大事件；"
        "图表：里程碑时间轴（Timeline），标注公司成立、融资、并购、上市等核心事件；"
        "数据：公司成立时间、各轮融资规模与时间点、并购标的名称与金额、上市日期及发行价格等原始时间节点数据；",

    '主营业务':
        "结构：先概述主要业务板块划分，再依次描述各板块的产品/服务类型、客户群体与市场份额；"
        "图表：饼图或环形图展示各业务板块收入占比；折线图展示最近 3–5 年各板块收入走势；"
        "数据：各业务板块近 3–5 年年度收入、收入占比、主要客户名单及其贡献度；",

    '股权结构':
        "结构：先给出主要股东分类（创始/管理层、机构投资者、公众股东），再回顾历次股权变动；"
        "图表：堆积柱状图展示各股东类别持股比例变化；表格列出 Top N 股东及持股比例；"
        "数据：最新一期股东名册（持股比例、持股数量）、历次股权融资/增持数据、限售股解禁数据；",

    '股权激励':
        "结构：说明激励计划框架（激励对象、激励工具、授予/行权时间），并总结执行进度与行权情况；"
        "图表：甘特图展示授予与行权时间节点；表格汇总激励对象人数与激励工具数量；"
        "数据：激励计划文件（授予数量、行权价格、行权期限）、已行权和未行权明细、参与人员名单；",

    '财务数据':
        "结构：分“主要利润表指标”“资产负债表要点”“现金流量表要素”三部分，分别列出核心财务指标及同比/环比；"
        "图表：双Y轴折线图展示营业收入与净利润趋势；雷达图或条形图对比关键财务比率（毛利率、净利率、ROE 等）；"
        "数据：近 3–5 年或更多年度的损益表、资产负债表、现金流量表原始数值，以及各期同比/环比计算结果；",

    '产品矩阵':
        "结构：列出主要产品/服务线及其定位，说明研发/在研/量产状态和目标市场；"
        "图表：流程图或结构图展示产品/服务生态；表格列产品名称、进度与市场贡献；"
        "数据：产品清单（名称、型号）、在研/测试/量产状态、各产品年度出货量与销售收入；",

    # 第二部分：行业与竞争
    '公司所在行业规模':
        "结构：先给出行业总体规模和近年增长率，再分析宏观经济和政策驱动因素；"
        "图表：柱状图或折线图展示近 5 年行业市场规模及 CAGR；可选地图热力图展示区域分布；"
        "数据：行业市场规模（年度）、CAGR 计算输入数据、政策文件摘要和宏观经济指标；",

    '细分领域规模':
        "结构：拆分主要细分市场，分别描述各细分领域的规模、增长逻辑与市场份额；"
        "图表：分组柱状图展示细分市场在不同年份的规模；饼图展示各细分领域份额；"
        "数据：各细分领域年度市场规模、各领域主要参与者占比、细分市场增长率；",

    '上下游产业链情况':
        "结构：绘制产业链框架，说明上游原材料/技术环节与下游应用场景及典型客户；"
        "图表：产业链流程图；Sankey 图或箭线图展示价值/物料流向；"
        "数据：上游原材料价格与采购量、中游生产成本、下游客户订单量及合作伙伴名单；",

    '可比公司情况':
        "结构：选取若干可比公司，先说明可比标准，再对比估值水平、盈利能力和增长性；"
        "图表：条形图或柱状图对比 P/E、P/B 等估值指标；散点图展示盈利率 vs 增长率；"
        "数据：可比公司财务数据（收入、利润、估值指标）、行业研究报告中的估值区间；",

    '公司所持有的技术的增长潜力等内容':
        "结构：说明核心技术类型和成熟度，分析研发投入与未来市场应用前景；"
        "图表：漏斗图或分层图展示技术从实验室到量产的成熟度；折线图对比研发投入与专利/成果数量；"
        "数据：年度 R&D 投入、专利数量与类型、技术验证/测试里程碑数据；",

    # 第三部分：核心竞争力
    '客户结构':
        "结构：按行业/区域/规模对客户进行分类，分别列出各类客户的营收贡献及变化；"
        "图表：旭日图或分级饼图展示客户层级结构；条形图展示 Top 客户年度营收占比；"
        "数据：客户清单（行业分类、地域分布）、各客户年度采购额及占比；",

    '专利情况':
        "结构：统计总专利数量及分类，重点点评核心专利及其地理/技术分布；"
        "图表：柱状图展示各技术类别专利数量；世界/国内地图展示主要专利布局；"
        "数据：专利申请/授权数量（年度）、专利技术分类、地区分布数据；",

    '技术人才':
        "结构：给出研发团队整体规模与结构，简要介绍主要研发/管理团队背景与分布；"
        "图表：堆积柱状图展示研发 vs 非研发人员占比；组织结构图或关系网展示核心团队；"
        "数据：员工总数、研发/非研发人员数量、管理层背景及任职年限；",

    '产品特色':
        "结构：挑选 2–3 个代表性产品/服务，用“功能—优势—应用案例”模板撰写；"
        "图表：功能模块示意图；案例前后效果对比图（柱状或对比表）；"
        "数据：各产品技术参数、用户反馈或案例数据（性能指标、成本效益）；",

    '最新情况':
        "结构：梳理最近 6–12 个月的重要进展（如新合作、新订单、新产品发布）；"
        "图表：时间轴（Timeline）标注关键进展；折线图展示新增订单或新增用户增长；"
        "数据：近年重要合同/订单金额与时间、合作伙伴名单、产品发布/版本号；",

    '发行可转债、再融资的募投项目介绍':
        "结构：分别说明融资规模、发行/使用要点及募投项目背景与预期回报；"
        "图表：饼图展示募资用途分布；甘特图或里程碑图展示项目实施计划与关键节点；"
        "数据：可转债/再融资金额、发行价格、募集资金投向明细、项目预算与预期收益；"
    }
     return special_title_prompts, desc_keyword_prompts

def get_llm():
    api_key = os.environ.get("OPENAI_API_KEY")
    base_url = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
    model = os.environ.get("OPENAI_MODEL", "gpt-4")
    llm_config = LLMConfig(api_key=api_key, base_url=base_url, model=model)
    return LLMHelper(llm_config)

def generate_outline(llm, background, report_content):
    outline_prompt = f"""
你是一位顶级金融分析师和研报撰写专家。请基于以下背景和财务研报汇总内容，生成一份详尽的《商汤科技公司研报》分段大纲，要求：
- 以yaml格式输出，务必用```yaml和```包裹整个yaml内容，便于后续自动分割。
- 大纲需为多级结构，具体要求如下：
  - 一级标题：为每个主要部分生成有吸引力的章节标题（不能照抄章节结构部分），如“稀缺光芯片资产，高速率产品持续放量”。
  - 每个一级标题下，包含最多三个二级标题，每个二级标题需包含：
    - title: 二级标题
    - desc: 对应段落的内容，要与章节结构下的名词一致，可以自由组合
- 章节结构如下：
  - 第一章 
    - 公司发展历程
    - 主营业务
    - 股权结构
    - 股权激励
    - 产品矩阵
    - 财务数据
  - 第二章 
    - 公司所在行业规模
    - 细分领域规模
    - 上下游产业链情况
    - 可比公司情况
    - 公司所持有的技术的增长潜力等内容
  - 第三章 公司的核心竞争力
    - 客户结构
    - 专利情况
    - 技术人才
    - 产品特色
    - 最新情况
    - 发行可转债、再融资的募投项目介绍
  - 盈利预测与估值
  - 风险提示
- 章节标题不能照抄章节结构部分，而是需要根据内容生成生动且有吸引力的标题，但是绝对不能与章节内容无关。"源杰科技"示例如下
  - 第一章：稀缺光芯片资产，高速率产品持续放量
    - 国内光芯片领军企业，产品升级驱动业绩增长
    - 光芯片持续迭代升级，高速率产品占比提升
    - 业绩稳健增长，盈利能力突出
  - 第二章：光芯片迭代升级，高端产品国产化空间广阔.
  - 第三章：研发+制造能力领先，推动高端光芯片国产化
  - 盈利预测与估值
  - 风险提示

- 只输出yaml格式的分段大纲，不要输出正文内容。
- 输出示例（仅供格式参考）：
```yaml
- 一级标题: 稀缺光芯片资产，高速率产品持续放量
  二级标题:
    - title: 国内光芯片领军企业，产品升级驱动业绩增长
      desc: 公司发展历程、管理团队、股权结构、股权激励
    - title: 光芯片持续迭代升级，高速率产品占比提升
      desc: 主营业务、产品矩阵
    - title: 业绩稳健增长，盈利能力突出
      desc: 财务数据
- 一级标题: 光芯片迭代升级，高端产品国产化空间广阔
  二级标题:
    - title: 光芯片是光通信产业链核心元件
      desc: 公司所在行业规模、上下游产业链情况
    - title: 数字经济驱动流量增长，光芯片量价齐升
      desc: 细分领域规模
```

【背景说明开始】
{background}
【背景说明结束】

【财务研报汇总内容开始】
{report_content}
【财务研报汇总内容结束】
"""
    outline_list = llm.call(
        outline_prompt,
        system_prompt="你是一位顶级金融分析师和研报撰写专家，善于结构化、分段规划输出，分段大纲必须用```yaml包裹，便于后续自动分割。",
        max_tokens=4096,
        temperature=0.3
    )
    print("\n===== 生成的分段大纲如下 =====\n")
    print(outline_list)
    
    parts = []
    try:
        # 提取YAML内容
        if '```yaml' in outline_list:
            yaml_block = outline_list.split('```yaml')[1].split('```')[0]
        elif '```' in outline_list:
            # 处理没有yaml标识的情况
            yaml_block = outline_list.split('```')[1].split('```')[0]
        else:
            yaml_block = outline_list.strip()
        
        def extract_outline(yaml_block):
            """
            输入：yaml_block（字符串，已去除markdown包裹）
            输出：嵌套结构的列表，每项为{'一级标题':..., '二级标题': [{'title':..., 'desc':...}, ...]}
            """
            parsed_data = yaml.safe_load(yaml_block)
            outline = []
            for first_level in parsed_data:
                first_title = first_level.get('一级标题', '')
                second_list = first_level.get('二级标题', [])
                outline.append({
                    '一级标题': first_title,
                    '二级标题': [
                        {'title': s.get('title', ''), 'desc': s.get('desc', '')}
                        for s in second_list
                    ]
                })
            return outline

        # 尝试解析YAML        
        parts = extract_outline(yaml_block)
        
    except Exception as e:
        print(f"[大纲yaml解析失败] {e}")
        parts = []

    return parts

def generate_section(llm, first_title, second_title, prev_content, background, requirements, part_desc, report_content, is_last):
    section_prompt = f"""
你是一位顶级金融分析师和研报撰写专家。请基于以下内容，直接输出"{second_title}"这一部分的完整研报内容。

**重要要求：**
1. 直接输出完整可用的研报内容，以"### {second_title}"开头，可以分段但不要再生成新的子标题。
2. 在正文中引用数据、事实、图片等信息时，适当位置插入参考资料符号（如[1][2][3]），符号需与文末引用文献编号一致
3. **图片引用要求（务必严格遵守）：**
   - 只允许引用【财务研报汇总内容】中真实存在的图片地址（格式如：./images/图片名字.png），必须与原文完全一致。
   - 禁止虚构、杜撰、改编、猜测图片地址，未在【财务研报汇总内容】中出现的图片一律不得引用。
   - 如需插入图片，必须先在【财务研报汇总内容】中查找，未找到则不插入图片，绝不编造图片。
   - 如引用了不存在的图片，将被判为错误输出。
4. 不要输出任何【xxx开始】【xxx结束】等分隔符
5. 不要输出"建议补充"、"需要添加"等提示性语言
6. 不要编造图片地址或数据
7. 内容要详实、专业，可直接使用

**数据来源标注：**
- 财务数据标注：（数据来源：东方财富-港股-财务报表[1]）
- 主营业务信息标注：（数据来源：同花顺-主营介绍[2]）
- 股东结构信息标注：（数据来源：同花顺-股东信息网页爬虫[3]）

【本次任务标题】

{second_title}

【本次任务一级标题】
{first_title}

【本次任务描述】
{part_desc}


【已生成前文】
{prev_content}

【格式要求】
{requirements}

【背景说明开始】
{background}
【背景说明结束】

【财务研报汇总内容开始】
{report_content}
【财务研报汇总内容结束】
"""
    if is_last:
        section_prompt += """
请在本节最后以"引用文献"格式，列出所有正文中用到的参考资料，格式如下：
[1] 东方财富-港股-财务报表: https://emweb.securities.eastmoney.com/PC_HKF10/FinancialAnalysis/index
[2] 同花顺-主营介绍: https://basic.10jqka.com.cn/new/000066/operate.html
[3] 同花顺-股东信息: https://basic.10jqka.com.cn/HK0020/holder.html
"""
    section_text = llm.call(
        section_prompt,
        system_prompt="你是顶级金融分析师，专门生成完整可用的研报内容。输出必须是完整的研报正文，无需用户修改。严格禁止输出分隔符、建议性语言或虚构内容。只允许引用真实存在于【财务研报汇总内容】中的图片地址，严禁虚构、猜测、改编图片路径。如引用了不存在的图片，将被判为错误输出。",
        max_tokens=8192,
        temperature=0.5
    )
    return section_text

def main():
    # ====== 图片路径预处理，自动生成本地 images 路径的 markdown 文件 ======
    raw_md_path = "财务研报汇总_20250630_215159.md"  # 原始输入 markdown
    new_md_path = "财务研报汇总_20250630_215159_images.md"  # 处理后输出 markdown
    images_dir = os.path.join(os.path.dirname(raw_md_path), 'images')
    extract_images_from_markdown(raw_md_path, images_dir, new_md_path)

    # 后续流程用 new_md_path
    report_content = load_report_content(new_md_path)
    background = get_background()

    requirements = get_requirements()
    special_title_prompts, desc_keyword_prompts = get_section_prompt()
    llm = get_llm()
    # parts 是一个列表，内含 dict：{'一级标题': str, '二级标题': [ {'title': str, 'desc': str}, ... ]}
    parts = generate_outline(llm, background, report_content)

    # ====== 为Markdown创建目录 ======
    toc = create_markdown_toc(parts)
    print(toc)

    # ====== 初始化报告内容，添加标题、摘要和目录 ======
    full_report = [f"# 商汤科技公司研报\n\n{toc}"]
    prev_content = ''
    # 计算所有小节总数，方便判断最后一个
    total_sections = sum(len(p['二级标题']) for p in parts)
    section_counter = 0
    for part in parts:
        first_title = '## ' + part['一级标题']
        full_report.append(first_title)
        for sub in part['二级标题']:
            section_counter += 1
            section_title = sub['title']
            section_desc  = sub['desc']
            print(f"\n===== 正在生成小节：{section_title} =====\n")

            # 2. 先看一级标题专项覆盖
            if first_title in special_title_prompts:
                system_prompt = special_title_prompts[first_title]
            else:
                # 3. 按 desc 切分关键词，累加匹配到的提示词
                system_prompt_parts = []
                for kw, tip in desc_keyword_prompts.items():
                    # 若 desc 中含有 kw，则添加对应提示
                    if kw in section_desc:
                        system_prompt_parts.append(tip)
                # 如果有至少一条匹配，就合并，否则用默认
                if system_prompt_parts:
                    system_prompt = ' '.join(system_prompt_parts)
                else:
                    system_prompt = ''

            # 最后一节时 is_last=True
            is_last = (section_counter == total_sections)

            section_text = generate_section(llm, first_title, section_title, prev_content, background, requirements, system_prompt, report_content, is_last)
            full_report.append(section_text)
            print(f"\n===== 已生成小节：{section_title}（预览前2000字符） =====\n")
            print(section_text[:2000])
            print("\n===== 本小节内容结束 =====\n")
            prev_content = '\n'.join(full_report)
    final_report = '\n\n'.join(full_report)
    output_file = f"深度财务研报分析_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
    save_markdown(final_report, output_file)
    format_markdown(output_file)
    convert_to_docx(output_file)

if __name__ == "__main__":
    main()
