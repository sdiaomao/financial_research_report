# ✅ 基础导入（保留原有功能 + 新增 Word 操作）
import os
import re
import shutil
import yaml
import requests
from datetime import datetime
from urllib.parse import urlparse

# ✅ Word 生成报告相关库（用于封面、摘要、目录）
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def generate_summary(llm, background, report_content):
    """使用LLM生成摘要内容"""
    summary_prompt = f"""
你是一位顶级金融分析师和研报撰写专家。请基于以下背景和财务研报汇总内容，生成一份简洁的摘要（不超过500字）。
摘要需涵盖公司核心亮点、行业地位、财务表现和投资价值等关键信息。

【背景说明开始】
{background}
【背景说明结束】

【财务研报汇总内容开始】
{report_content}
【财务研报汇总内容结束】
"""
    summary_text = llm.call(
        summary_prompt,
        system_prompt="你是顶级金融分析师，专门生成简洁、有洞察力的研报摘要。",
        max_tokens=1024,
        temperature=0.5
    )
    return summary_text

def create_markdown_toc(parts):
    """为Markdown文件创建手动目录"""
    toc = ["# 目录", ""]
    for part in parts:
        if isinstance(part, dict):
            title = part.get('part_title', '')
            # 转换标题为Markdown链接格式（移除特殊字符，用短横线连接）
            anchor = title.lower().replace(' ', '-').replace(':', '').replace('(', '').replace(')', '')
            toc.append(f"- [{title}](#{anchor})")
    toc.append("")  # 添加空行分隔目录和正文
    return "\n".join(toc)


def load_report_content(md_path):
    with open(md_path, "r", encoding="utf-8") as f:
        return f.read()

def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def is_url(path):
    return path.startswith('http://') or path.startswith('https://')

def download_image(url, save_path):
    try:
        resp = requests.get(url, stream=True, timeout=10)
        resp.raise_for_status()
        with open(save_path, 'wb') as f:
            for chunk in resp.iter_content(1024):
                f.write(chunk)
        return True
    except Exception as e:
        print(f"[下载失败] {url}: {e}")
        return False

def copy_image(src, dst):
    try:
        shutil.copy2(src, dst)
        return True
    except Exception as e:
        print(f"[复制失败] {src}: {e}")
        return False

def extract_images_from_markdown(md_path, images_dir, new_md_path):
    ensure_dir(images_dir)
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 匹配 ![alt](path) 形式的图片
    pattern = re.compile(r'!\[[^\]]*\]\(([^)]+)\)')
    matches = pattern.findall(content)
    used_names = set()
    replace_map = {}
    not_exist_set = set()

    for img_path in matches:
        img_path = img_path.strip()
        # 取文件名
        if is_url(img_path):
            filename = os.path.basename(urlparse(img_path).path)
        else:
            filename = os.path.basename(img_path)
        # 防止重名
        base, ext = os.path.splitext(filename)
        i = 1
        new_filename = filename
        while new_filename in used_names:
            new_filename = f"{base}_{i}{ext}"
            i += 1
        used_names.add(new_filename)
        new_img_path = os.path.join(images_dir, new_filename)
        # 下载或复制
        img_exists = True
        if is_url(img_path):
            success = download_image(img_path, new_img_path)
            if not success:
                img_exists = False
        else:
            # 支持绝对和相对路径
            abs_img_path = img_path
            if not os.path.isabs(img_path):
                abs_img_path = os.path.join(os.path.dirname(md_path), img_path)
            if not os.path.exists(abs_img_path):
                print(f"[警告] 本地图片不存在: {abs_img_path}")
                img_exists = False
            else:
                copy_image(abs_img_path, new_img_path)
        # 记录替换
        if img_exists:
            replace_map[img_path] = f'./images/{new_filename}'
        else:
            not_exist_set.add(img_path)

    # 替换 markdown 内容，不存在的图片直接删除整个图片语法
    def replace_func(match):
        orig = match.group(1).strip()
        if orig in not_exist_set:
            return ''  # 删除不存在的图片语法
        return match.group(0).replace(orig, replace_map.get(orig, orig))

    new_content = pattern.sub(replace_func, content)
    with open(new_md_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print(f"图片处理完成！新文件: {new_md_path}")


from data_analysis_agent.config.llm_config import LLMConfig
from data_analysis_agent.utils.llm_helper import LLMHelper

desc_list = [
    """
    公司简介：适当精炼即可。
    业务分布：将公司业务最新情况更新在此，包括图片（可截图）和细节描述（用表格），分为分地区业务收入和分板块的业务收入。
    历史沿革：公司的官网或者招股书会详细的总结公司的变化，如并购、重大产品更新等，那么运用流程图+年份变化的形式是最合适的方式
    股权结构:
    股权激励：股权激励是公司为绑定主要的核心业务人员、技术人才的红利，实行股权激励是一项利好，意味着公司的主要架构与发展方向不会出现重大变化，劲儿往一处使
    团队介绍：一般来讲，管理层和核心技术人员团队良好的学历背景、丰富的履历经验，一般在公司官网或招股书会披露董监高的简历情况
    财务数据：公司的营收、净利润、毛净利率、期间费用，合同负债、应收账款等情况简要概括，用较少的图表展示详尽的信息，合理运用饼图表示比例、柱线组合图表示变化和YOY、堆积条形、堆积柱状图表示业务构成、多线图表期间费用变化。在分析财务数据时，主要笔墨放在特殊拐点或是突变的地方，并解释变化的原因。
    """,
    """
    行业规模：行业的整体规模增长是公司拥有潜力的最基本条件，也预示着有一个良好的前景。来源一般是在行业的协会、或是大型的专业行业咨询公司数据。
    细分领域规模：细分领域则是公司所在市场的规模，其规模同企业的营收也是非常相关。
    国家竞争格局：一般是运用饼图将现如今的市场分布以国家方式分类。
    上下游产业链：界定公司所处的产业链的位置，判断是否拥有话语权，围绕“供给、需求”两个最为基准的方面，可以判断公司无论是向供应商还是向客户的议价能力，展示公司营收的稳定程度及对成本的控制能力。
    可比公司：可比公司进行对比可以展示公司相对于其他行业内的公司的独特优势或是特殊的细分赛道布局等。一般毛利率、研发费用、资本支出、产品特点是主要可以进行对比的指标。判断公司的竞争实力，持续发展情况等的重要指标。
    """,
    """
    客户优势：稳定、较为分散的客户群体是公司保持营收稳定核心，如果出现最大或前五大客户集中度过于高，那么一旦出现变动，则公司营收会出现大幅变动，影响公司股价。同时如果客户是业内的龙头公司，那么也可以侧面反映公司的产品十分受认可，其技术含量及质量有保证。
    技术专利与获奖情况：技术专利等无形资产反映公司的技术硬实力，技术专利较多，获得的奖项含金量高，也能体现公司的研发水准。
    在研项目：在研项目将演变为公司后续推出的新产品，新业务，当在研项目取得重大进展时，则对公司的收入具有较大利好。
    募投扩产：企业发行可转债、专项债券、募集配套资金完成某些产线或项目的建设，在达产后，也能助于公司业务成熟度提升，体量增大。
    """,
    "风险提示及盈利预测: 风险提示一般仅有4-5个小短句，盈利预测则是估值建模的内容。几大估值建模方法:现金流折现、可比公司、可比交易。首先去预测收入(分业务)，成本，折旧等情况，即每年增长的百分比，预测后几年的市盈率，股价情况等。",

]


def get_part_desc(idx):
    if 0 <= idx < len(desc_list):
        return desc_list[idx] or ''
    else:
        return ''


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
- 每一项为一个主要部分，每部分需包含：
  - part_title: 章节标题
  - part_desc: 本部分内容简介
- 章节结构如下：
  - 第一章 公司发展历程、主营业务、股权结构、股权激励、产品矩阵、财务数据。
  - 第二章 公司所在行业规模、公司所在细分领域规模、上下游产业链情况、可比公司情况、公司所持有的技术的增长潜力等内容。
  - 第三章 公司的核心竞争力:客户结构、专利情况、技术人才、产品特色、最新情况、发行可转债、再融资的募投项目介绍。
  - 风险提示及盈利预测: 风险提示一般仅有4-5个小短句，盈利预测则是估值建模的内容。几大估值建模方法:现金流折现、可比公司、可比交易。首先去预测收入(分业务)，成本，折旧等情况，即每年增长的百分比，预测(2024Eexpectation)后几年的市盈率，股价情况等。
  - 数据来源: 列明本报告所引用的主要数据来源和参考资料，确保报告的权威性和可追溯性。
- 章节标题不能照抄章节结构部分，而是需要根据内容生成生动且有吸引力的标题，但是绝对不能与章节内容无关。"源杰科技"示例如下
  - 第一章：稀缺光芯片资产，高速率产品持续放量
  - 第二章：光芯片迭代升级，高端产品国产化空间广阔.
  - 第三章：研发+制造能力领先，推动高端光芯片国产化
  - 投资建议
  - 风险提示

- 只输出yaml格式的分段大纲，不要输出正文内容。

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
    try:
        if '```yaml' in outline_list:
            yaml_block = outline_list.split('```yaml')[1].split('```')[0]
        else:
            yaml_block = outline_list
        parts = yaml.safe_load(yaml_block)

        # 修复：确保 parts 是正确的格式
        if isinstance(parts, dict):
            # 如果 parts 是字典，检查其值的类型
            if all(isinstance(v, dict) for v in parts.values()):
                parts = list(parts.values())
            else:
                # 如果值是字符串，转换为正确的格式
                parts = [{'part_title': k, 'part_desc': v} for k, v in parts.items()]
        elif isinstance(parts, list):
            # 如果 parts 是列表，确保每个元素都是字典
            if parts and not isinstance(parts[0], dict):
                # 如果列表元素是字符串，转换为字典格式
                parts = [{'part_title': f'部分{i + 1}', 'part_desc': part} for i, part in enumerate(parts)]
        else:
            parts = []
    except Exception as e:
        print(f"[大纲yaml解析失败] {e}")
        parts = []
    return parts


def generate_section(llm, part_title, prev_content, background, part_desc, report_content, is_last):
    section_prompt = f"""
你是一位顶级金融分析师和研报撰写专家。请基于以下内容，直接输出"{part_title}"这一部分的完整研报内容。

**重要要求：**
1. 直接输出完整可用的研报内容，以"## {part_title}"开头
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
{part_title}

【本次任务描述】
{part_desc}

【已生成前文】
{prev_content}

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
from datetime import datetime

def save_markdown(content, output_file):
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"\n📁 深度财务研报分析已保存到: {output_file}")

def format_markdown(output_file):
    try:
        import subprocess
        format_cmd = ["mdformat", output_file]
        subprocess.run(format_cmd, check=True, capture_output=True, text=True, encoding='utf-8')
        print(f"✅ 已用 mdformat 格式化 Markdown 文件: {output_file}")
    except Exception as e:
        print(f"[提示] mdformat 格式化失败: {e}\n请确保已安装 mdformat (pip install mdformat)")

def convert_to_docx(output_file, docx_output="Company_Research_Report.docx"):
    try:
        import subprocess
        import os
        pandoc_cmd = [
            "pandoc",
            output_file,
            "-o",
            docx_output,
            "--standalone",
            "--resource-path=.",
            "--extract-media=."
        ]
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        subprocess.run(pandoc_cmd, check=True, capture_output=True, text=True, encoding='utf-8', env=env)
        print(f"\n📄 Word版报告已生成: {docx_output}")
    except subprocess.CalledProcessError as e:
        print(f"[提示] pandoc转换失败。错误信息: {e.stderr}")
        print("[建议] 检查图片路径是否正确，或使用 --extract-media 选项")
    except Exception as e:
        print(f"[提示] 若需生成Word文档，请确保已安装pandoc。当前转换失败: {e}")


def main():
    # ====== 图片路径预处理，自动生成本地 images 路径的 markdown 文件 ======
    raw_md_path = "财务研报汇总_20250630_215159.md"  # 原始输入 markdown
    new_md_path = "财务研报汇总_20250630_215159_images.md"  # 处理后输出 markdown
    images_dir = os.path.join(os.path.dirname(raw_md_path), 'images')
    extract_images_from_markdown(raw_md_path, images_dir, new_md_path)

    # 后续流程用 new_md_path
    report_content = load_report_content(new_md_path)

    def get_background():
        # 这里可以写具体的背景说明内容
        return "商汤科技是一家专注于人工智能视觉技术的领先企业，提供先进的AI解决方案。"

    background = get_background()
    llm = get_llm()

    # ====== 新增：生成大纲 ======
    parts = generate_outline(llm, background, report_content)

    # ====== 新增：生成摘要 ======
    print("\n===== 正在生成摘要 =====\n")
    summary_text = generate_summary(llm, background, report_content)
    print(f"\n===== 摘要生成完成 =====\n")

    # ====== 新增：为Markdown创建目录 ======
    toc = create_markdown_toc(parts)

    # ====== 初始化报告内容，添加标题、摘要和目录 ======
    full_report = [f"# 商汤科技公司研报\n\n## 摘要\n{summary_text}\n\n{toc}"]
    prev_content = '\n'.join(full_report)

    # ====== 原有：按章节生成内容 ======
    for idx, part in enumerate(parts):
        part_desc = get_part_desc(idx)
        # 修复：安全地获取 part_title
        if isinstance(part, dict):
            part_title = part.get('part_title', f'部分{idx + 1}')
        else:
            part_title = str(part) if part else f'部分{idx + 1}'
        print(f"\n===== 正在生成：{part_title} =====\n")
        is_last = (idx == len(parts) - 1)
        section_text = generate_section(
            llm, part_title, prev_content, background, part_desc, report_content, is_last
        )
        full_report.append(section_text)
        print(f"\n===== 已生成：{part_title}（预览前2000字符） =====\n")
        print(section_text[:2000])
        print("\n===== 本部分内容结束 =====\n")
        prev_content = '\n'.join(full_report)

    # ====== 保存报告 ======
    final_report = '\n\n'.join(full_report)
    output_file = f"深度财务研报分析_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
    save_markdown(final_report, output_file)
    format_markdown(output_file)
    convert_to_docx(output_file)

if __name__ == "__main__":
    main()
