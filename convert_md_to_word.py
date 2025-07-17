import sys
from utils import convert_to_docx

def convert_md_to_word(md_file, docx_file=None):
    if docx_file is None:
        # 如果没有指定 docx 文件名，自动生成
        docx_file = md_file.replace('.md', '.docx')
    
    try:
        convert_to_docx(md_file, docx_file)
        print(f"✅ 转换成功！Word 文件已保存为: {docx_file}")
    except Exception as e:
        print(f"❌转换失败: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("使用方法: python convert_md_to_word.py <markdown文件> [word文件名]")
        print("示例: python convert_md_to_word.py 深度财务研报分析_202507172412")
        sys.exit(1)
    
    md_file = sys.argv[1]
    docx_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    convert_md_to_word(md_file, docx_file) 