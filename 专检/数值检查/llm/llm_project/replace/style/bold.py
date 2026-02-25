from llm.llm_project.parsers.body_extractor import extract_body_text
from llm.llm_project.parsers.footer_extractor import extract_footers
from llm.llm_project.parsers.header_extractor import extract_headers

if __name__ == "__main__":
    # 示例文件路径
    original_path = r"C:\Users\Administrator\Desktop\project\效果\TP251222006，香港资翻译，中译英（字数1.7w）\原文-RX-96 LAT Report Vol 19 - Zongtian Contract (pages 4-30).docx"  # 请替换为原文文件路径
    translated_path = r"C:\Users\Administrator\Desktop\project\效果\TP251222006，香港资翻译，中译英（字数1.7w）\译文-RX-96 LAT Report Vol 19 - Zongtian Contract (pages 4-30).docx"  # 请替换为译文文件路径
    #处理页眉
    original_header_text=extract_headers(original_path)
    translated_header_text=extract_headers(translated_path)
    #处理页脚
    original_footer_text = extract_footers(original_path)
    translated_footer_text = extract_footers(translated_path)
    #处理正文(含脚注/表格/自动编号)
    original_body_text=extract_body_text(original_path)
    translated_body_text=extract_body_text(translated_path)
    print("======页眉===========")
    print(original_header_text)
    print(translated_header_text)
    print("======页脚===========")
    print(original_footer_text)
    print(translated_footer_text)
    print("======正文===========")
    print(original_body_text)
    print(translated_body_text)