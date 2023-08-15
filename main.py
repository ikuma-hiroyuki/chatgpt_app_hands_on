import chatgpt
from output_excel import excel_path, is_open_output_excel, output_excel

if excel_path.exists() and is_open_output_excel():
    print(f"{excel_path.name} が開かれているため処理を中断します。")
else:
    log, summary = chatgpt.chat_runner()
    output_excel(log, summary)
