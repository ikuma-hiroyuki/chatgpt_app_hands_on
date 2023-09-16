from chatgpt import chat_runner
from output_excel import output_excel, is_output_open_excel, excel_path

is_excel_open = is_output_open_excel()
if not is_excel_open:
    log, summary = chat_runner()
    output_excel(chat_log=log, chat_summary=summary)
else:
    print(f"{excel_path.name}が開かれているため、チャットを開始できませんでした。")
