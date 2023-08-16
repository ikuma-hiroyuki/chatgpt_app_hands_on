import os

import openai
from dotenv import load_dotenv
from colorama import Fore

EXIT_COMMAND = "exit()"
DEFAULT_MODEL = "gpt-3.5-turbo"


def print_error_message(message):
    """エラーメッセージを表示する"""
    print(f"{Fore.RED}{message}{Fore.RESET}")


def fetch_gpt_model_list() -> list[str] | None:
    """
    GPTモデルの一覧を取得する
    :return: GPTモデルのリストおよびNone(エラーのとき)
    """

    try:
        # APIから使えるモデルの一覧を取得する
        all_model_list = openai.Model.list()
    except (openai.error.APIError, openai.error.ServiceUnavailableError):
        print_error_message("OpenAI側でエラーが発生しています。少し待ってから再度試してください。")
        print("サービス稼働状況は https://status.openai.com/ で確認できます。")
    except openai.error.APIConnectionError:
        print_error_message("ネットワークに問題があります。設定を見直すか少し待ってから再度試してください。")
    except openai.error.AuthenticationError:
        print_error_message("APIキーまたはトークンが無効もしくは期限切れです。")
    except openai.error.OpenAIError:
        print_error_message("エラーが発生しました。少し待ってから再度試してください。")
    else:
        # GPTモデルのみを抽出する
        gpt_model_list = []
        for model in all_model_list.data:
            if "gpt" in model.id:
                gpt_model_list.append(model.id)

        # モデル名でソートする
        gpt_model_list.sort()

        # GPTモデルのリストを返す
        return gpt_model_list


def choice_model(model_list: list[str]) -> str:
    """
    GPTを選択させ、モデル名を返す
    :param model_list: GPTモデルのリスト
    :return: 選択されたモデル名
    """

    # モデルの一覧を表示する
    print("\nAIとのチャットに使用するモデルの番号を入力しEnterキーを押してしてください。")
    for i, model in enumerate(model_list):
        print(f"{i}: {model}")
    input_number = input(f"何も入力しない場合は{Fore.GREEN} {DEFAULT_MODEL} {Fore.RESET}が使われます。: ")

    while True:
        # 何も入力されなかったとき
        if not input_number:
            return DEFAULT_MODEL
        # 数字以外が入力されたとき
        if not input_number.isdigit():
            print(f"{Fore.RED}数字を入力してください。{Fore.RESET}")
        # 選択肢に存在しない番号が入力されたとき
        elif not int(input_number) in range(len(model_list)):
            print(f"{Fore.RED}その番号は選択肢に存在しません。{Fore.RESET}")
        # 正常な入力
        elif int(input_number) in range(len(model_list)):
            return model_list[int(input_number)]


def give_role_to_system():
    """
    AIアシスタントに与える役割を入力させる
    :return: 役割の辞書
    """
    print(f"\nAIアシスタントとチャットを始めます。")
    system_content = input("AIアシスタントに与える役割がある場合は入力してください。"
                           "ない場合はそのままEnterキーを押してください。: ")
    return {"role": "system", "content": system_content}


def input_user_prompt() -> str:
    """
    ユーザーからの入力を受け付ける
    :return: ユーザーのプロンプト
    """

    user_prompt = ""
    while not user_prompt:
        user_prompt = input(f"{Fore.CYAN}\nあなた: {Fore.RESET}")
        if not user_prompt:
            print(f"{Fore.YELLOW}プロンプトを入力してください。{Fore.RESET}")

    return user_prompt


def generate_chat_log(model: str, system_role: dict) -> list[dict]:
    """
    チャットのログを生成する
    :param model: チャットするときのモデル
    :param system_role: システムに与える役割
    :return: やりとりを格納したリスト
    """

    print("AIアシスタントとチャットを始めます。チャットを終了させる場合は"
          f"{Fore.GREEN} {EXIT_COMMAND} {Fore.RESET}と入力してください。")

    # チャットのログを保存するリスト
    chat_log: list[dict] = []

    # 役割がある場合は、役割を追加する
    if system_role["content"]:
        chat_log.append(system_role)

    while True:
        prompt = input_user_prompt()
        if prompt == 'exit()':
            break

        # プロンプトをログに追加
        chat_log.append({"role": "user", "content": prompt})

        # AIの応答を取得
        response = openai.ChatCompletion.create(model=model, messages=chat_log)
        # 応答のみを取り出す
        content = response.choices[0].message.content
        role = response.choices[0].message.role
        # ログに追加
        chat_log.append({"role": role, "content": content})
        # 応答を表示
        print(f"\n{Fore.GREEN}AIアシスタント:{Fore.RESET} {content}")

    # ループを抜けたら終了メッセージを表示
    print("\nチャットを終了します。")
    return chat_log


def get_initial_prompt(chat_log: list[dict]) -> str | None:
    """
    チャットの履歴からユーザーの最初のプロンプトを取得する。
    :param chat_log:
    :return: ユーザーの最初のプロンプト
    """
    initial_prompt = ""
    for log in chat_log:
        if log["role"] == "user":
            initial_prompt = log["content"]
            break

    if initial_prompt:
        return initial_prompt
    else:
        return None


def generate_summary(initial_prompt: str, gpt_model: str, summary_length: int = 10) -> str:
    """
    チャットの履歴から要約を生成し返す。

    要約する文字数は summary_length で指定する。
    ただし、GPTに文字数を指定して要約を生成させると、指定した文字数よりも多くなる場合がある。
    その場合、要約の文字数を summary_length に合わせ、最後に ... を追加する。
    :param initial_prompt: ユーザーの最初のプロンプト
    :param gpt_model: GPTモデルの名前
    :param summary_length: 要約する文字数
    :return: チャットの要約
    """

    # messages の先頭に要約の依頼を追加
    summary_request = {"role": "system",
                       "content": "あなたはユーザーの依頼を要約する役割を担います。"
                                  f"以下のユーザーの依頼を必ず全角{summary_length}文字以内で要約してください"}
    # GPTによる要約を取得
    messages = [summary_request, {"role": "user", "content": initial_prompt}]
    response = openai.ChatCompletion.create(model=gpt_model,
                                            messages=messages,
                                            max_tokens=summary_length)
    summary = response.choices[0].message.content

    # 要約を調整
    if len(summary) > summary_length:
        return summary[:summary_length] + "..."
    else:
        return summary


def chat_runner() -> tuple[list[dict], str]:
    """
    チャットを実行する
    :return: チャットのログと要約
    """

    # モデルの一覧を取得し、ユーザーに選択させる
    models = fetch_gpt_model_list()
    # エラーが発生した場合は終了
    if not models:
        exit()

    chose_model = choice_model(models)

    # ロールの設定
    give_role = give_role_to_system()

    # チャットを開始しログを生成する
    chat_log = generate_chat_log(chose_model, give_role)
    # ログが空の場合(いきなりexit()コマンド入力)は終了
    if not chat_log:
        exit()

    # チャットの最初のプロンプトを取得
    init_prompt = get_initial_prompt(chat_log)

    # 要約を生成
    generated_summary = ""
    if init_prompt:
        generated_summary = generate_summary(init_prompt, chose_model)

    return chat_log, generated_summary


# .envファイルからAPIキーを読み込む
load_dotenv()
openai.api_key = os.getenv("API_KEY")

if __name__ == "__main__":
    chat_runner()
