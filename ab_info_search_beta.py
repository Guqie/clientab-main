import streamlit as st
import os
import re
from pathvalidate import sanitize_filename
import base64
import threading
from ab_time import now_in_filename, scheduled_run
from ab_utils import clean_yesterday_files
from ab_tools import get_prompt, get_response_format, get_tools, chat, ai_dict, openrouter, excellence2

st.session_state["ai"] = st.query_params.get("ai", st.session_state.get("ai", "GPT for text chat"))
st.session_state["chat_history"] = st.session_state.get("chat_history", [{"role": "assistant", "content": "请上传Excel或CSV文件，或者提供你搜索到的文章的URLs"}])
st.session_state["chat_history_editable"] = st.session_state.get("chat_history_editable", "")
st.session_state["is_chat_history_edited"] = st.session_state.get("is_chat_history_edited", False)
st.session_state["files_info"] = st.session_state.get("files_info", [])
st.session_state["doc_content"] = st.session_state.get("doc_content", "")


def select():
    st.query_params.update({"ai": st.session_state["ai"]})


def update_chat_history():
    pattern = r"(User:\n|AI \([^)]+\):\n)"
    segments = ["User:\n"] + [segment for segment in re.split(pattern, st.session_state["chat_history_editable"]) if segment.strip()]
    return [{"role": "user" if segments[i] == "User:\n" else "assistant", "content": segments[i + 1].strip()} for i in range(len(segments) - 1) if re.match(pattern, segments[i]) and not re.match(pattern, segments[i + 1])]


def sync_chat_history_editable():
    return "\n\n".join([f"User:\n{message['content']}" if message["role"] == "user" else f"AI ({st.session_state['ai']}):\n{message['content']}" for message in st.session_state["chat_history"]])


def append_user_message(content):
    st.session_state["chat_history"].append({"role": "user", "content": content})
    st.session_state["chat_history_editable"] = sync_chat_history_editable()


def append_assistant_message(content):
    st.session_state["chat_history"].append({"role": "assistant", "content": content})
    st.session_state["chat_history_editable"] = sync_chat_history_editable()
    st.rerun()


def is_chat_history_edited():
    st.session_state["is_chat_history_edited"] = True


def get_file_paths(file):
    file_path = f"uploaded-files/{sanitize_filename(os.path.splitext(file.name)[0])} {now_in_filename()}{os.path.splitext(file.name)[1]}"
    with open(file_path, "wb") as f:
        f.write(file.getbuffer())
    return file_path


def update_doc_content(contents, cutoff_length):
    doc_content = f"{st.session_state.get('doc_content')}\n\n{'\n\n'.join(contents[key] for key in sorted(contents))}"
    st.session_state["doc_content"] = doc_content[-cutoff_length:]


def append_txt_path(contents):
    txt_path = f"temp-data/{now_in_filename()}.txt"
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(str(contents))
    append_user_message(f"The path of the TXT file containing the doc_content: {txt_path}")


def get_new_files(files):
    return [file for file in files if (file.name, file.size) not in set(st.session_state["files_info"])]


def within_length_limit(user_message, results=None):
    ai = st.session_state["ai"]
    max_length = ai_dict[ai]["max_length"]
    current_length = len(st.session_state["chat_history_editable"]) + len(user_message) + (len(results) if results else 0) + 1000
    if current_length > max_length:
        st.warning("Please shorten the chat history to continue the thread.")
        return False
    if current_length > max_length * 0.9:
        st.warning(f"The thread is approaching the length limit ({current_length}/{max_length}). Consider cutting the chat history or starting a new thread.")
    return True


def text_chat(user_message, results=None):
    ai = st.session_state["ai"]
    llms = ai_dict[ai]["llms"]
    system_message = get_prompt(ai_dict[ai]["system_message"])
    response_format = get_response_format(ai_dict[ai]["response_format"])
    tools = get_tools(ai_dict[ai]["tools"])
    messages = [{"role": "system", "content": system_message}] + st.session_state["chat_history"] + ([{"role": "assistant", "content": f"{results}"}] if results else []) + [{"role": "user", "content": user_message}]
    return chat(llms, messages, response_format=response_format, tools=tools)


def images_chat(user_message, image_paths):
    ai = st.session_state["ai"]
    llms = ai_dict[ai]["llms"]
    messages = st.session_state["chat_history"] + [{"role": "user", "content": [{"type": "text", "text": user_message}, *[{"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64.b64encode(open(image_path, 'rb').read()).decode('utf-8')}"}} for image_path in image_paths]]}]
    return chat(llms, messages)


st.set_page_config(layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.markdown("<h1 style='text-align: center; font-size: 18px;'>安邦智库信息二部内测</h1>", unsafe_allow_html=True)
    ais = list(ai_dict.keys())[:1]
    captions = [ai_dict[ai]["intro"] for ai in ais]
    st.radio("AI options", ais, key="ai", on_change=select, captions=captions, label_visibility="collapsed")

chat_view_tab, editable_view_tab = st.tabs(["Chat view", "Editable view"])

user_message = (st.chat_input("Input a message") or "").strip()

file_uploader_column, download_button_column = st.columns([3, 1])

with file_uploader_column:
    file = st.file_uploader("Upload a CSV or Excel file", type=["csv", "xlsx"], accept_multiple_files=False, label_visibility="collapsed")

with download_button_column:
    if st.session_state.get("is_chat_history_edited"):
        st.download_button("Download chat history", f"Chat history after editing:\n{st.session_state['chat_history_editable']}\n\nChat history before editing:\n{sync_chat_history_editable()}", f"Chat history {now_in_filename()}.txt", "text/plain", use_container_width=True)
    else:
        st.download_button("Download chat history", st.session_state["chat_history_editable"], f"Chat history {now_in_filename()}.txt", "text/plain", use_container_width=True)

if user_message:
    if st.session_state.get("is_chat_history_edited"):
        st.session_state["chat_history"] = update_chat_history()
        st.session_state["is_chat_history_edited"] = False

    ai = st.session_state["ai"]
    category = ai_dict[ai]["category"]

    try:
        if category == "function_calling":
            if file:
                file_path = get_file_paths(file)
                if file_path.endswith(".xlsx"):
                    append_user_message(f"The path of the Excel file: {file_path}")
                else:
                    append_user_message(f"The path of the CSV file: {file_path}")

            if within_length_limit(user_message):
                with st.spinner("Let me think... 🧠"):
                    results = text_chat(user_message)
                    append_user_message(user_message)
                    if results:
                        if results.startswith("The following dictionary contains the results:\n"):
                            user_message = get_prompt("reply_with_results")
                            with st.spinner("Further processing... 🧠"):
                                results = text_chat(user_message, results)
                                if results:
                                    append_assistant_message(results)
                        else:
                            append_assistant_message(results)

    except Exception as e:
        st.warning(f"An exception occurred: {e}")

with chat_view_tab:
    with st.container(height=462, border=True):
        for message in st.session_state["chat_history"]:
            role = message["role"]
            content = message["content"]
            if role == "user":
                with st.chat_message("user"):
                    st.write(f"User:\n{content}")
            else:
                ai = st.session_state["ai"]
                with st.chat_message("assistant"):
                    st.write(f"AI ({ai}):\n{content}")

with editable_view_tab:
    st.text_area("Editable view", height=460, key="chat_history_editable", on_change=is_chat_history_edited, label_visibility="collapsed")

st.components.v1.html(
    """
    <script>
    let lastContent = '';

    function getLast30Chars(str) {
        return str.slice(-30);
    }

    function checkAndScroll() {
        const textArea = window.parent.document.querySelector('textarea[aria-label="Editable view"]');
        if (!textArea) return;

        const currentContent = textArea.value;
        const currentLast30 = getLast30Chars(currentContent);
        const lastLast30 = getLast30Chars(lastContent);

        if (currentLast30 !== lastLast30) {
            textArea.scrollTop = textArea.scrollHeight;
            lastContent = currentContent;
        }
    }

    // 每500毫秒检查一次
    setInterval(checkAndScroll, 500);
    </script>
    """,
    height=0
)

cleanup_thread = threading.Thread(target=scheduled_run, args=(10, 0, clean_yesterday_files), daemon=True)
cleanup_thread.start()
