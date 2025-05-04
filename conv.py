import zipfile
import os
import shutil
import tempfile
import pandas as pd
import subprocess


def get_label(index, style="1"):
    if style == "1":
        return f"{index}."
    elif style == "a":
        return chr(96 + index) + "."   # 1→a., 2→b. ...
    elif style == "A":
        return chr(64 + index) + "."   # 1→A., 2→B. ...
    elif style == "ア":
        return "アイウエオカキクケコサシス"[index-1] + "."
    else:
        return f"{index}."


def convert_file_to_html(filepath):
    """
    .docx または .tex ファイルを自動判定して HTMLに変換する
    """
    if not os.path.exists(filepath):
        return ""

    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".docx":
        input_format = "docx"
    elif ext == ".tex":
        input_format = "latex"
    else:
        return ""

    with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp_html:
        html_path = tmp_html.name

    try:
        subprocess.run([
            "pandoc", filepath,
            "-f", input_format,
            "-t", "html",
            "--mathjax",
            "-o", html_path
        ], check=True)

        with open(html_path, encoding="utf-8") as f:
            html_content = f.read()
    finally:
        os.remove(html_path)

    return html_content


def make_image_tag(image_file):
    if not image_file or pd.isna(image_file):
        return ""
    ext = image_file.split('.')[-1].lower()
    mime = {
        "png": "image/png", "jpg": "image/jpeg",
        "jpeg": "image/jpeg", "gif": "image/gif"
    }.get(ext, "image/png")
    return f'''
    <material>
      <matimage imagtype="{mime}" uri="images/{image_file}" label="{image_file}" />
    </material>
    '''

# --- 問題形式別の関数（cloze, radio, checkbox, text） ---

def convert_to_cloze_dropdown(question_name, answer, question, options, label_style, is_html="plain"):
    answers = answer.split("!#!")
    num_blanks = len(answers)

    xml = f'''
    <item title="{question_name}" ident="{question_name}">
      <itemmetadata>
        <qtimetadata>
          <qtimetadatafield>
            <fieldlabel>question_type</fieldlabel>
            <fieldentry>multiple_dropdowns_question</fieldentry>
          </qtimetadatafield>
        </qtimetadata>
      </itemmetadata>
      <presentation>
        <material>
          <mattext texttype="text/html">
              <![CDATA[\n{question}<br/>
    '''
    
    placeholders = " ".join([f"{get_label(j+1, label_style)} [q{j+1}]" for j in range(num_blanks)])
    xml += f"{placeholders}<br/>]]>\n</mattext>\n</material>\n"

    for j in range(num_blanks):
        qid = f"q{j+1}"
        xml += f'''<response_lid ident="response_{qid}">
          <render_choice>
        '''
        for i, opt in enumerate(options):
            if pd.notna(opt):
                xml += f'''
            <response_label ident="opt{i+1}">
              <material><mattext>{opt}</mattext></material>
            </response_label>
                '''
        xml += '''
          </render_choice>
        </response_lid>
        '''

    xml += '</presentation>\n<resprocessing>\n<outcomes>\n'
    xml += '<decvar maxvalue="100" minvalue="0" varname="SCORE" vartype="Decimal"/>\n</outcomes>\n'

    point_per_blank = 100 / num_blanks
    for j, correct_ans in enumerate(answers):
        qid = f"q{j+1}"
        correct_id = int(correct_ans) if correct_ans.isdigit() else None
        xml += f'''
        <respcondition>
          <conditionvar>
            <varequal respident="response_{qid}">opt{correct_id}</varequal>
          </conditionvar>
          <setvar varname="SCORE" action="Add">{point_per_blank}</setvar>
        </respcondition>
        '''
        break

    xml += '</resprocessing>\n'
    return xml


def make_single_choice_item(question_name, question, options, correct_answer, is_html="plain"):
    if is_html == "html":
        question = "<![CDATA[" + question + "]]>"
    xml = f'''
    <item title="{question_name}" ident="{question_name}">
      <itemmetadata>
        <qtimetadata>
          <qtimetadatafield>
            <fieldlabel>question_type</fieldlabel>
            <fieldentry>multiple_choice_question</fieldentry>
          </qtimetadatafield>
        </qtimetadata>
      </itemmetadata>
      <presentation>
        <material>
          <mattext texttype="text/{is_html}">{question}</mattext>
        </material>
        <response_lid ident="response1" rcardinality="Single">
          <render_choice>
    '''
    correct_id = int(correct_answer) if correct_answer.isdigit() else None
    for i, opt in enumerate(options):
        if pd.notna(opt):
            opt_id = f"opt{i+1}"
            xml += f'''
            <response_label ident="{opt_id}">
              <material><mattext>{opt}</mattext></material>
            </response_label>
            '''
       

    xml += '''
          </render_choice>
        </response_lid>
      </presentation>
      <resprocessing>
        <outcomes>
          <decvar maxvalue="100" minvalue="0" varname="SCORE" vartype="Decimal"/>
        </outcomes>
    '''
    if correct_id:
        xml += f'''
        <respcondition>
          <conditionvar>
            <varequal respident="response1">opt{correct_id}</varequal>
          </conditionvar>
          <setvar varname="SCORE" action="Set">100</setvar>
        </respcondition>
        '''

    xml += '</resprocessing>\n'
    return xml


def make_multiple_choice_item(question_name, question, options, correct_answers, is_html="plain"):
    if is_html == "html":
        question = "<![CDATA[" + question + "]]>"
    xml = f'''
    <item title="{question_name}" ident="{question_name}">
      <itemmetadata>
        <qtimetadata>
          <qtimetadatafield>
            <fieldlabel>question_type</fieldlabel>
            <fieldentry>multiple_answers_question</fieldentry>
          </qtimetadatafield>
        </qtimetadata>
      </itemmetadata>
      <presentation>
        <material>
          <mattext texttype="text/{is_html}">{question}</mattext>
        </material>
        <response_lid ident="response1" rcardinality="Multiple">
          <render_choice>
    '''
    correct_ids = []
    for i, opt in enumerate(options):
        if pd.notna(opt):
            opt_id = f"opt{i+1}"
            xml += f'''
            <response_label ident="{opt_id}">
              <material><mattext>{opt}</mattext></material>
            </response_label>
            '''
            if str(opt).strip() in [s.strip() for s in correct_answers]:
                correct_ids.append(opt_id)

    xml += '''
          </render_choice>
        </response_lid>
      </presentation>
      <resprocessing>
        <outcomes>
          <decvar maxvalue="100" minvalue="0" varname="SCORE" vartype="Decimal"/>
        </outcomes>
    '''

    if correct_ids:
        point_per_ans = 100 / len(correct_ids)
        for cid in correct_ids:
            xml += f'''
        <respcondition>
          <conditionvar>
            <varequal respident="response1">{cid}</varequal>
          </conditionvar>
          <setvar varname="SCORE" action="Add">{point_per_ans}</setvar>
        </respcondition>
            '''

    xml += '</resprocessing>\n'
    return xml


def convert_to_cloze_text_input(question_name, answer, question, label_style, is_html="plain"):
    """
    WebClassのwordinput形式をCanvasの自由入力（非ドロップダウン）空欄補充問題に変換。
    - 各空欄にrender_fibを使い、受験者が自由入力
    - 正解は完全一致によって判定（大文字小文字無視）
    """
    answers = answer.split("!#!")
    num_blanks = len(answers)

    xml = f'''
    <item title="{question_name}" ident="{question_name}">
      <itemmetadata>
        <qtimetadata>
          <qtimetadatafield>
            <fieldlabel>question_type</fieldlabel>
            <fieldentry>fill_in_multiple_blanks_question</fieldentry>
          </qtimetadatafield>
        </qtimetadata>
      </itemmetadata>
      <presentation>
        <material>
          <mattext texttype="text/html">
            <![CDATA[
              {question}<br/>
    '''
    

    placeholders = " ".join([f"{get_label(j+1, label_style)} [q{j+1}]" for j in range(num_blanks)])
    xml += f"{placeholders}<br/>]]>\n</mattext>\n</material>\n"

    for j in range(num_blanks):
        qid = f"q{j+1}"
        xml += f'''
        <response_str ident="response_{qid}" rcardinality="Single">
          <render_fib fibtype="String" prompt="Box" maxchars="60"/>
        </response_str>
        '''

    xml += '</presentation>\n<resprocessing>\n<outcomes>\n'
    xml += '<decvar maxvalue="100" minvalue="0" varname="SCORE" vartype="Decimal"/>\n</outcomes>\n'

    point_per_blank = 100 / num_blanks
    for j, correct_ans in enumerate(answers):
        qid = f"q{j+1}"
        xml += f'''
        <respcondition>
          <conditionvar>
            <varequal respident="response_{qid}" case="no">{correct_ans.strip()}</varequal>
          </conditionvar>
          <setvar varname="SCORE" action="Add">{point_per_blank}</setvar>
        </respcondition>
        '''

    xml += '</resprocessing>\n '
    return xml


def find_main_csv(root_dir):
    csv_paths = []
    for root, dirs, files in os.walk(root_dir):
        for f in files:
            if f.lower().endswith(".csv") and not f.startswith("._"):
                full_path = os.path.join(root, f)
                csv_paths.append(full_path)

    if not csv_paths:
        raise FileNotFoundError("CSVファイルが見つかりませんでした")

    # "list.csv" があればそれを使う
    for path in csv_paths:
        if os.path.basename(path).lower() == "list.csv":
            return path

    # なければ最初のものを使う（または別ルールでソート）
    return csv_paths[0]




# --- 統合関数 ---
def convert_webclass_zip_to_qti(zip_path, output_zip_path, label_style="1", assessment_title="WebClass Import"):
    extract_dir = "webclass_temp"
    image_dir = os.path.join(extract_dir, "images")
    os.makedirs(image_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    csv_path = find_main_csv(extract_dir)
    df = pd.read_csv(csv_path, encoding="shift_jis")


    xml_header = '''<?xml version="1.0" encoding="UTF-8"?>
<questestinterop>
  <assessment ident="webclass_assess" title="WebClass Import">
    <section ident="root_section">
'''
    xml_footer = '''
    </section>
  </assessment>
</questestinterop>
'''

    items_xml = ""
    for idx, row in df.iterrows():
        style = str(row.get("style", "")).strip()
        answer = str(row.get("answer", "")).strip()
        qfile = os.path.join(extract_dir, str(row.get("question_file", "")).strip())
        dfile = os.path.join(extract_dir, str(row.get("description_file", "")).strip())
        image = str(row.get("image", "")).strip() if "image" in row else None
        options = [row.get(f"option{i}", None) for i in range(1, 27)]
        question_id = f"q_{idx+1}"

        qfile_raw = str(row.get("question_file", "")).strip()
        qfile = os.path.join(extract_dir, qfile_raw) if qfile_raw else None

        if qfile and os.path.exists(qfile):
            question_text = convert_file_to_html(qfile)
            is_html = "html"
        else:
            question_text = str(row.get("question", "")).strip()
            is_html = "plain"  # テキスト入力なら text/plain 扱い


        # === 解説の取得とHTML判定 ===
        if dfile and os.path.exists(dfile):
            description_text = "<![CDATA["+ convert_file_to_html(dfile) + "]]>"
            is_desc_html = "html"
        else:
            description_text = str(row.get("description", "")).strip()
            is_desc_html = "plain"

        img_tag = make_image_tag(image)

        feedback= f'''
            <itemfeedback ident="general_fb">
            <flow_mat>
                <material>
                <mattext texttype="text/{is_desc_html}">
                    {description_text}
                </mattext>
                </material>
            </flow_mat>
            </itemfeedback>
            </item>
            '''

        if style == "radio":
            items_xml += make_single_choice_item(question_id, question_text + img_tag, options, answer, is_html) + feedback
            
        elif style == "checkbox":
            corrects = answer.split("!#!")
            items_xml += make_multiple_choice_item(question_id, question_text + img_tag, options, corrects, is_html) + feedback
        elif style == "dropdown":
            items_xml += convert_to_cloze_dropdown(question_id,  answer, question_text + img_tag, options, label_style, is_html) + feedback
        elif style == "wordinput":
            items_xml += convert_to_cloze_text_input(question_id, answer, question_text + img_tag, label_style, is_html) + feedback
        # 必要に応じて他形式を追加（例：line, text）

        
    xml_path = "quiz.xml"
    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(xml_header + items_xml + xml_footer)

    meta_path = "assessment_meta.xml"
    with open(meta_path, "w", encoding="utf-8") as f:
        f.write(f'''<?xml version="1.0" encoding="UTF-8"?>
<quiz identifier="{assessment_title}"
xmlns="http://canvas.instructure.com/xsd/cccv1p0"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
xsi:schemaLocation="http://canvas.instructure.com/xsd/cccv1p0 https://canvas.instructure.com/xsd/cccv1p0.xsd">
<title>{assessment_title}</title>
<scoring_policy>keep_highest</scoring_policy>
</quiz>
''')

    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(xml_path, arcname="quiz.xml")
        zipf.write(meta_path, arcname="assessment_meta.xml")
        for f in os.listdir(extract_dir):
            if f.lower().endswith((".png", ".jpg", ".jpeg", ".gif")):
                zipf.write(os.path.join(extract_dir, f), arcname=f"images/{f}")

    shutil.rmtree(extract_dir)
    os.remove(xml_path)
    os.remove(meta_path)


import streamlit as st
import os

st.set_page_config(page_title="WebClass → Canvas QTI 変換", layout="centered")
st.title("WebClass → Canvas QTI 変換ツール")
assessment_title = st.text_input("クイズのタイトル", "WebClass Import", key="assessment_title")

st.text_input("変換する問題のラベル形式", "1", key="label_style")

uploaded_file = st.file_uploader("WebClass形式のZIPファイルをアップロード", type="zip")

if uploaded_file:
    input_zip_path = "temp_input.zip"
    with open(input_zip_path, "wb") as f:
        f.write(uploaded_file.read())

    output_zip_path = "canvas_qti.zip"

    with st.spinner("変換中...しばらくお待ちください"):
        convert_webclass_zip_to_qti(input_zip_path, output_zip_path, label_style=st.session_state.label_style, assessment_title=assessment_title)

    st.success("変換完了！Canvas用QTI ZIPファイルを以下からダウンロードできます。")
    with open(output_zip_path, "rb") as f:
        st.download_button("QTI ZIPをダウンロード", f, file_name="canvas_qti.zip")
