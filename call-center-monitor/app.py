"""
Call Center Monitoring Application for KM Next
建設資材商社向けコールセンターモニタリングアプリケーション
"""

import os
import re
import tempfile
from datetime import datetime

import streamlit as st
import google.generativeai as genai
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# System prompt for the AI model
SYSTEM_PROMPT = """
## あなたの役割
あなたは、建設資材商社（KMネクスト）のベテラン営業事務であり、スタッフの「良いところ」を見つけて伸ばす、人望の厚いスーパーバイザーです。
アップロードされた通話音声を聴取し、スタッフが「明日も頑張ろう」と思えるようなフィードバックを作成してください。

## ⚠️ 重要：言語と出力の絶対ルール
1. **完全日本語指定:** すべての出力を必ず「日本語」で行ってください。
2. **担当スタッフ名の特定:** ファイル名から取得した名前を最優先で採用してください。
3. **タイトル固定:** 出力1行目は必ず `# 【応対FB】{{スタッフ名}}_{{用件カテゴリ}}_{{今日の日付}}` としてください。

## 💡 評価のマインドセット（加点法）
* **目的:** スタッフのモチベーション向上と自信の醸成。
* **NG:** 咳、言い淀み、噛んでしまった等の「生理的なミス」や「些細なノイズ」は**完全に無視**してください。
* **OK:** 「お客様の要望を解決できたか」「安心感を与えられたか」という**成果**に焦点を当ててください。

## 📊 採点基準（貢献度スコア）
減点方式ではなく、「どれだけ良かったか」の加点方式で評価します。

1. **ヒアリング力**
    * [5]: 相手の要望を完全に把握し、スムーズに案内できた。
    * [4]: 必要な情報は概ね聞き取れている。
    * [3]: 一部聞き返しがあったが、業務に支障はない。

2. **スピード感**
    * [5]: お客様をお待たせした印象を与えない、素晴らしいテンポ。
    * [4]: 通常業務として問題ないスピード。
    * [3]: 少し時間がかかったが、許容範囲内。

3. **好感度・マナー（※重要）**
    * [5]: 明るい声、親身な対応で、お客様に安心感を与えた。
    * [4]: 失礼がなく、丁寧な対応ができている。
    * [3]: 事務的な対応。

## 分析プロセス
1. **【Good探しの旅】:** まず「この対応で良かった点」を3つ以上探す。（例：復唱確認した、在庫を即答した、声が明るかった等）
2. **【Next Stepの選定】:** 否定的な指摘は避け、「さらにプロになるためのヒント（＋αの提案など）」を1つだけ選ぶ。

## 出力フォーマット
# 【応対FB】{{スタッフ名}}_{{用件カテゴリ}}_{{今日の日付}}

---
### 🛠️ モニタリング・フィードバックシート

**■ 基本情報**
* **担当スタッフ名:** {{スタッフ名}}
* **会話の趣旨:** {{用件カテゴリ}}
* **キーワード:** （音声で聞こえたもののみ）

**■ 📊 パフォーマンス・スコア**
* **ヒアリング力:** [ 4 ] （※5段階評価）
* **スピード感:** [ 4 ]
* **好感度・マナー:** [ 4 ]
* **総合評価ランク:** [ A ] （S/A/B）

**■ 案件概要**
（要約）

**■ 素晴らしいポイント（Good Points）** 🌟ここが現場の助けになりました！
* **[項目]:** （〇〇さんは… ※具体的に褒める）
* **[項目]:** （〇〇さんの対応により、お客様は…）
* **[項目]:** （些細な気遣いも見逃さずに褒める）

**■ さらなるレベルアップへ（Next Step）** 🚀ここを磨けば完璧です
* （※注意や叱責はNG。「こうするともっと良くなる」という未来志向のヒントを1点のみ）

**■ SVからのエール**
（〇〇さんの強みに触れながら、温かい励ましのメッセージ）
---
"""

# Supported audio file extensions
SUPPORTED_EXTENSIONS = ["mp3", "mp4", "m4a", "wav"]


def extract_staff_name(filename: str) -> str:
    """
    Extract staff name from filename.
    Examples:
        "Tanaka_20251225.mp3" -> "Tanaka"
        "田中_在庫確認.mp3" -> "田中"
        "山田太郎_注文対応_20251225.wav" -> "山田太郎"
    """
    # Remove file extension
    name_without_ext = os.path.splitext(filename)[0]

    # Split by underscore and take the first part as staff name
    parts = name_without_ext.split("_")
    if parts:
        return parts[0]

    return name_without_ext


def extract_category_from_response(response_text: str) -> str:
    """
    Extract category from the AI response.
    Looks for pattern like 【応対FB】スタッフ名_カテゴリ_日付
    """
    # Try to extract category from the title line
    pattern = r"【応対FB】[^_]+_([^_]+)_"
    match = re.search(pattern, response_text)
    if match:
        return match.group(1)
    return "対応記録"


def get_mime_type(filename: str) -> str:
    """Get MIME type based on file extension."""
    ext = filename.lower().split(".")[-1]
    mime_types = {
        "mp3": "audio/mp3",
        "mp4": "audio/mp4",
        "m4a": "audio/mp4",
        "wav": "audio/wav",
    }
    return mime_types.get(ext, "audio/mpeg")


def analyze_audio_with_gemini(audio_file, staff_name: str) -> str:
    """
    Upload audio to Gemini API and analyze it.
    """
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise ValueError("GOOGLE_API_KEY not found in environment variables")

    genai.configure(api_key=api_key)

    # Save uploaded file to temporary location
    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{audio_file.name.split('.')[-1]}") as tmp_file:
        tmp_file.write(audio_file.getvalue())
        tmp_path = tmp_file.name

    try:
        # Upload file to Gemini
        uploaded_file = genai.upload_file(tmp_path, mime_type=get_mime_type(audio_file.name))

        # Wait for file to be processed
        import time
        while uploaded_file.state.name == "PROCESSING":
            time.sleep(1)
            uploaded_file = genai.get_file(uploaded_file.name)

        if uploaded_file.state.name == "FAILED":
            raise ValueError("Audio file processing failed")

        # Create the model
        model = genai.GenerativeModel("gemini-1.5-flash")

        # Create prompt with staff name and today's date
        today = datetime.now().strftime("%Y年%m月%d日")
        user_prompt = f"""
担当スタッフ名: {staff_name}
今日の日付: {today}

上記の情報を使用して、アップロードされた音声ファイルを分析し、フィードバックを作成してください。
"""

        # Generate response
        response = model.generate_content(
            [SYSTEM_PROMPT, uploaded_file, user_prompt]
        )

        # Clean up uploaded file from Gemini
        genai.delete_file(uploaded_file.name)

        return response.text

    finally:
        # Clean up temporary file
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def main():
    """Main Streamlit application."""
    st.set_page_config(
        page_title="コールセンターモニタリング | KM Next",
        page_icon="📞",
        layout="wide"
    )

    st.title("📞 コールセンターモニタリングアプリ")
    st.markdown("**KM Next** - 建設資材商社向け通話品質フィードバックシステム")

    # Check for API key
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        st.error("⚠️ GOOGLE_API_KEY が設定されていません。`.env` ファイルにAPIキーを設定してください。")
        st.code("GOOGLE_API_KEY=your_api_key_here", language="text")
        st.stop()

    st.divider()

    # File uploader
    st.subheader("📁 音声ファイルをアップロード")
    st.markdown("""
    対応形式: **MP3**, **MP4**, **M4A**, **WAV**

    ファイル名の形式: `スタッフ名_その他情報.拡張子`
    例: `田中_在庫確認_20251225.mp3`
    """)

    uploaded_file = st.file_uploader(
        "音声ファイルを選択",
        type=SUPPORTED_EXTENSIONS,
        help="通話録音ファイルをアップロードしてください"
    )

    if uploaded_file is not None:
        # Extract staff name from filename
        staff_name = extract_staff_name(uploaded_file.name)

        st.info(f"📋 **検出されたスタッフ名:** {staff_name}")

        # Allow user to override staff name
        staff_name_input = st.text_input(
            "スタッフ名を修正（必要な場合）",
            value=staff_name,
            help="ファイル名から自動検出されたスタッフ名を修正できます"
        )

        if staff_name_input:
            staff_name = staff_name_input

        st.divider()

        # Analyze button
        if st.button("🔍 音声を分析する", type="primary", use_container_width=True):
            with st.spinner("音声を分析中...（しばらくお待ちください）"):
                try:
                    feedback_text = analyze_audio_with_gemini(uploaded_file, staff_name)

                    # Store result in session state
                    st.session_state["feedback_text"] = feedback_text
                    st.session_state["staff_name"] = staff_name
                    st.session_state["category"] = extract_category_from_response(feedback_text)

                except Exception as e:
                    st.error(f"エラーが発生しました: {str(e)}")
                    st.stop()

        # Display results if available
        if "feedback_text" in st.session_state:
            st.divider()
            st.subheader("📝 フィードバック結果")

            # Display the feedback
            st.markdown(st.session_state["feedback_text"])

            st.divider()

            # Download button
            today = datetime.now().strftime("%Y%m%d")
            download_filename = f"【応対FB】{st.session_state['staff_name']}_{st.session_state['category']}_{today}.txt"

            st.download_button(
                label="📥 テキストファイルをダウンロード",
                data=st.session_state["feedback_text"],
                file_name=download_filename,
                mime="text/plain",
                use_container_width=True
            )

    # Footer
    st.divider()
    st.markdown(
        "<div style='text-align: center; color: gray;'>"
        "© 2025 KM Next - Call Center Monitoring System"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
