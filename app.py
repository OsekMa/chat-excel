import streamlit as st
import pandas as pd
import google.generativeai as genai
import io

# 页面标题
st.set_page_config(page_title="Excel AI 助手", page_icon="🤖")
st.title("🤖 AI Excel 对话助手")

# 获取 API Key (稍后在网页后台填，不要写在代码里)
api_key = st.secrets.get("GOOGLE_API_KEY")

if not api_key:
    st.error("请在后台设置 GOOGLE_API_KEY")
    st.stop()

genai.configure(api_key=api_key)

# 侧边栏：上传文件
with st.sidebar:
    uploaded_file = st.file_uploader("第一步：上传 Excel 文件", type=["xlsx"])

# 核心逻辑
if uploaded_file:
    # 读取 Excel
    if 'df' not in st.session_state:
        st.session_state.df = pd.read_excel(uploaded_file)
    
    df = st.session_state.df
    
    # 展示前几行
    st.subheader("当前数据预览：")
    st.dataframe(df.head())

    # 聊天输入框
    user_query = st.chat_input("输入指令，例如：把销售额大于100的行标红...")

    if user_query:
        st.write(f"🗣️ **你的指令:** {user_query}")
        
        # 调用 AI
        try:
            model = genai.GenerativeModel('gemini-2.5-flash') # 或者 gemini-pro
            
            prompt = f"""
            你是一个 Python pandas 专家。变量名必须用 `df`。
            列名: {list(df.columns)}
            前3行数据: {df.head(3).to_markdown()}
            用户需求: "{user_query}"
            请只生成 Python 代码，不要解释，不要 ```python 标记。
            必须直接修改 `df` 变量。
            """
            
            response = model.generate_content(prompt)
            code = response.text.replace("```python", "").replace("```", "").strip()
            
            st.code(code, language='python') # 展示生成的代码
            
            # 执行代码
            local_vars = {'df': df, 'pd': pd}
            exec(code, globals(), local_vars)
            st.session_state.df = local_vars['df'] # 更新状态
            
            st.success("✅ 执行成功！")
            st.rerun() # 刷新页面显示新数据
            
        except Exception as e:
            st.error(f"❌ 出错: {e}")

    # 下载按钮
    st.divider()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.df.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 下载修改后的 Excel",
        data=output.getvalue(),
        file_name="modified_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👈 请先在左侧上传 Excel 文件")
