import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import time


# --- 页面配置 ---
st.set_page_config(page_title="AI Excel 超级助手", page_icon="🚀", layout="wide")
st.title("🚀 AI Excel 超级助手")
# --- 🎨 CSS 样式美化区 ---
st.markdown("""
<style>
    /* 1. 隐藏默认的菜单和页脚 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* 2. 全局字体优化 */
    html, body, [class*="css"] {
        font-family: 'PingFang SC', 'Microsoft YaHei', sans-serif;
    }

    /* 3. 按钮美化 (渐变色+圆角) */
    div.stButton > button {
        background: linear-gradient(45deg, #4b6cb7, #182848);
        color: white;
        border: none;
        border-radius: 20px;
        padding: 10px 24px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    div.stButton > button:hover {
        transform: scale(1.05);
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }

    /* 4. 侧边栏美化 */
    section[data-testid="stSidebar"] {
        background-color: #f8f9fa;
        border-right: 1px solid #e0e0e0;
    }
    
    /* 5. 表格区域加个卡片阴影效果 */
    div[data-testid="stDataFrame"] {
        background: white;
        padding: 10px;
        border-radius: 10px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)
# -------------------------
# --- 1. 获取 API Key ---
api_key = st.secrets.get("GOOGLE_API_KEY")
if not api_key:
    st.error("请在 Streamlit 后台 Secrets 设置 GOOGLE_API_KEY")
    st.stop()
genai.configure(api_key=api_key)

# --- 2. 侧边栏 ---
with st.sidebar:
    st.header("📂 文件操作区")
    uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])
    
    if st.button("🔄 重置/清除缓存"):
        st.cache_data.clear()
        if 'df' in st.session_state:
            del st.session_state['df']
        st.rerun()

# --- 3. 核心逻辑 ---
if uploaded_file:
    # 读取文件
    if 'df' not in st.session_state:
        try:
            st.session_state.df = pd.read_excel(uploaded_file)
            st.toast("✅ 文件上传成功！")
        except Exception as e:
            st.error(f"文件读取失败: {e}")
            st.stop()
    
    df = st.session_state.df

    # --- 🔥 关键修复：类型安全检查 🔥 ---
    # 如果 df 变质了（不是表格了），就强制恢复
    if not isinstance(df, pd.DataFrame):
        st.warning("⚠️ 检测到数据格式异常（可能 AI 把表格变成了一个值），已自动重置数据。")
        st.session_state.df = pd.read_excel(uploaded_file)
        df = st.session_state.df
        st.rerun()

    # --- 数据展示区 ---
    st.subheader("📊 数据全览")
    st.dataframe(df, use_container_width=True, height=400)
    
    # 这里加了保护，确保 df 真的是个表格才读取 shape
    if hasattr(df, 'shape'):
        st.caption(f"当前共 {df.shape[0]} 行, {df.shape[1]} 列数据")

    # --- 聊天输入 ---
    st.divider()
    user_query = st.chat_input("💡 请下达指令，例如：'把销售额大于500的标红'...")

    if user_query:
        with st.status("🤖 AI 正在干活...", expanded=True) as status:
            st.write("1️⃣ 正在思考 Python 解决方案...")
            
            try:
                # 模型加载
                try:
                    model = genai.GenerativeModel('gemini-2.5-flash')
                except:
                    model = genai.GenerativeModel('gemini-pro')

                # 获取列信息（防止报错）
                dtypes_info = df.dtypes.to_string() if hasattr(df, 'dtypes') else "无"

                # Prompt
                prompt = f"""
                你是一个 Python Pandas 高级专家。
                
                【当前数据情况】
                变量名: `df`
                列信息:
                {dtypes_info}
                
                【用户任务】
                "{user_query}"
                
                【绝对禁令】
                1. 严禁将 `df` 赋值为非 DataFrame 对象（如数字、列表、Series）。
                2. 如果用户要求计算（如“求和”、“计数”），请不要修改 `df`，而是新建变量并使用 `print()` 输出结果。
                3. 只有在需要修改表格结构/内容时，才对 `df` 进行赋值。
                
                【输出要求】
                只输出 Python 代码，不要 ```python 标记。
                """
                
                # 生成
                response = model.generate_content(prompt)
                code = response.text.replace("```python", "").replace("```", "").strip()
                st.code(code, language='python')
                
                # 执行
                st.write("2️⃣ 正在执行...")
                # 捕获 print 输出
                from io import StringIO
                import sys
                captured_output = StringIO()
                sys.stdout = captured_output
                
                local_vars = {'df': df, 'pd': pd, 'st': st}
                exec(code, globals(), local_vars)
                
                # 恢复标准输出
                sys.stdout = sys.__stdout__
                output_str = captured_output.getvalue()
                
                if output_str:
                    st.info(f"🖨️ 计算结果:\n{output_str}")

                # 检查执行后的 df 是否还是个表格
                new_df = local_vars.get('df')
                if isinstance(new_df, pd.DataFrame):
                    st.session_state.df = new_df
                    status.update(label="✅ 表格已修改！", state="complete", expanded=False)
                    time.sleep(1)
                    st.rerun()
                else:
                    status.update(label="✅ 计算完成 (表格未变动)", state="complete", expanded=False)
                
            except Exception as e:
                status.update(label="❌ 执行失败", state="error")
                st.error(f"出错: {e}")

    # --- 下载 ---
    if not df.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.df.to_excel(writer, index=False)
        st.download_button("📥 下载结果", output.getvalue(), "AI_Result.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("👈 请先上传 Excel 文件")
