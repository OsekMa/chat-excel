import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import time
import matplotlib.pyplot as plt
import matplotlib

# --- 解决 Matplotlib 中文乱码和后端问题 ---
matplotlib.use('Agg') # 这是一个非交互式后端，适合服务器环境
# 尝试设置中文字体，Streamlit Cloud 默认没有中文字体，通常会回退到 sans-serif
# 如果需要完美中文支持，建议让 AI 使用 Streamlit 原生图表 (st.bar_chart)
plt.rcParams['font.sans-serif'] = ['sans-serif'] 
plt.rcParams['axes.unicode_minus'] = False 

# --- 页面配置 ---
st.set_page_config(page_title="AI Excel 超级助手", page_icon="🚀", layout="wide")
st.title("🚀 AI Excel 超级助手")

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

    # 类型安全检查
    if not isinstance(df, pd.DataFrame):
        st.warning("⚠️ 数据异常，已自动重置。")
        st.session_state.df = pd.read_excel(uploaded_file)
        df = st.session_state.df
        st.rerun()

    # --- 数据展示区 ---
    st.subheader("📊 数据全览")
    st.dataframe(df, use_container_width=True, height=400)
    
    if hasattr(df, 'shape'):
        st.caption(f"当前共 {df.shape[0]} 行, {df.shape[1]} 列数据")

    # --- 聊天输入 ---
    st.divider()
    user_query = st.chat_input("💡 请输入指令，例如：'画一个柱状图展示各分类的数量'...")

    if user_query:
        with st.status("🤖 AI 正在干活...", expanded=True) as status:
            st.write("1️⃣ 正在思考 Python 解决方案...")
            
            try:
                try:
                    model = genai.GenerativeModel('gemini-2.5-flash')
                except:
                    model = genai.GenerativeModel('gemini-pro')

                dtypes_info = df.dtypes.to_string() if hasattr(df, 'dtypes') else "无"

                # --- 🔥 升级后的提示词：教 AI 画图 🔥 ---
                prompt = f"""
                你是一个 Python Pandas 和 Streamlit 专家。
                
                【当前数据情况】
                变量名: `df`
                列信息:
                {dtypes_info}
                
                【用户任务】
                "{user_query}"
                
                【关键规则 - 必读】
                1. **关于画图**：
                   - 优先使用 Streamlit 原生图表，因为它们支持中文且可交互：
                     - 柱状图用 `st.bar_chart(data)`
                     - 折线图用 `st.line_chart(data)`
                     - 散点图用 `st.scatter_chart(data)`
                   - 如果必须使用 `matplotlib`：
                     - **严禁**使用 `plt.show()` (在网页里无效)。
                     - 画完图后，必须调用 `st.pyplot(plt)` 来把图展示出来。
                     - 设置 `plt.figure(figsize=(10, 5))`。
                
                2. **关于数据修改**：
                   - 如果是修改数据，直接操作 `df`。
                   - 严禁把 `df` 变成非 DataFrame 对象。
                
                3. **关于输出**：
                   - 只输出 Python 代码，不要 markdown 标记。
                """
                
                # 生成
                response = model.generate_content(prompt)
                code = response.text.replace("```python", "").replace("```", "").strip()
                st.code(code, language='python')
                
                # 执行
                st.write("2️⃣ 正在执行...")
                
                # 捕获文字输出
                from io import StringIO
                import sys
                captured_output = StringIO()
                sys.stdout = captured_output
                
                # --- 把绘图库传给 AI ---
                local_vars = {
                    'df': df, 
                    'pd': pd, 
                    'st': st, 
                    'plt': plt, # 把 matplotlib 传进去
                    'matplotlib': matplotlib
                }
                
                exec(code, globals(), local_vars)
                
                # 恢复标准输出
                sys.stdout = sys.__stdout__
                output_str = captured_output.getvalue()
                
                if output_str:
                    st.info(f"🖨️ 计算结果:\n{output_str}")

                # 更新数据状态
                new_df = local_vars.get('df')
                if isinstance(new_df, pd.DataFrame):
                    st.session_state.df = new_df
                    status.update(label="✅ 任务完成！", state="complete", expanded=False)
                    time.sleep(1)
                    st.rerun()
                else:
                    status.update(label="✅ 任务完成", state="complete", expanded=False)
                
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
