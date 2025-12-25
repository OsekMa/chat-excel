import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import time # 引入时间库，为了展示工作流效果

# --- 页面配置 ---
st.set_page_config(page_title="AI Excel 超级助手", page_icon="🚀", layout="wide") # layout="wide" 让表格展示更宽
st.title("🚀 AI Excel 超级助手")

# --- 1. 获取 API Key ---
# 依然从后台 Secrets 获取，安全第一
api_key = st.secrets.get("GOOGLE_API_KEY")

if not api_key:
    st.error("请在 Streamlit 后台 Secrets 设置 GOOGLE_API_KEY")
    st.stop()

genai.configure(api_key=api_key)

# --- 2. 侧边栏：文件操作 ---
with st.sidebar:
    st.header("📂 文件操作区")
    uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])
    
    # 添加一个重置按钮
    if st.button("🔄 重置所有操作"):
        if 'df' in st.session_state:
            del st.session_state['df']
        st.rerun()

# --- 3. 核心逻辑 ---
if uploaded_file:
    # 初始化 session_state，保证刷新页面数据不丢
    if 'df' not in st.session_state:
        try:
            st.session_state.df = pd.read_excel(uploaded_file)
            st.toast("✅ 文件上传成功！")
        except Exception as e:
            st.error(f"文件读取失败: {e}")
            st.stop()
    
    df = st.session_state.df

    # --- 升级点 1：全量展示数据 ---
    st.subheader("📊 数据全览")
    # use_container_width=True 会让表格自动铺满屏幕宽度
    # height=400 限制高度，超过会有滚动条，防止数据太多把页面撑爆
    st.dataframe(df, use_container_width=True, height=400) 
    
    # 展示一下当前的行数和列数
    st.caption(f"当前共 {df.shape[0]} 行, {df.shape[1]} 列数据")

    # --- 聊天输入框 ---
    st.divider()
    user_query = st.chat_input("💡 请下达指令，例如：'把销售额大于500的标红' 或 '删除空行'...")

    if user_query:
        # --- 升级点 3：可视化工作流 ---
        # 使用 st.status 创建一个状态容器
        with st.status("🤖 AI 正在干活，请稍候...", expanded=True) as status:
            
            # 第一步：分析
            st.write("1️⃣ 正在阅读表格结构和数据类型...")
            # 获取列名和数据类型，帮助 AI 更准确判断
            dtypes_info = df.dtypes.to_string()
            
            # 第二步：思考
            st.write("2️⃣ 正在思考 Python 解决方案...")
            
            try:
                # 尝试使用更强的模型，如果失败会自动回退
                try:
                    model = genai.GenerativeModel('gemini-1.5-flash-001')
                except:
                    model = genai.GenerativeModel('gemini-pro')

                # --- 升级点 2：更完善的提示词 (Prompt Engineering) ---
                prompt = f"""
                你是一个 Python Pandas 高级专家。
                
                【当前数据情况】
                1. DataFrame 变量名为: `df`
                2. 列名及数据类型如下:
                {dtypes_info}
                3. 前 3 行数据样例:
                {df.head(3).to_markdown()}
                
                【用户任务】
                "{user_query}"
                
                【严格约束】
                1. 必须生成可执行的 Python 代码。
                2. 代码必须修改 `df` 变量（例如 `df = ...` 或 `df.drop(..., inplace=True)`）。
                3. 不需要导入 pandas，环境已预置 `pd` 和 `df`。
                4. 不要包含 ```python 或 ``` 标记，只输出纯代码。
                5. 如果涉及字符串匹配，请注意处理空值和大小写问题，增强代码鲁棒性。
                6. 如果用户问的是查询（如“计算总和”），请使用 `st.write()` 将结果打印出来。
                """
                
                # 第三步：生成代码
                st.write("3️⃣ 正在生成执行代码...")
                response = model.generate_content(prompt)
                code = response.text.replace("```python", "").replace("```", "").strip()
                
                # 展示生成的代码（让用户看到 AI 做了什么）
                st.code(code, language='python')
                
                # 第四步：执行
                st.write("4️⃣ 正在执行修改...")
                
                # 准备执行环境
                local_vars = {'df': df, 'pd': pd, 'st': st} # 把 st 传进去，让 AI 可以打印结果
                exec(code, globals(), local_vars)
                
                # 更新状态
                st.session_state.df = local_vars['df']
                
                # 更新工作流状态为完成
                status.update(label="✅ 任务圆满完成！", state="complete", expanded=False)
                
                # 稍微停顿一下让用户看到“完成”的状态，然后刷新
                time.sleep(1.5)
                st.rerun()
                
            except Exception as e:
                status.update(label="❌ 任务执行失败", state="error", expanded=True)
                st.error(f"出错啦: {e}")
                st.error("建议：尝试换一种说法，或者检查指令是否符合当前表格结构。")

    # --- 下载区域 ---
    st.divider()
    if not df.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 下载修改后的 Excel",
            data=output.getvalue(),
            file_name="AI_Modified_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True # 按钮也变宽
        )

else:
    # 没上传文件时的欢迎界面
    st.info("👈 请先在左侧侧边栏上传 Excel 文件开始工作")
