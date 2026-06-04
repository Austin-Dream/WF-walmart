import streamlit as st
import pandas as pd
import io
import base64
from datetime import datetime

st.set_page_config(page_title="沃尔玛批量填写追踪号（粘贴版）", page_icon="📋", layout="wide")

def parse_pasted_data(pasted_text):
    """将粘贴的文本解析为DataFrame，支持Tab、逗号、空格分隔"""
    import re
    lines = pasted_text.strip().splitlines()
    if not lines:
        return None
    # 判断分隔符：优先Tab，其次逗号，再空格
    delimiter = None
    if '\t' in lines[0]:
        delimiter = '\t'
    elif ',' in lines[0]:
        delimiter = ','
    else:
        # 尝试按空格分割（连续空格视为一个）
        rows = [re.split(r'\s+', line) for line in lines]
        max_cols = max(len(row) for row in rows)
        # 填充不一致的行
        for row in rows:
            if len(row) < max_cols:
                row.extend([''] * (max_cols - len(row)))
        return pd.DataFrame(rows)
    # 使用pandas读取
    from io import StringIO
    try:
        df = pd.read_csv(StringIO(pasted_text), sep=delimiter, engine='python')
        return df
    except:
        # 如果失败，手动分割
        rows = [line.split(delimiter) for line in lines]
        max_cols = max(len(row) for row in rows)
        for row in rows:
            if len(row) < max_cols:
                row.extend([''] * (max_cols - len(row)))
        return pd.DataFrame(rows)

def generate_walmart_template(df, order_col, tracking_col, carrier_col):
    """生成沃尔玛批量上传格式"""
    result = pd.DataFrame()
    result['Order ID'] = df[order_col].astype(str).str.strip()
    result['Tracking Number'] = df[tracking_col].astype(str).str.strip()
    if carrier_col and carrier_col in df.columns:
        result['Carrier Name'] = df[carrier_col].astype(str).str.strip()
    else:
        result['Carrier Name'] = ''
    # 可选：添加当前日期作为发货日期（沃尔玛可能需要）
    result['Ship Date'] = datetime.now().strftime('%Y-%m-%d')
    # 删除全空的行
    result = result.dropna(subset=['Order ID', 'Tracking Number'], how='all')
    return result

def get_download_link(df, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    b64 = base64.b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">点击下载沃尔玛批量上传文件</a>'
    return href

def main():
    st.title("📋 沃尔玛批量填写追踪号（粘贴版）")
    st.markdown("""
    ### 使用说明
    1. 从同事更新的Excel表格中**复制**数据（包括表头行，例如订单号、追踪号、承运商等列）
    2. 在下方的文本框中**粘贴**（Ctrl+V 或 Cmd+V）
    3. 程序自动解析表格，让您选择对应的列
    4. 点击生成，即可下载**沃尔玛标准批量上传模板**（Excel格式）
    
    > ✅ 无需上传任何文件，数据仅在您浏览器中处理，安全便捷。
    """)
    
    # 粘贴区域
    pasted_text = st.text_area("📌 请在此处粘贴数据（支持从Excel复制多行多列）", height=200,
                               help="从Excel选中数据区域（包含表头），复制后粘贴到这里")
    
    if pasted_text:
        # 解析粘贴内容
        df = parse_pasted_data(pasted_text)
        if df is None or df.empty:
            st.error("粘贴内容为空，请检查")
            return
        
        st.success(f"成功解析 {len(df)} 行 × {len(df.columns)} 列")
        st.markdown("#### 解析后的数据预览（前10行）")
        st.dataframe(df.head(10))
        
        # 列映射
        st.markdown("### 列映射设置")
        col_list = df.columns.tolist()
        
        order_col = st.selectbox("选择【订单号】列", col_list, index=find_column_index(col_list, ['order', 'po', 'order id', '订单号', '订单编号']))
        tracking_col = st.selectbox("选择【追踪号】列", col_list, index=find_column_index(col_list, ['tracking', 'track', '追踪号', '单号', '快递单号']))
        carrier_col = st.selectbox("选择【承运商】列（可选，没有则留空）", [''] + col_list, index=0)
        
        if st.button("🚀 生成沃尔玛批量上传文件"):
            if not order_col or not tracking_col:
                st.error("请至少选择订单号和追踪号列")
                return
            result_df = generate_walmart_template(df, order_col, tracking_col, carrier_col)
            if result_df.empty:
                st.warning("没有有效的订单数据，请检查列内容")
                return
            
            st.success(f"成功生成 {len(result_df)} 条记录")
            st.dataframe(result_df.head(20))
            
            # 下载
            filename = f"walmart_bulk_tracking_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.markdown(get_download_link(result_df, filename), unsafe_allow_html=True)
            st.info("下载后可直接在沃尔玛卖家中心上传（订单管理 → 批量上传追踪号）")

def find_column_index(col_list, keywords):
    """根据关键词找到第一个匹配的列索引"""
    for i, col in enumerate(col_list):
        col_lower = str(col).lower()
        for kw in keywords:
            if kw.lower() in col_lower:
                return i
    return 0

if __name__ == "__main__":
    main()
