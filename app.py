import streamlit as st
import pandas as pd
import io
import base64
from datetime import datetime

st.set_page_config(page_title="沃尔玛批量上传单号生成器", page_icon="📦", layout="wide")

def parse_pasted_data(pasted_text):
    """将粘贴的文本解析为DataFrame（支持Tab、逗号、空格分隔）"""
    import re
    lines = pasted_text.strip().splitlines()
    if not lines:
        return None
    # 优先尝试Tab分隔
    if '\t' in lines[0]:
        sep = '\t'
    elif ',' in lines[0]:
        sep = ','
    else:
        # 按空格分割（连续空格视为一个）
        rows = [re.split(r'\s+', line) for line in lines]
        max_cols = max(len(row) for row in rows)
        for row in rows:
            if len(row) < max_cols:
                row.extend([''] * (max_cols - len(row)))
        return pd.DataFrame(rows)
    from io import StringIO
    df = pd.read_csv(StringIO(pasted_text), sep=sep, engine='python')
    return df

def generate_walmart_bulk(df, order_col, tracking_col):
    """生成沃尔玛批量上传格式，承运商固定为 FedEx"""
    result = pd.DataFrame()
    result['Order ID'] = df[order_col].astype(str).str.strip()
    result['Tracking Number'] = df[tracking_col].astype(str).str.strip()
    result['Carrier Name'] = 'FedEx'   # 固定承运商
    result['Ship Date'] = datetime.now().strftime('%Y-%m-%d')
    # 删除订单号或追踪号为空的记录
    result = result.dropna(subset=['Order ID', 'Tracking Number'], how='any')
    result = result[result['Order ID'] != '']
    result = result[result['Tracking Number'] != '']
    return result

def get_download_link(df, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    b64 = base64.b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">点击下载沃尔玛批量上传文件</a>'
    return href

def main():
    st.title("📦 沃尔玛批量上传单号生成器")
    st.markdown("""
    ### 使用说明
    1. 从同事更新后的Excel中**复制**所有数据（包括表头行）
    2. 粘贴到下方文本框
    3. 选择**订单号列**和**追踪号列**（同事填入的列，如 V 列）
    4. 点击生成，下载沃尔玛标准批量上传文件
    
    > ✅ 承运商自动设为 **FedEx**，发货日期为今天
    """)
    
    pasted = st.text_area("📋 请在此处粘贴数据（从Excel复制，含表头）", height=250)
    
    if pasted:
        df = parse_pasted_data(pasted)
        if df is None or df.empty:
            st.error("解析失败，请检查粘贴内容")
            return
        
        st.success(f"成功解析 {len(df)} 行 × {len(df.columns)} 列")
        st.dataframe(df.head(10))
        
        # 列选择
        cols = df.columns.tolist()
        # 智能推荐订单号列
        default_order = 0
        for i, c in enumerate(cols):
            if str(c).lower() in ['po#', 'order#', 'order id', '订单号']:
                default_order = i
                break
        order_col = st.selectbox("选择【订单号】列", cols, index=default_order)
        
        # 智能推荐追踪号列（优先匹配 "Tracking Number" 或 "Update Tracking Number"）
        default_track = 0
        for i, c in enumerate(cols):
            c_lower = str(c).lower()
            if 'tracking' in c_lower or '单号' in c_lower:
                default_track = i
                break
        tracking_col = st.selectbox("选择【追踪号】列（同事填写的列，如 V 列）", cols, index=default_track)
        
        if st.button("🚀 生成沃尔玛批量上传文件"):
            result = generate_walmart_bulk(df, order_col, tracking_col)
            if result.empty:
                st.warning("没有有效的订单-追踪号记录，请检查列内容")
                return
            st.success(f"生成 {len(result)} 条记录")
            st.dataframe(result)
            
            filename = f"walmart_bulk_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.markdown(get_download_link(result, filename), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
