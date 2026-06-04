import streamlit as st
import pandas as pd
import io
import base64
import re
from datetime import datetime

st.set_page_config(page_title="沃尔玛批量上传单号生成器", page_icon="📦", layout="wide")

def detect_separator(text):
    """检测文本的分隔符，返回最可能的分隔符"""
    lines = text.strip().splitlines()
    if not lines:
        return None
    # 取前5行样本
    sample = "\n".join(lines[:5])
    # 候选分隔符
    separators = ['\t', ',', ';', '|', r'\s+']
    best_sep = None
    best_score = 0
    for sep in separators:
        if sep == r'\s+':
            # 按连续空格分割，检查每行字段数是否一致
            rows = [re.split(r'\s+', line) for line in lines[:5]]
        else:
            rows = [line.split(sep) for line in lines[:5]]
        if not rows:
            continue
        # 计算每行的字段数
        col_counts = [len(r) for r in rows]
        if max(col_counts) == min(col_counts) and max(col_counts) > 1:
            # 如果所有行字段数相同且大于1，视为有效
            score = col_counts[0]
            if score > best_score:
                best_score = score
                best_sep = sep
    return best_sep

def parse_pasted_data(pasted_text, manual_sep=None):
    """解析粘贴数据，返回DataFrame"""
    if not pasted_text:
        return None
    lines = pasted_text.strip().splitlines()
    if len(lines) == 0:
        return None
    
    # 如果手动指定了分隔符
    if manual_sep and manual_sep != "自动检测":
        sep = manual_sep
        if sep == "空格":
            # 按连续空格分割
            rows = [re.split(r'\s+', line) for line in lines]
        else:
            rows = [line.split(sep) for line in lines]
        # 确保每行列数相同
        max_cols = max(len(r) for r in rows)
        for r in rows:
            if len(r) < max_cols:
                r.extend([''] * (max_cols - len(r)))
        df = pd.DataFrame(rows)
        # 尝试将第一行作为表头
        try:
            df.columns = df.iloc[0].astype(str)
            df = df[1:].reset_index(drop=True)
        except:
            pass
        return df
    
    # 自动检测分隔符
    sep = detect_separator(pasted_text)
    if sep is None:
        st.error("无法自动检测分隔符，请手动选择")
        return None
    
    try:
        from io import StringIO
        # 如果分隔符是正则空格，不能用read_csv
        if sep == r'\s+':
            rows = [re.split(r'\s+', line) for line in lines]
            max_cols = max(len(r) for r in rows)
            for r in rows:
                if len(r) < max_cols:
                    r.extend([''] * (max_cols - len(r)))
            df = pd.DataFrame(rows)
            # 尝试将第一行作为表头
            try:
                df.columns = df.iloc[0].astype(str)
                df = df[1:].reset_index(drop=True)
            except:
                pass
        else:
            # 使用pandas读取
            df = pd.read_csv(StringIO(pasted_text), sep=sep, engine='python')
        return df
    except Exception as e:
        st.error(f"解析错误：{e}")
        return None

def generate_walmart_bulk(df, order_col, tracking_col):
    """生成沃尔玛批量上传格式，承运商固定为 FedEx"""
    result = pd.DataFrame()
    result['Order ID'] = df[order_col].astype(str).str.strip()
    result['Tracking Number'] = df[tracking_col].astype(str).str.strip()
    result['Carrier Name'] = 'FedEx'
    result['Ship Date'] = datetime.now().strftime('%Y-%m-%d')
    # 过滤空值
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
    1. 从同事更新后的Excel中**复制所有数据**（包括表头行）
    2. 粘贴到下方文本框
    3. 如果自动解析失败，请手动选择**分隔符**（通常是 **Tab** 或 **逗号**）
    4. 选择**订单号列**和**追踪号列**（同事填入的列，如 `Update Tracking Number`）
    5. 点击生成，下载沃尔玛标准批量上传文件
    
    > 承运商自动设为 **FedEx**，发货日期为今天
    """)
    
    pasted = st.text_area("📋 请粘贴数据（从Excel复制，含表头）", height=250)
    
    if pasted:
        # 让用户选择分隔符
        sep_choice = st.selectbox("分隔符选择", ["自动检测", "Tab (\\t)", "逗号 (,)", "分号 (;)", "竖线 (|)", "空格"], index=0)
        manual_sep_map = {
            "Tab (\\t)": "\t",
            "逗号 (,)": ",",
            "分号 (;)": ";",
            "竖线 (|)": "|",
            "空格": "空格",
            "自动检测": None
        }
        manual_sep = manual_sep_map.get(sep_choice, None)
        
        df = parse_pasted_data(pasted, manual_sep)
        if df is None or df.empty:
            st.error("解析失败，请检查粘贴内容或手动选择正确的分隔符")
            # 显示原始文本的前几行，帮助调试
            lines = pasted.strip().splitlines()
            st.text("粘贴内容前5行：")
            for i, line in enumerate(lines[:5]):
                st.text(f"{i+1}: {line[:200]}")
            return
        
        st.success(f"成功解析 {len(df)} 行 × {len(df.columns)} 列")
        st.dataframe(df.head(10))
        
        cols = df.columns.tolist()
        # 智能推荐订单号列
        default_order = 0
        for i, c in enumerate(cols):
            c_low = str(c).lower()
            if c_low in ['po#', 'order#', 'order id', '订单号', 'po']:
                default_order = i
                break
        order_col = st.selectbox("选择【订单号】列", cols, index=default_order)
        
        # 智能推荐追踪号列（优先匹配 tracking / 单号）
        default_track = 0
        for i, c in enumerate(cols):
            c_low = str(c).lower()
            if 'tracking' in c_low or '单号' in c_low:
                default_track = i
                break
        tracking_col = st.selectbox("选择【追踪号】列（同事填写的列，如 V 列或 Update Tracking Number）", cols, index=default_track)
        
        if st.button("🚀 生成沃尔玛批量上传文件"):
            result = generate_walmart_bulk(df, order_col, tracking_col)
            if result.empty:
                st.warning("没有有效的订单-追踪号记录，请检查对应列是否有值")
                return
            st.success(f"生成 {len(result)} 条记录")
            st.dataframe(result)
            
            filename = f"walmart_bulk_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.markdown(get_download_link(result, filename), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
