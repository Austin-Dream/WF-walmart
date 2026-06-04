import streamlit as st
import pandas as pd
import io
import base64
from datetime import datetime

st.set_page_config(page_title="沃尔玛批量填写追踪号", page_icon="🚚", layout="wide")

def load_excel(file):
    """加载Excel文件，返回DataFrame和原始文件名"""
    try:
        df = pd.read_excel(file, sheet_name=0)  # 默认第一个sheet
        return df, file.name
    except Exception as e:
        st.error(f"读取文件失败: {e}")
        return None, None

def match_and_update(to_update_df, shipped_df, match_col, update_cols):
    """
    根据匹配列，用shipped_df中的对应列值更新to_update_df
    update_cols: dict {目标列名: 源列名}  例如 {"Tracking Number": "Tracking Number", "Carrier": "Carrier"}
    返回更新后的df，及更新统计
    """
    # 复制，避免修改原数据
    updated_df = to_update_df.copy()
    
    # 将匹配列转换为字符串，确保匹配
    shipped_df[match_col] = shipped_df[match_col].astype(str).str.strip()
    updated_df[match_col] = updated_df[match_col].astype(str).str.strip()
    
    # 建立映射字典：订单号 -> 行数据（只取需要更新的列）
    shipped_dict = {}
    for _, row in shipped_df.iterrows():
        key = row[match_col]
        if key not in shipped_dict:
            shipped_dict[key] = row
        else:
            # 如果同一个订单有多个行（例如多个包裹），可以选择合并或取第一个，这里提示用户
            st.warning(f"订单号 {key} 在已发货文件中出现多次，将只使用第一行的追踪信息。")
    
    updated_count = 0
    for idx, row in updated_df.iterrows():
        key = row[match_col]
        if key in shipped_dict:
            shipped_row = shipped_dict[key]
            for target_col, source_col in update_cols.items():
                if source_col in shipped_row and pd.notna(shipped_row[source_col]):
                    updated_df.at[idx, target_col] = shipped_row[source_col]
            updated_count += 1
    
    return updated_df, updated_count

def get_download_link(df, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    b64 = base64.b64encode(processed_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">点击下载更新后的文件</a>'
    return href

def main():
    st.title("📦 沃尔玛批量填写追踪号（单号）")
    st.markdown("根据已发货订单的追踪信息，自动批量填写到未发货订单表中。")
    
    with st.expander("📖 使用说明"):
        st.markdown("""
        1. **上传待更新文件**：未发货的订单表格（需包含订单号列，以及需要填写追踪号的列，例如V列）。
        2. **上传已发货文件**：已经完成发货的表格（包含订单号和对应的追踪号、承运商等）。
        3. **设置匹配列**：选择两个文件中共同的订单标识列（如 `PO#` 或 `Order#`）。
        4. **设置更新列**：指定要将哪些信息填到待更新文件的哪些列。
        5. 点击 **开始匹配更新**，预览结果并下载更新后的文件。
        """)
    
    col1, col2 = st.columns(2)
    with col1:
        to_update_file = st.file_uploader("📂 上传待更新文件（未发货）", type=["xlsx", "xls"])
    with col2:
        shipped_file = st.file_uploader("📦 上传已发货文件", type=["xlsx", "xls"])
    
    if to_update_file and shipped_file:
        # 加载数据
        to_update_df, to_update_name = load_excel(to_update_file)
        shipped_df, shipped_name = load_excel(shipped_file)
        
        if to_update_df is not None and shipped_df is not None:
            st.success(f"待更新文件加载成功：{len(to_update_df)} 行")
            st.success(f"已发货文件加载成功：{len(shipped_df)} 行")
            
            with st.expander("🔍 预览待更新文件前5行"):
                st.dataframe(to_update_df.head())
            with st.expander("🔍 预览已发货文件前5行"):
                st.dataframe(shipped_df.head())
            
            # 列选择
            st.markdown("### 1. 选择匹配列（订单唯一标识）")
            match_col = st.selectbox(
                "用于匹配两个文件的订单号列",
                options=to_update_df.columns.tolist(),
                index=to_update_df.columns.tolist().index("PO#") if "PO#" in to_update_df.columns else 0
            )
            # 确保两个文件都有该列
            if match_col not in shipped_df.columns:
                st.error(f"已发货文件中没有找到列「{match_col}」，请检查文件格式或手动选择匹配列。")
                # 让用户手动选择已发货文件的匹配列
                shipped_match_col = st.selectbox("在已发货文件中选择对应的订单号列", shipped_df.columns.tolist())
            else:
                shipped_match_col = match_col
            
            st.markdown("### 2. 选择要更新的列")
            st.info("将已发货文件中的信息填入待更新文件的对应列。")
            
            # 预设常见字段：追踪号、承运商、追踪链接
            update_mappings = {}
            
            # 追踪号
            tracking_col_in_shipped = st.selectbox(
                "已发货文件中的【追踪号】列",
                options=[""] + shipped_df.columns.tolist(),
                index=0
            )
            if tracking_col_in_shipped:
                target_col_in_update = st.selectbox(
                    "待更新文件中要填入追踪号的列（例如V列）",
                    options=to_update_df.columns.tolist(),
                    index=to_update_df.columns.tolist().index("Tracking Number") if "Tracking Number" in to_update_df.columns else 0
                )
                update_mappings[target_col_in_update] = tracking_col_in_shipped
            
            # 承运商
            carrier_col_in_shipped = st.selectbox(
                "已发货文件中的【承运商】列（可选）",
                options=[""] + shipped_df.columns.tolist(),
                index=0
            )
            if carrier_col_in_shipped:
                target_carrier_col = st.selectbox(
                    "待更新文件中要填入承运商的列",
                    options=to_update_df.columns.tolist(),
                    index=to_update_df.columns.tolist().index("Carrier") if "Carrier" in to_update_df.columns else 0
                )
                update_mappings[target_carrier_col] = carrier_col_in_shipped
            
            # 追踪链接
            url_col_in_shipped = st.selectbox(
                "已发货文件中的【追踪链接】列（可选）",
                options=[""] + shipped_df.columns.tolist(),
                index=0
            )
            if url_col_in_shipped:
                target_url_col = st.selectbox(
                    "待更新文件中要填入追踪链接的列",
                    options=to_update_df.columns.tolist(),
                    index=to_update_df.columns.tolist().index("Tracking Url") if "Tracking Url" in to_update_df.columns else 0
                )
                update_mappings[target_url_col] = url_col_in_shipped
            
            if st.button("🚀 开始匹配更新"):
                with st.spinner("正在匹配并更新数据..."):
                    updated_df, updated_count = match_and_update(
                        to_update_df, 
                        shipped_df, 
                        match_col, 
                        update_mappings,
                        shipped_match_col  # 注意上面函数需要支持不同列名，修改一下
                    )
                    # 修正：上面函数只用一个match_col，需要修改函数定义
                    # 为了简洁，直接修改上面函数，支持两个不同的列名参数
                    # 这里重新实现一下匹配逻辑，更清晰
                    # 为了避免混乱，我重新写一个匹配函数在这里
                    def do_match_update(df_target, df_source, target_key_col, source_key_col, mappings):
                        df_target = df_target.copy()
                        df_source[source_key_col] = df_source[source_key_col].astype(str).str.strip()
                        df_target[target_key_col] = df_target[target_key_col].astype(str).str.strip()
                        source_dict = {}
                        for _, row in df_source.iterrows():
                            key = row[source_key_col]
                            if key not in source_dict:
                                source_dict[key] = row
                        updated = 0
                        for idx, row in df_target.iterrows():
                            key = row[target_key_col]
                            if key in source_dict:
                                src_row = source_dict[key]
                                for t_col, s_col in mappings.items():
                                    if s_col in src_row and pd.notna(src_row[s_col]):
                                        df_target.at[idx, t_col] = src_row[s_col]
                                updated += 1
                        return df_target, updated
                    
                    updated_df, updated_count = do_match_update(
                        to_update_df, shipped_df,
                        match_col, shipped_match_col,
                        update_mappings
                    )
                    
                st.success(f"匹配完成！共更新了 {updated_count} 行记录。")
                
                # 显示更新预览（仅显示被更新的行）
                updated_rows = updated_df[updated_df[match_col].isin(shipped_df[shipped_match_col].astype(str).str.strip())]
                if not updated_rows.empty:
                    st.subheader("📋 更新预览（前10条）")
                    st.dataframe(updated_rows.head(10))
                else:
                    st.warning("未找到任何匹配的订单，请检查匹配列是否正确。")
                
                # 下载
                output_filename = f"updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                st.markdown(get_download_link(updated_df, output_filename), unsafe_allow_html=True)
    
    st.markdown("---")
    st.caption("提示：如列名与示例不符，请手动选择正确的列。匹配时订单号应完全一致（包括数字格式）。")

if __name__ == "__main__":
    main()
