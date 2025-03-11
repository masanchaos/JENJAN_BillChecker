import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO
from datetime import datetime

def process_excel(uploaded_file):
    # 使用上傳的文件 (不依賴檔名) 讀取所有分頁
    try:
        xls = pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"讀取 Excel 檔案錯誤: {e}")
        return None

    # 取得「費用編號表」分頁
    df_fee = xls.get('費用編號表')
    if df_fee is None:
        st.error("找不到『費用編號表』分頁。")
        return None

    # 檢查必需欄位
    required_fee_columns = ['費用編號', '所屬', '項目']
    for col in required_fee_columns:
        if col not in df_fee.columns:
            st.error(f"『費用編號表』中缺少必需欄位：{col}")
            return None

    # 建立映射字典 {費用編號: {'所屬': 所屬, '項目': 項目}}
    fee_mapping = df_fee.set_index('費用編號')[['所屬', '項目']].to_dict('index')

    # 取得「客戶列表」分頁
    df_customer = xls.get('客戶列表')
    if df_customer is None:
        st.error("找不到『客戶列表』分頁。")
        return None

    if '客戶名稱' not in df_customer.columns:
        st.error("『客戶列表』中缺少必需欄位：客戶名稱")
        return None

    # 在「客戶編號」前新增「請檢查」欄（如果沒有「客戶編號」，則插在最前面）
    if '客戶編號' in df_customer.columns:
        pos = df_customer.columns.get_loc('客戶編號')
        df_customer.insert(pos, '請檢查', '')
    else:
        df_customer.insert(0, '請檢查', '')

    # 新增「分倉應收賬款（未稅）」欄，預設為 0
    df_customer['分倉應收賬款（未稅）'] = 0

    # 遍歷每個客戶
    for idx, cust_row in df_customer.iterrows():
        # 檢查原始客戶名稱是否為 NaN
        cust_raw = cust_row['客戶名稱']
        if pd.isna(cust_raw):
            df_customer.loc[idx, '請檢查'] = ""
            continue

        cust_name = str(cust_raw).strip()
        # 若經過 strip 處理後為空，則不填入 *** 且不做後續處理
        if not cust_name:
            df_customer.loc[idx, '請檢查'] = ""
            continue

        cust_code = cust_name[:2]

        # 從所有分頁中找出名稱開頭為 cust_code 的分頁（排除「費用編號表」和「客戶列表」）
        target_sheet_name = None
        for sheet_name in xls.keys():
            if sheet_name in ['費用編號表', '客戶列表']:
                continue
            if sheet_name.startswith(cust_code):
                target_sheet_name = sheet_name
                break

        if target_sheet_name is None:
            df_customer.loc[idx, '請檢查'] = "***"
            continue

        # 取得該客戶分頁數據
        try:
            df_cust = xls[target_sheet_name].copy()
        except Exception as e:
            st.error(f"讀取客戶分頁 {target_sheet_name} 失敗: {e}")
            continue

        # 檢查必需欄位
        required_customer_columns = ['費用編號', '費用', '總計']
        missing_cols = [col for col in required_customer_columns if col not in df_cust.columns]
        if missing_cols:
            st.error(f"客戶分頁 {target_sheet_name} 缺少必需欄位：{', '.join(missing_cols)}")
            continue

        # 新增「項目名」欄：根據「費用編號」查找 fee_mapping 中的「項目」值
        df_cust['項目名'] = df_cust['費用編號'].map(lambda x: fee_mapping[x]['項目'] if x in fee_mapping else "")

        # 將「項目名」欄插入到「費用編號」和「費用」之間
        if '費用編號' in df_cust.columns and '費用' in df_cust.columns:
            fee_code_index = df_cust.columns.get_loc('費用編號')
            cols = list(df_cust.columns)
            cols.remove('項目名')
            new_cols = cols[:fee_code_index+1] + ['項目名'] + cols[fee_code_index+1:]
            df_cust = df_cust[new_cols]

        # 計算該客戶的賬款總數
        total = 0
        for i, fee_row in df_cust.iterrows():
            fee_code = fee_row['費用編號']
            amount = fee_row['總計']
            if fee_code in fee_mapping:
                belong = fee_mapping[fee_code]['所屬']
                if belong == 2:
                    total += amount
        df_customer.loc[idx, '分倉應收賬款（未稅）'] = total

        # 更新 xls 中該客戶分頁
        xls[target_sheet_name] = df_cust

    # 新增「分倉應收賬款（含稅）」欄，公式為「分倉應收賬款（未稅）」 * 1.05 四捨五入取整
    df_customer['分倉應收賬款（含稅）'] = (df_customer['分倉應收賬款（未稅）'] * 1.05).round().astype(int)
    xls['客戶列表'] = df_customer

    # 將所有分頁寫入到 BytesIO 中
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        for sheet_name, df in xls.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output_buffer.seek(0)

    # 用 openpyxl 調整「客戶列表」中「請檢查」欄為 "***" 的行文字顏色為紅色
    wb = openpyxl.load_workbook(output_buffer)
    ws = wb['客戶列表']
    header = [cell.value for cell in ws[1]]
    check_idx = header.index("請檢查") + 1
    for row in ws.iter_rows(min_row=2):
        if row[check_idx-1].value == "***":
            for cell in row:
                cell.font = Font(color="FF0000")
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# Streamlit 界面
st.title("JENJAN 對賬小幫手 - 賬單生成")
st.write("請上傳您的 Excel 賬單（檔名不限），點擊【生成新賬單】後即可下載新賬單。")

uploaded_file = st.file_uploader("選擇 Excel 文件", type=["xlsx"])
if uploaded_file is not None:
    if st.button("生成新賬單"):
        result = process_excel(uploaded_file)
        if result is not None:
            # 根據上傳檔案的檔名與當前時間組合輸出檔名
            input_filename = uploaded_file.name
            if input_filename.lower().endswith('.xlsx'):
                input_filename = input_filename[:-5]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"對賬後_{input_filename}_{timestamp}.xlsx"

            st.download_button(
                label="下載生成的新賬單",
                data=result,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
