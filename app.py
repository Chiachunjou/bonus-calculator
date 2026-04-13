import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="業務獎金結算系統", layout="centered")
st.title("📊 業務獎金與訂單結算系統")
st.write("上傳當月業績表，系統將自動過濾退貨、計算獎金，並產出各業務專屬的分頁報表。")

# 1. 前端輸入
total_goal = st.number_input("1. 填寫全公司總體目標", min_value=1.0, value=1000000.0, step=10000.0)
uploaded_file = st.file_uploader("2. 上傳原始訂單 Excel / CSV 檔", type=["xlsx", "csv"])

if st.button("🚀 開始結算並產出報表") and uploaded_file is not None:
    try:
        with st.spinner("雲端運算中，請稍候..."):
            # 讀取檔案
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, header=None, dtype=str).fillna("")
            else:
                df = pd.read_excel(uploaded_file, header=None, dtype=str).fillna("")

            # ==========================================
            # 【關鍵修正】改回你正常表格的欄位索引 (A=0, B=1...)
            # ==========================================
            COL_STATUS = 6   # G欄：退貨狀態
            COL_AMOUNT = 7   # H欄：金額
            COL_SALES  = 25  # Z欄：商家負責業務
            COL_AA     = 26  # AA欄：數值 1
            COL_AB     = 27  # AB欄：數值 2
            COL_COMM   = 29  # AD欄：傭金
            COL_DISC   = 30  # AE欄：平台折扣
            
            # 【防呆機制】如果 Excel 最後幾欄完全空白，Pandas 可能會讀不到那些欄位
            # 這裡強制把欄位補齊到至少 31 欄 (AE欄)，避免 KeyError
            max_required_col = max([COL_STATUS, COL_AMOUNT, COL_SALES, COL_AA, COL_AB, COL_COMM, COL_DISC])
            for c in range(df.shape[1], max_required_col + 1):
                df[c] = ""

            headers = df.iloc[0].tolist()
            # 同步確保表頭長度足夠
            while len(headers) <= max_required_col:
                headers.append(f"未命名欄位_{len(headers)}")
                
            data = df.iloc[1:].copy()

            # 清理文字與轉換數值
            data[COL_STATUS] = data[COL_STATUS].astype(str).str.strip()
            data[COL_SALES] = data[COL_SALES].astype(str).str.strip()
            for col in [COL_AMOUNT, COL_AA, COL_AB, COL_COMM, COL_DISC]:
                data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0)

            # 篩選未申請的有效訂單
            valid_orders = data[data[COL_STATUS] == "未申請"]

            # 計算公司總表
            sum_h = valid_orders[COL_AMOUNT].sum()
            company_achieve_rate = sum_h / total_goal if total_goal > 0 else 0

            sales_names = [name for name in valid_orders[COL_SALES].unique() if name]
            sales_count = len(sales_names)
            individual_goal = (total_goal * 0.2) / sales_count if sales_count > 0 else 0

            # 準備產出 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 【工作表 1：總表】
                summary_data = [
                    ["【全公司業績獎金】", "", ""],
                    ["總體目標", total_goal, ""],
                    ["總金額加總", sum_h, ""],
                    ["公司達成率", f"{company_achieve_rate:.2%}", ""],
                    ["", "", ""],
                    ["【業務個人獎金】", "", ""],
                    ["個人目標", individual_goal, ""],
                    ["業務姓名", "實質收入(佣金-折扣)", "達成率"]
                ]

                sales_stats = []
                for name in sales_names:
                    p_orders = valid_orders[valid_orders[COL_SALES] == name]
                    sum_comm = p_orders[COL_COMM].sum()
                    sum_disc = p_orders[COL_DISC].sum()
                    real_income = sum_comm - sum_disc
                    rate = real_income / individual_goal if individual_goal > 0 else 0
                    
                    summary_data.append([name, real_income, f"{rate:.2%}"])
                    sales_stats.append((name, p_orders, sum_comm, sum_disc, real_income))

                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(writer, sheet_name='總表', index=False, header=False)

                # 【工作表 2：原始資料】
                df.to_excel(writer, sheet_name='原始資料', index=False, header=False)

                # 【工作表 3+：各業務分頁】
                for name, p_orders, sum_comm, sum_disc, real_income in sales_stats:
                    if len(p_orders) > 0:
                        p_df = pd.DataFrame(p_orders.values, columns=headers)
                        
                        order_count = len(p_orders)
                        avg_aa = p_orders[COL_AA].sum() / order_count if order_count > 0 else 0
                        avg_ab = p_orders[COL_AB].sum() / order_count if order_count > 0 else 0

                        # 製作底部的合計列
                        bottom_row = {col: "" for col in headers}
                        bottom_row[headers[COL_SALES]] = "合計/平均："
                        bottom_row[headers[COL_AA]] = avg_aa
                        bottom_row[headers[COL_AB]] = avg_ab
                        bottom_row[headers[COL_COMM]] = sum_comm
                        bottom_row[headers[COL_DISC]] = sum_disc
                        
                        # AF 欄 (即 COL_DISC 的下一欄) 填寫實質收入
                        if COL_DISC + 1 < len(headers):
                            bottom_row[headers[COL_DISC + 1]] = real_income

                        # 合併並寫入
                        p_df = pd.concat([p_df, pd.DataFrame([bottom_row])], ignore_index=True)
                        p_df.to_excel(writer, sheet_name=name, index=False)

            st.success("✅ 報表產出成功！")
            st.download_button(
                label="📥 點此下載統整好的 Excel",
                data=output.getvalue(),
                file_name="獎金結算報表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"系統發生錯誤，請確認上傳的表格格式是否正確。錯誤細節：{e}")
