import streamlit as st
import pandas as pd
from io import BytesIO

# ファイルアップロード
st.title('ピッキングリストと在庫表の処理')
picking_list_1 = st.file_uploader('ピッキングリスト1をアップロード (CSV)', type='csv')
picking_list_2 = st.file_uploader('ピッキングリスト2をアップロード (CSV)', type='csv')
inventory_file = st.file_uploader('在庫表をアップロード (Excel)', type='xlsx')

if picking_list_1 and picking_list_2 and inventory_file:
    try:
        # ピッキングリストの読み込み（エンコーディングを指定）
        picking_df1 = pd.read_csv(picking_list_1, encoding='shift_jis')  # Shift-JISとして読み込む
        picking_df2 = pd.read_csv(picking_list_2, encoding='shift_jis')  # Shift-JISとして読み込む
        
        # ピッキングリストを結合
        picking_df = pd.concat([picking_df1, picking_df2])
        
        # 受注数をJANコードごとに合計
        order_summary = picking_df.groupby('JANコード')['受注数'].sum().reset_index()
        order_summary.columns = ['JAN', '受注数']  # 列名を一致させる
        
        # 在庫表を読み込み
        inventory_sheets = pd.read_excel(inventory_file, sheet_name=None)
        
        # 結果を新しいExcelファイルに保存
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, sheet_df in inventory_sheets.items():
                if 'JAN' in sheet_df.columns:
                    # 在庫表のJANとピッキングリストの受注数を結合
                    result_df = sheet_df[['JAN']].merge(order_summary, on='JAN', how='left')
                    result_df['受注数'] = result_df['受注数'].fillna(0).astype(int)
                    result_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        st.success('処理が完了しました。以下のボタンから結果をダウンロードしてください。')
        
        # ダウンロードボタン
        st.download_button(
            label="結果をダウンロード",
            data=output.getvalue(),
            file_name="在庫表_結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
