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
        picking_df1 = pd.read_csv(picking_list_1, encoding='shift_jis', dtype={'JANコード': str})
        picking_df2 = pd.read_csv(picking_list_2, encoding='shift_jis', dtype={'JANコード': str})
        
        # 数値型として読み込まれた場合に備えて文字列に変換
        picking_df1['JANコード'] = picking_df1['JANコード'].astype(str)
        picking_df2['JANコード'] = picking_df2['JANコード'].astype(str)

        # ピッキングリストを結合
        picking_df = pd.concat([picking_df1, picking_df2])
        
        # 受注数をJANコードごとに合計
        order_summary = picking_df.groupby('JANコード')['受注数'].sum().reset_index()
        order_summary.columns = ['JAN', '受注数']  # 列名を一致させる
        
        # デバッグ用の出力
        st.write("ピッキングリストの受注数合計:", order_summary)

        # 在庫表を読み込み
        inventory_sheets = pd.read_excel(inventory_file, sheet_name=None)
        
        # デバッグ用の出力
        st.write("在庫表のシート一覧:", list(inventory_sheets.keys()))
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, sheet_df in inventory_sheets.items():
                st.write(f"シート名: {sheet_name}")

                # 列名をすべて小文字に変換してチェック
                sheet_df.columns = sheet_df.columns.str.lower()

                if 'jan' in sheet_df.columns:
                    # 数値型として読み込まれた場合に備えて文字列に変換
                    sheet_df['jan'] = sheet_df['jan'].astype(str)
                    
                    # デバッグ用の出力
                    st.write("元のシートデータ:", sheet_df.head())

                    # 在庫表のJANとピッキングリストの受注数を結合
                    result_df = sheet_df.merge(order_summary, on='jan', how='left')
                    result_df['受注数'] = result_df['受注数'].fillna(0).astype(int)
                    
                    # デバッグ用の出力
                    st.write("結合後のデータ:", result_df.head())
                    
                    result_df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    st.write(f"シート {sheet_name} に JAN 列が見つかりません")
        
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
