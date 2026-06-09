import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Sheet Merger Automation", layout="centered")

st.title("📊 Excel Sheet Merger Automation")
st.write("अपनी Purchase Order (PO) की फाइलें यहाँ अपलोड करें और कस्टमाइज्ड शीट तुरंत डाउनलोड करें।")
st.markdown("---")

uploaded_files = st.file_uploader(
    "सभी Excel या CSV फाइल्स एक साथ यहाँ अपलोड करें:", 
    type=["csv", "xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    sheet1_list = []
    sheet2_list = []
    
    for file in uploaded_files:
        filename = file.name.lower()
        if filename.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
            
        df.columns = [c.lower().strip() for c in df.columns]
        
        if "sheet1" in filename or "netprice" in df.columns:
            sheet1_list.append(df)
        elif "sheet2" in filename or ("quantity" in df.columns and "unit_name" in df.columns):
            sheet2_list.append(df)

    if st.button("🚀 फाइलें मर्ज करें (Fix Format)"):
        if not sheet1_list or not sheet2_list:
            st.error("कृपया सुनिश्चित करें कि आपने Sheet1 और Sheet2 दोनों टाइप की फाइलें अपलोड की हैं।")
        else:
            try:
                df_s1_all = pd.concat(sheet1_list, ignore_index=True)
                df_s2_all = pd.concat(sheet2_list, ignore_index=True)
                
                lookup = df_s1_all[['vendor_name', 'item_name', 'pack_type', 'netprice']].drop_duplicates()
                
                cols_to_drop = [c for c in ['vendoraddress', 'unitaddress'] if c in df_s2_all.columns]
                df_s2_clean = df_s2_all.drop(columns=cols_to_drop)
                
                merged_df = pd.merge(df_s2_clean, lookup, on=['vendor_name', 'item_name'], how='left')
                merged_df = merged_df.rename(columns={'netprice': 'rate'})
                merged_df['amount'] = merged_df['rate'] * merged_df['quantity']
                
                final_columns = ['unit_name', 'item_name', 'pack_type', 'quantity', 'rate', 'amount', 'vendor_name']
                final_df = merged_df[final_columns]
                
                st.success("डेटा सफलतापूर्वक मर्ज हो गया है!")
                st.dataframe(final_df.head(10))
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name="Master_Report")
                processed_data = output.getvalue()
                
                st.download_button(
                    label="📥 कस्टमाइज्ड एक्सेल फाइल डाउनलोड करें",
                    data=processed_data,
                    file_name="Automated_Master_Sheet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"त्रुटि: {e}")
                
