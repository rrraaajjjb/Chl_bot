import logging
import io
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# Logging setup
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# ⚠️ यहाँ अपना असली टेलीग्राम बॉट टोकन डालें
BOT_TOKEN = "8879189917:AAEr9YsDVpfd2_R7L3Cw-zM6Uws83CgH5og"

user_data_store = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    user_data_store[user_id] = {'sheet1': [], 'sheet2': []}
    await update.message.reply_text(
        "👋 नमस्ते! मैं आपका Excel Automation Bot हूँ।\n\n"
        "📁 अपनी Sheet1 and Sheet2 वाली Excel (.xlsx) या CSV फाइलें मुझे एक-एक करके भेजें।\n"
        "सभी फाइलें भेजने के बाद, **/merge** कमांड टाइप करें। मैं तय फॉर्मेट में फाइल तैयार कर दूंगा!"
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    file = await update.message.document.get_file()
    filename = update.message.document.file_name.lower()
    
    if user_id not in user_data_store:
        user_data_store[user_id] = {'sheet1': [], 'sheet2': []}
        
    await update.message.reply_text(f"⏳ '{update.message.document.file_name}' को रीड किया जा रहा है...")
    file_bytes = await file.download_as_bytearray()
    
    try:
        if filename.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(file_bytes))
        elif filename.endswith('.xlsx') or filename.endswith('.xls'):
            df = pd.read_excel(io.BytesIO(file_bytes))
        else:
            await update.message.reply_text("❌ कृपया केवल Excel (.xlsx) या CSV फाइल ही भेजें।")
            return

        df.columns = [c.lower().strip() for c in df.columns]

        if "sheet1" in filename or "netprice" in df.columns:
            user_data_store[user_id]['sheet1'].append(df)
            await update.message.reply_text("✅ Sheet1 (प्राइस डिटेल्स) मिल गई है।")
        elif "sheet2" in filename or ("quantity" in df.columns and "unit_name" in df.columns):
            user_data_store[user_id]['sheet2'].append(df)
            await update.message.reply_text("✅ Sheet2 (क्वांटिटी डिटेल्स) मिल गई है।")
        else:
            await update.message.reply_text("⚠️ फाइल की पहचान नहीं हो सकी। फाइल के नाम में 'Sheet1' या 'Sheet2' होना चाहिए।")
            
    except Exception as e:
        await update.message.reply_text(f"❌ फाइल एरर: {str(e)}")

async def merge_files(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    
    if user_id not in user_data_store or (not user_data_store[user_id]['sheet1'] and not user_data_store[user_id]['sheet2']):
        await update.message.reply_text("❌ पहले फाइलें भेजें, फिर /merge चलाएं।")
        return
        
    s1_list = user_data_store[user_id]['sheet1']
    s2_list = user_data_store[user_id]['sheet2']
    
    if not s1_list or not s2_list:
        await update.message.reply_text("❌ प्रोसेस करने के लिए Sheet1 और Sheet2 दोनों फाइलें होनी जरूरी हैं।")
        return
        
    await update.message.reply_text("⚙️ डेटा मर्ज किया जा रहा है... कृपया रुकें।")
    
    try:
        df_s1_all = pd.concat(s1_list, ignore_index=True)
        df_s2_all = pd.concat(s2_list, ignore_index=True)
        
        lookup = df_s1_all[['vendor_name', 'item_name', 'pack_type', 'netprice']].drop_duplicates()
        
        cols_to_drop = [c for c in ['vendoraddress', 'unitaddress'] if c in df_s2_all.columns]
        df_s2_clean = df_s2_all.drop(columns=cols_to_drop)
        
        merged_df = pd.merge(df_s2_clean, lookup, on=['vendor_name', 'item_name'], how='left')
        merged_df = merged_df.rename(columns={'netprice': 'rate'})
        merged_df['amount'] = merged_df['rate'] * merged_df['quantity']
        
        final_columns = ['unit_name', 'item_name', 'pack_type', 'quantity', 'rate', 'amount', 'vendor_name']
        final_df = merged_df[final_columns]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name="Automated_Report")
        output.seek(0)
        
        await update.message.reply_document(
            document=output, 
            filename="Automated_Master_Sheet.xlsx",
            caption="🎉 आपकी फाइनल फाइल तैयार है!"
        )
        user_data_store[user_id] = {'sheet1': [], 'sheet2': []}
        
    except Exception as e:
        await update.message.reply_text(f"❌ मर्ज एरर: {str(e)}")

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("merge", merge_files))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    print("बॉट चालू है...")
    app.run_polling()

if __name__ == '__main__':
    main()
  
