## 📖 အသုံးပြုနည်းလမ်းညွှန် / User Guide

### 🇲🇲 မြန်မာဘာသာဖြင့် အသုံးပြုနည်းလမ်းညွှန်
1. **ပရိုဂရမ်အား စတင်ဖွင့်လှစ်ခြင်း:** Terminal သို့မဟုတ် Command Prompt တွင် `python main.py` ဟု ရိုက်နှိပ်၍ ပရိုဂရမ်ကို ဆောင်ရွက်ပါ။
2. **Word Template ဖိုင်ရွေးချယ်ခြင်း:** **Load Word Template** ခလုတ်ကို နှိပ်ပြီး မိမိအသုံးပြုမည့် ဇယား (Table Structure) ပါဝင်သော `.docx` ဖိုင်ကို ရွေးချယ်ပါ။
3. **Excel Data ဖိုင်ရွေးချယ်ခြင်း:** **Load Excel Data** ခလုတ်ကို နှိပ်ပြီး အချက်အလက်များ ပါရှိသည့် `.xlsx` ဖိုင်ကို ရွေးချယ်ပါ။ ထို့နောက် ပေါ်လာသော Dialog Box တွင် အသုံးပြုလိုသည့် Sheet ကို ရွေးချယ်ပါ။
4. **ကော်လံများ ချိတ်ဆက်ခြင်း (Mapping Rules):**
   - 'Word Columns (Available)' အကွက်ထဲမှ ခေါင်းစဉ်တစ်ခုကို ကလစ်နှစ်ချက် (Double Click) နှိပ်၍ ရွေးချယ်ပါ။ (၎င်းသည် 'Word Selected' ထဲသို့ ရောက်သွားပါမည်)
   - 'Excel Columns (Available)' အကွက်ထဲမှ ၎င်းနှင့် တွဲဖက်လိုသော အချက်အလက် ကော်လံကို ကလစ်နှစ်ချက်နှိပ်၍ ရွေးချယ်ပါ။ (၎င်းသည် 'Excel Selected' ထဲသို့ ရောက်သွားပါမည်)
   - ထို့နောက် **Add Rule** ခလုတ်ကို နှိပ်ပါ။ စည်းမျဉ်းအသစ်ကို 'Current Mapping Rules' အကွက်ထဲတွင် မြင်တွေ့ရမည် ဖြစ်သည်။
5. **ဖိုင်ထုတ်ယူခြင်း:** စည်းမျဉ်းများ အားလုံး သတ်မှတ်ပြီးပါက **🚀 GENERATE REPORT** ခလုတ်ကို နှိပ်ပါ။ ပရိုဂရမ်သည် စာမျက်နှာများကို အလိုအလျောက် ပေါင်းစည်းပေးသွားမည် ဖြစ်သည်။
6. **ရလဒ်ဖိုင်အား စစ်ဆေးခြင်း:** **📂 OPEN FOLDER** ခလုတ်ကို နှိပ်၍ `output` ဖိုင်တွဲအတွင်းရှိ `Final_Combined_Report.docx` ဖိုင်ကို ဖွင့်လှစ်ပြီး ရလဒ်အား စစ်ဆေးနိုင်ပါသည်။

---

### 🇺🇸 English User Guide
1. **Run the Application:** Open your Terminal or Command Prompt, type `python main.py`, and press Enter to launch the tool.
2. **Load Word Template:** Click the **Load Word Template** button and select your baseline `.docx` file that contains the target table layout.
3. **Load Excel Data:** Click the **Load Excel Data** button, select your `.xlsx` data source, and choose the specific sheet you want to process from the popup dialog.
4. **Map the Columns (Mapping Rules):**
   - Double-click a header from the 'Word Columns (Available)' list to move it to 'Word Selected'.
   - Double-click the corresponding column(s) from the 'Excel Columns (Available)' list to move to 'Excel Selected'.
   - Click the **Add Rule** button. The new mapping rule will appear in the 'Current Mapping Rules' panel.
5. **Generate the Document:** Once all rules are defined, click the **🚀 GENERATE REPORT** button. The application will process and automatically merge all pages seamlessly.
6. **View the Output:** Click the **📂 OPEN FOLDER** button to open the `output` directory and access your final consolidated report named `Final_Combined_Report.docx`.