# Daily Diary Automation Script 📅

Automatically update dates in your UVT ICT Industrial Training Daily Diary Word document for 26 weeks.

---

## ⚠️ IMPORTANT - READ FIRST

### 🇺🇸 English:
**Before running this script:**
- The script **ONLY fills existing tables** - it does NOT create new pages or tables
- Your Word document (`Daily Dairy uvt ict.docx`) must already have **26 empty weekly tables**
- If your file only has 1 table, only Week 1 will be updated

### 🇱🇰 සිංහල:
**Script එක Run කිරීමට පෙර:**
- මෙම මෘදුකාංගය **දැනට තිබෙන වගු පමණක් පුරවයි** - අලුත් පිටු හෝ වගු සාදන්නේ නැහැ
- ඔබේ Word ලේඛනයේ (`Daily Dairy uvt ict.docx`) **සති 26ක හිස් වගු 26** කලින් තිබිය යුතුයි
- ෆයිල් එකේ එක වගුවක් තියෙනවා නම්, Week 1 විතරක් update වෙයි

---

## ✨ Features

- ✅ Automatically fills dates for 26 weeks (Nov 20, 2025 to May 19, 2026)
- ✅ Updates "Sunday" date in week headers
- ✅ Preserves existing formatting (Bold, Font Size)
- ✅ Marks **Saturdays & Sundays** as "WEEKEND" (Black, Bold, Centered)
- ✅ Marks **Holidays** with holiday name (Red, Bold, Centered)
- ✅ Smart logic: Holidays on weekends stay as "WEEKEND"

---

## 🚀 Installation & Usage

### Step 1: Install Python
Download and install Python from [python.org](https://www.python.org/downloads/)

### Step 2: Install Required Package
```bash
pip install python-docx
```

### Step 3: Prepare Your Document
1. Make sure your Word file has **26 empty weekly tables**
2. Rename it to: `Daily Dairy uvt ict.docx`
3. Place it in the same folder as `update_dates.py`
4. **Close the Word file** before running the script

### Step 4: Run the Script
```bash
python update_dates.py
```

### Step 5: Check Output
- A new file `Daily Dairy Final by A_L_E_X.docx` will be created
- Open it and verify the dates!

<img width="742" height="341" alt="image" src="https://github.com/user-attachments/assets/6ca02120-1547-40e9-af75-a556e4fa5c46" />

---

## 🔧 Configuration

You can customize dates and holidays by editing `update_dates.py`:

### Change Start Date:
```python
start_date_user = datetime(2025, 11, 20)  # Change this date
```

### Add/Remove Holidays:
```python
holiday_map = {
    (2025, 12, 25): "Christmas Day",
    (2026, 1, 1): "New Year's Day",
    # Add more holidays here
}
```
## ❓ Troubleshooting

**Problem:** "File not found" error
- ✅ Make sure `Daily Dairy uvt ict.docx` is in the same folder as `main.py`

**Problem:** Only Week 1 updated
- ✅ Your document needs 26 empty weekly tables before running

**Problem:** Dates not changing
- ✅ Make sure the Word file is **closed** before running the script

**Problem:** Formatting lost
- ✅ The script preserves existing formatting in date cells

---

**Happy Training! 🎉**
