# Nándor Magyar

**Automation & Excel Developer · Food Industry Engineer · Agricultural Commodity Trading**
Malta · magyarnana97@gmail.com

---

I'm a food industry engineer with a background in agricultural commodity trading. Alongside trading, I built tools that took the repetitive manual work off the desk.

A lot of time in trading goes into routine tasks. Opening emails one by one, copying their contents into Excel, pulling reports from SAP, sending the same kind of delivery notice to every customer. I automated these with Python and VBA. A macro reads the inbox and imports the data, every partner gets a ready-to-check email about yesterday's pickups, and the sales report that used to take an hour every morning now runs on its own. The tools are not rigid. They adapt to the date, the product, the customer and other changing details.

I'm now based in Malta and looking for a role where I can help people replace manual work with automation and Excel development. My aim is always the same. People should get to the point where their experience is really needed much faster, and spend that time on decision making instead of copy-paste.

---

## What I Work With

![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-150458?style=flat&logo=pandas&logoColor=white)
![VBA](https://img.shields.io/badge/Excel_VBA-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Outlook](https://img.shields.io/badge/Outlook_Automation-0078D4?style=flat&logo=microsoft-outlook&logoColor=white)
![SAP](https://img.shields.io/badge/SAP_GUI_Scripting-0FAAFF?style=flat&logo=sap&logoColor=white)
![Git](https://img.shields.io/badge/Git-F05032?style=flat&logo=git&logoColor=white)

**Automation:** Python (`pandas`, `openpyxl`, `win32com`, `tkinter`, `folium`) · Excel & Outlook VBA · SAP GUI Scripting API
**Domain:** Agricultural commodity trading · Logistics operations · B2B pricing workflows · FX normalization

---

## Projects

### 🗺️ [SAP sales report and logistics map](https://github.com/NandesHungarian/SAP_AgriTrade_Automation)

Every morning the sales report had to be pulled from SAP, converted to EUR with the day's exchange rates and formatted for management. This took 45 to 60 minutes. Now a Python script does it in a single run. It logs into SAP, exports the report, converts the prices, asks for any missing freight cost in a small popup window and builds the summary table. At the end it can draw an interactive map showing where the goods are going.

![Interactive logistics map](https://raw.githubusercontent.com/NandesHungarian/SAP_AgriTrade_Automation/main/docs/images/logistics_map.jpg)

### 📈 [Historical Trade Data Analyzer](https://github.com/NandesHungarian/Automation-Portfolio/blob/main/Python/Historical_Trade_Data_Analyzer.py)

When sales need to be analysed over several years, this tool comes in. To recalculate past sales from HUF to EUR or USD accurately, it needs the forward rates that were valid on each day. These are stored in hundreds of daily Excel files. The script opens all of them, matches every contract to the right day's rates and converts the prices. pandas keeps this fast even with years of data.

### ⚙️ [Everyday Excel and Outlook automation](https://github.com/NandesHungarian/Automation-Portfolio/tree/main/VBA)

Smaller VBA macros that replace the manual tasks of the day. Together they save more than 20 hours a month.

- [Finds today's report email](https://github.com/NandesHungarian/Automation-Portfolio/blob/main/VBA/Email_Attachment_Importer.vba) in Outlook with one click and pastes its Excel attachment into the right sheet. Easy to reuse for other daily emails
- [Prepares a daily email for each partner](https://github.com/NandesHungarian/Automation-Portfolio/blob/main/VBA/Aviso_Automation.vba) showing what they collected and what is still waiting, using their own template. Usually 5–15 partners a day. The emails open as drafts for a final check, saving about an hour a day
- [Sends the daily price table](https://github.com/NandesHungarian/Automation-Portfolio/blob/main/VBA/FCA_Price_Indication_Emailer.vba) as an image in an email and saves the prices to a log sheet, so the price history is always ready for a chart
- [Splits the weekly report](https://github.com/NandesHungarian/Automation-Portfolio/blob/main/VBA/Regional_Coverage_Report_Distributor.vba) and emails each regional colleague their part
- [Merges data files](https://github.com/NandesHungarian/Automation-Portfolio/blob/main/VBA/Incoming_Data_Consolidator.vba) from regional teams into one clean table without duplicates
- [Calculates weighted average prices](https://github.com/NandesHungarian/Automation-Portfolio/blob/main/VBA/Period_Average_Calculator.vba) for any chosen period
- [Flags contracts](https://github.com/NandesHungarian/Automation-Portfolio/blob/main/VBA/Tiered_Pricing_and_Deviation_Tracker.vba) booked outside the approved price grid

---

*All company-specific data in public repositories has been anonymized.*
