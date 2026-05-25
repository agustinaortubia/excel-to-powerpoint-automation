# 🚀 From 45 Minutes of Copy-Pasting to a 5-Minute Automated Workflow

## 📌 The TL;DR
Let's be real: creating periodic market intelligence reports manually is a drain. It used to take me 45 minutes of tedious copying, pasting, resizing, and positioning charts from Excel into PowerPoint. 

So, I built a VBA pipeline to do it for me. Now? **It takes 5 minutes.** That's an 88% reduction in busywork, with zero manual formatting errors and perfect pixel alignment.

*(Note: The original corporate Excel models are proprietary, so this repo contains the core VBA engine (`ppt_export_macro.vba`) and the architectural breakdown).*

## 🏗️ How it actually works: The Control Sheet
Instead of hardcoding a messy script, the macro reads from a clean `Control2` sheet in Excel that acts as mission control. Any user can adjust the presentation without touching a single line of code by mapping:

* **Slide #:** Where it goes in PowerPoint.
* **Object Name:** The specific Excel asset (Charts or Named Ranges).
* **Target_Top & Target_Left:** Exact placement coordinates.
* **Target_Height & Target_Width:** Precise resizing dimensions.

## 🧠 The Clever Bits (Handling Data Complexity)
The pipeline doesn't just copy and paste blindly. It evaluates the content first to handle the messy reality of dynamic data:

1. **Smart Trimming (Overlap Tables):** Data volume changes every week. The script scans down Column A and automatically trims the capture area exactly where the data ends. No blank spaces exported.
2. **Dodging the Junk:** For complex sheets, the macro uses dual-stop conditions (like detecting `"Sum of Available"`). This prevents it from accidentally taking screenshots of backend pivot tables.
3. **Aspect-Ratio Cloning:** The biggest issue with pasting Excel tables into PPT is distortion. To fix this, the code calculates the exact Height/Width ratio required by the PowerPoint slide, dynamically recalculates the Excel column widths, takes the screenshot, and then reverts Excel to normal. Result? Perfectly scaled tables without pixelation.

## 🛠️ Tech Stack
* **VBA:** Core automation engine and cross-application control.
* **Excel:** Dynamic ranges, data manipulation, and configuration hub.
* **PowerPoint:** Automated slide population.
