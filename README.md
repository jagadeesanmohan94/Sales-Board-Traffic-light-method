
Dashboard Link: https://capgemini-my.sharepoint.com/personal/jagadeesan_mohan_capgemini_com/Documents/Desktop/UDEMY%20cOURSE/Data%20Analysis%20Projects/Excel%20Project/Traffic%20Model%20-%20Dashboard.xlsb?web=1

#Calculation Board


![Image](https://github.com/user-attachments/assets/789f49cc-397f-4429-aa3d-0fda232242ae)

OFFSET (range, rows, columns, [height], [width] )				
Parameters or Arguments				
Range: The starting range from which the offset will be applied.				
Rows: The number of rows to apply as the offset to the range. This can be a positive or negative number.				
Columns: The number of columns to apply as the offset to the range. This can be a positive or negative number.				
Height (Optional):. It is the number of rows that you want the returned range to be. If this parameter is omitted, it is assumed to be the height of range.				
Width (Optional): It is the number of columns that you want the returned range to be. If this parameter is omitted, it is assumed to be the width of range				
				
				
				
MATCH(lookup_value, lookup_array, [match_type])				
Parameters or Arguments				
lookup_value    Required. The value that you want to match in lookup_array. For example, when you look up someone's number in a telephone book, you are using the person's name as the lookup value, but the telephone number is the value you want.				
The lookup_value argument can be a value (number, text, or logical value) or a cell reference to a number, text, or logical value.				
lookup_array    Required. The range of cells being searched.				
match_type    Optional. The number -1, 0, or 1. The match_type argument specifies how Excel matches lookup_value with values in lookup_array. The default value for this argument is 1.				
				
#Validation  for months				
				
Conditional format for the month column for blink				
				
Total sales for each month, if incase we need check jan to feb or march not full year sum amount to design a formula for that is 				
				
	SUM(OFFSET(C4,0,0,1,MATCH($B$10,$C$3:$N$3,0)))			
				
				
Low	R	IF(C3<=$C$9,"=","")		
Median	Y	IF(C3<$C$10,IF(C3>$C$9,"=",""),"")		
High	G	IF(C6>=$C$10,"=","")		
				
				
Trafic light image attached and change the font size into webdings then copy past to that location				
				
Change the Location into that formula sheet cells 				
				
Adjusting with the dashboards and the select each rows and shows sparkline and bar charts				



#Shifting the Validation month and Dash board details changed

![Image](https://github.com/user-attachments/assets/706da5ab-6937-4ee6-afd2-e279fff37aac)

#Jan Month Dashboard


![Image](https://github.com/user-attachments/assets/1e8621ce-5dfd-4d14-9ca2-3dcdea908093)

#Feb Month Dashboard

![Image](https://github.com/user-attachments/assets/bde95a47-2557-4e46-8d23-36020bae28ef)

#Mar Month Dashboard

![Image](https://github.com/user-attachments/assets/709462b9-b9dc-469d-b9da-cb0ad8497344)

#Apr Month Dashboard

![Image](https://github.com/user-attachments/assets/66db9f1b-38e6-436a-9389-663d4d4705c6)

#Dec Month Dashboard

![Image](https://github.com/user-attachments/assets/a53ee84d-6827-46fe-b9db-26d2338289ff)


📊 Detailed Summary of the Traffic Sales Dashboard
The dashboard visualizes sales performance for four Microsoft Office products—Windows, Word, Excel, and PowerPoint—across all twelve months (Jan–Dec). It includes monthly sales data, total yearly sales, traffic‑light indicators, sparklines, and bar charts to help understand trends.

1. ⭐ Total Yearly Sales Overview	
	
The TOTAL SALES column shows how each product performed overall in the year:	
	
Product	Total Sales
Windows	1639
Word	2203
Excel	1888
PowerPoint	2768
	
Key Insight	
	
PowerPoint is the top-selling product with 2768 units, significantly outperforming all others.	
Windows has the lowest total sales at 1639 units.	
2. 📈 Monthly Sales Trends
Windows

Sales fluctuate moderately, with a major spike in November (260).
Lowest month: January (70).

Word

Strong and consistent performance throughout the year.
Best month: December (185).
Lowest month: March (120).

Excel

Stable upward trend with periodic spikes.
Best month: November (139).
Lowest month: January (90).

PowerPoint

Highest and most consistent growth.
Several peak months: July (259), October (217), November (239).
Lowest month: January (279) — still higher than others.


3. 🚦 Traffic Signal Indicators
The dashboard uses stoplight-style indicators to visually represent product performance.

Windows → Red = Low performance
Word → Amber/Yellow = Moderate performance
Excel → Yellow = Moderate performance
PowerPoint → Green = High performance

Interpretation

PowerPoint is performing well above expectations.
Windows requires attention due to weak performance.
Word and Excel show mid-range performance.


4. 📉 Sparklines (Mini Trendlines)
Each product has a sparkline showing its sales trend across 12 months.

Windows – fluctuating, no strong upward movement.
Word – somewhat upward trend with small dips.
Excel – steady overall, moderate peaks.
PowerPoint – strong pattern with multiple high points.

These lines help compare patterns quickly without detailed numbers.

5. 📊 Bar Charts
The dashboard includes bar-style visual bars (mini bar charts) showing variation in monthly sales:

Windows & Excel show more variation month to month.
Word & PowerPoint show consistently higher sales bars.

These give a visual sense of volume and stability.

📌 Final Insights

PowerPoint is the strongest performer across the year.
Windows lags behind significantly and may require review of marketing or sales strategy.
Excel and Word are stable performers—neither exceptionally high nor concerningly low.
Visualization tools (traffic lights, sparklines, bar charts) clearly highlight performance and trends in a clean, intuitive way.

