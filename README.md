# ItemUploadTool


EPIC PIPING
ITEM UPLOAD TOOL
BY JARED HICKS
July 25, 2018

This is a live / deployed Application being used by Epic Piping, its primary purpose is to speed up any processes that can be improved upon, while also decreaseing the risk for errors.

This documentation is not up to date as of 2-15-2019, updates to the documentation will be made shortly.

Below is an outline of each tab within the application, and some of the processes it performs.


**1.	Tab 1 - Fittings**
  - This Tab is for the processing and creation of all Partcodes not related to supports and originating from any source.	
- Checks Partcode on the fly to an Item Master table that’s updated every 15 minutes to see if exists.
-	If Partcode exists, click green status text to import all segments.
-	Description splits at 30 chars to ensure segment length is not exceeded.
-	Material helper - Click Blue text, select material within popup, auto populates Material field.
-	GL Class helper - Click Blue text, select GL Class option 3 and 4, auto populates GL Class field.
-	All segments must be populated to progress, weight and SA must be over 0, and SA must be less than weight.
-	Template is automatically selected based on Subcomm.
-	Partcode segments are automatically placed based on Subcomm.
-	Button; Import Missing, Imports missing Partcodes from auto spec building tool into dropdown list with or without descriptions based on checkbox.
-	Right click on ‘Import Missing’ for an option to import Partcodes from clipboard into dropdown list.
-	When Partcodes in dropdown list are selected the proper fields will fill with the create segments.
-	Button; 'Add Partcode' or 'Add and Clear text boxes' both buttons add the Partcode to be created in the correct structure to a DataGrid table and either clear and reset the text fields or leave them filled.
-	Right click on 'Add' or 'Add_Clear' for an option to perform buttons task but with an Action of 04 to revise a Partcode with the added function of either clearing and resetting the text fields or leaving them filled.
-	Button; Undo Sort, undoes the sorting that is given to the DataGrid table when the copy button is pressed.
-	Button; Undo Add, Removes the last added Partcodes from the DataGrid Table, can be used multiple times.
-	Button; Reset Fields, clears all text boxes of info.
-	Button; Clear, data clears the DataGrid table of any data.
-	Button; Copy, Sorts and copies the DataGrid tables info and allows you to paste it within JDE.
-	Button, Right Click, Copy, Copy as Req, Copies the first instance of every Partcode to the clipboard.


**2.	Tab 2 - Supports**
	- This tab is for the processing of all Support Partcodes.
-	Graphical arrow between Description and Partcode is clickable and converts the Description String into a Partcode (strips all special characters out and prefixes the string with an 'S').
-	Description splits at 30 chars to ensure segment length is not exceeded.
-	Material helper - Click Blue text, select material within popup, auto populates Material field.
-	GL Class helper - Click Blue text, select GL Class option 3 and 4, auto populates GL Class field.
-	All segments must be populated to progress, weight and SA must be over 0, and SA must be less than weight.
-	TO-BE-ADDED -- Checks Partcode on the fly to an Item Master table that’s updated every 15 minutes to see if exists.	-	Button; Undo Sort, undoes the sorting that is given to the DataGrid table when the copy button is pressed.
-	Button; Undo Add, Removes the last added Partcodes from the DataGrid Table, can be used multiple times.
-	Button; Reset Fields, clears all text boxes of info.
-	Button; Clear data, clears the DataGrid table of any data.
-	Button; Copy, copies the DataGrid tables info and allows you to paste it within JDE.
-	Button, Right Click, Copy, Copy as Req, Copies the first instance of every PartCode to the clipboard.
-	Button, Import, imports a list of PartCodes stored in the clipboard. When one is selected the it is placed in the Partcode field.
-	Checkbox, Req Builder, adds a new Data Grid to the bottom in the format of a Requisition, while adding new textboxes for Qty, Job number, and Reference.


**3.	Tab 3 - Partcodes / Breakout**
	- This tab is to aid in checking if Partcodes exist on a small or large scale, and for creating a Partcode Breakout which is used to check the accuracy of newly coded Partcodes before they are submitted into JDE.
-	Button; ‘Paste and check’ from clipboard a list of Partcodes, this will run a quick check to see if they exist and output the description.
-	Button; ‘Paste and Breakout’ from clipboard a list of Partcodes, this will run a quick breakout of the Partcodes and output all the segments in the DataGrid table on the right,
-	Right Click on the 'Paste and check' or 'Paste and Breakout' button to import from the missing Partcodes dropdown list on Tab1
-	Button 'Select All and Copy' will copy the DataGrid tables output for pasting in excel.
-	(note for above feature, make sure you have escaped out of the excel copy (the moving green dashes around a field of copied info) before pressing this button or it does not work.

**4.	Tab 4 - BOM Lookup**
	- This tab is to aid in Short report evaluations, a way to view the spool take off by only having the barcode, or job-spool number. The return info is pulled directly from the MTO DB.
-	to be added.

**5.	Future Adaptations / To be added**
	-	These are potential future adaptations I would like to add into the software assuming time is allocated to do so. I strongly believe each one of these features will greatly reduce the amount of time needed to perform our daily tasks.	
-	Partcode Description builder; drop downs / typing fields with auto complete to select segments and once complete the description is built.
-	(sub feature for above) Options to batch generate parts between size ranges.
-	Partcode checker; checks segment description against client description for matches.
-	Partcode builder; checks client description for strings and auto generates the Partcode for them, with options to have multiple settings / per job settings.
-	Build in a live weight and surface area table in SQL and have the info pulled in correctly when building new Partcodes / with option to submit new weights and surface area if one doesn’t exist.




This program shall stick to my personal philosphy of 70/30, where 70% of the programs functions shall be imediately recognizeable  without proper instruction, the other 30% shall be in a format that remains the same over the entire software, an example would be, all buttons that contain an Asterix have a right click option enabled. 

All additions / new functions shall be tested for a few weeks by my self and other like minded individuals before publishing to the company.

![IUT 1](https://user-images.githubusercontent.com/32394719/177992240-98a60940-9151-4795-87a9-089c4be11b5a.png)
![IUT 2](https://user-images.githubusercontent.com/32394719/177992241-4cce2c05-26cb-44ed-afcc-b979c66206d6.png)
![IUT 3](https://user-images.githubusercontent.com/32394719/177992242-da7c47ec-d99b-4388-a4a5-1801af496b8b.png)
![IUT 4](https://user-images.githubusercontent.com/32394719/177992243-a99bd063-f803-4ab5-96d9-026b2f6be21e.png)
![IUT 5](https://user-images.githubusercontent.com/32394719/177992233-ff89c139-97e2-4399-bb11-2583734e2509.png)
![IUT 6](https://user-images.githubusercontent.com/32394719/177992235-25e0c7e7-c0a5-4c02-b079-9b98bec21215.png)
![IUT 7](https://user-images.githubusercontent.com/32394719/177992238-fe2e9f2b-242b-4bd8-abe6-896d8ee76055.png)
