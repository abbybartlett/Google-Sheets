# Google-Sheets
Google Apps Scripts from 2017 - Created to aid with my work in creating data reports for Hopscotch clients.

- [`arrange.js`](/arrange.js) is a custom Google Apps Script I created in 2017 to programatically organize, manipulate and apply formatting to a weekly data report that I managaed and maintained to monitor usage and performance of the 200+ native mobile applications I was working on the time at Hopscotch (RIP).
- This scrpit transforms a raw data table in CSV format into an informative and visually appealing report built in Google Sheets. We would then distribute these polished reports to our customer success team to share and review with our clients.
- - example of how the data was exported from our database: https://docs.google.com/spreadsheets/d/1DU60xgTVLqvXHi0GT9zNrwBMT2eKHGol5f9ArFmH7r8/edit?gid=1576233508#gid=1576233508 
- i think i was improving the script here: ***Arrange Template (standalone)**:  https://script.google.com/home/projects/1J3Ln44TEfS3qRURHBj6QONHMs44n7xtExMenY_gExdyvoDvZsu2TFD2L/edit 

## Related Documents
- **UsageReport TEMPLATE**: https://docs.google.com/spreadsheets/d/1rpIvHN0vKI8_cYb7CxQDcAFij170csIinrfzoI2gLr0/edit?usp=sharing
- in the UsageReport, open **Extensions** > **Apps Script** to find the `code.gs` script that imports - [`arrange.js`](/arrange.js) from GitHub into Google Apps Script
- - I have [`code.gs`](/code.gs) bound to the **UsageReport TEMPLATE** -- im not sure if this is the best method, but it always worked for me.

# Connecting the Google Apps Scripts together?
_These steps below are unverifeid and need to be debuged/checked for compatiblity with latest instaces of Google Apps Script_

- You will need to use the **Script ID** that is provided in the comment in line 3 to connect `arrange.js` to this Apps Script. 
- click "+" button to **Add a Library** and paste in the script ID <img width="1832" height="1527" alt="Screenshot 2025-07-16 at 4 36 47 AM" src="https://github.com/user-attachments/assets/4068ebdb-9610-47d7-a781-40ae8df93559" />
- click **Look Up**
- click "Add" button <img width="546" height="513" alt="Screenshot 2025-07-16 at 4 39 53 AM" src="https://github.com/user-attachments/assets/c88bcc68-630b-437f-bd78-4c1d6aa2cf2c" />


## In Summary
Im starting to think that maybe I only used GitHub as a separate source of truth for my scripts and would copy/paste them manually into my Drive files.. instead of importing in a fancy way.

- it all worked out nicely but the library connections have broken over the years.
- original files can be found here: My Drive > Hopscotch > Data & Dashboards > Usage Report: https://drive.google.com/drive/folders/127zWCFLnWMWf2ARz-S4Jh1eKztdZV1ve 
