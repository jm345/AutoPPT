AutoPPT - Automatic Presenatation Creator

Project Description:
Began using Google Slides API to try to automate the more laborious tasks in the business. With information downloaded from the image downloader (planning to extend to include more useful data such as product descriptions, pricing etc.) I can use that data to generate presentations quickly. Some presentations can go beyond 100 slides which means a lot of tedious work including trivial tasks such as text and picture alignment (CTRL+C -> CTRL+V rinse and repeat). 

Road Map

Current Stage:
With information contained in a .csv retrieved from a web scraper, can automatically generate a fresh presentation using
Google Slides API.
Currently parses the .csv for images, prices and product specifications.
Creates the final layout including appropriately formatted text, colours and fonts.

Target Stage:
Short Term:
Generate a new presentation each time the program is run.
Add shadows underneath the price tag.
Automatically add product titles into the presentation.

Long Term
Order the data in the .csv to prevent users from needing to reorder slides.
Allow users to download the presentations in multiple file formats, most importantly in .pdf and .ppt.
Create a simple GUI to make creating the presentations even easier.
Create intermediary slides such as title and category slides.
Allow users to enter custom data (margins, apply discounts, custom logos etc.) into the GUI to customise the presentation.
Migrate codebase to C# so I can create a desktop application and a mobile application using Xamarin.

The program takes approximately 1.7 seconds per slide (all the slides are identical in that they have a title, specs and an image). For perspectives sake, a previous presentation created manually using data extracted from the same website took me approximately 9 hours to complete 161 slides (11 hours vs around 4 and a half minutes).
