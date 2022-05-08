# WoocommerceToGoogleSheets

# Project Description

This code is intended for everyone using the woocommerce plugin in wordpress. 
It pulls all products from woocommerce and adds them into your googlesheet.
This helps if you want to use your data for further apis. 

In my case, the spreadsheet was used to easily import all products into ebay. 
In the next step the data will be used to easily generate a feed for google merchant, so the products from the wordpress shop can also be found in googles shopping search.

# How to install and run the project

You need to change the following things in the python code to make it working for you. 

Woocommerce setup: 
- You will need a working wordpress site, with the woocommerce plugin installed.
- You need a consumer key and consumer secret. There are several guides on how to create them.
- Change "wcapi_consumer_key = os.environ.get("wcapi_consumer_key")" to your consumer key. For securitys sake i saved it in an enviromental variable. 
- Change wcapi_consumer_secret = os.environ.get("wcapi_consumer_secret") to your secret key. For securitys sake i saved it in an enviromental variable. 
- in the wcapi = API(...) call change the url to your homepage. 

Googlesheets setup:
- You will need a google development account
- check the documentation or google for Writing to Google Sheets API using service account OR
- i used the gspread library to help me with all of that. they have a usefull documentation about the authorization as well:
  - https://docs.gspread.org/en/latest/oauth2.html
- you need to save your credentials in a json file
- link to that json file  in the code here: gc = gspread.service_account(filename='keys.json')
- change the SAMPLE_SPREADSHEET_ID to the google spreadsheet ID (the ID, not the whole link) you want to use. 

You can decide which categories to pull from woocommerce into your spreadsheet. print(product) in the while loop to get an idea which options woocommerce gives you. 
Make sure to name your headers accordingly. 

You can also push changes you made in your spreadsheet back to woocommerce with the push_to_woocommerce() function. 
I'm not using it myself so it's greyed out and doesn't get called, but the option exists if needed. 

