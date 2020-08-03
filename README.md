# Excel Shopify Gift Card Tool
---
<sub>Credit for actually getting the VBA bit to work reliably goes to [this chap on Upwork](https://www.upwork.com/freelancers/~018f69f455ef453569), this is the guy you should contact if you want to extend the functionality of the tool or create something similar for another Shopify API.</sub>

### What
A macro enabled Excel workbook (XLSM) which will create a gift voucher per line in a Shopify Plus store using entered details (you must have a Plus level subscription to use the Gift Card API). Connects to Shopify via a private app which the user must first create (see Getting Started below).

### Why
I half finished this to help me migrate gift cards from our legacy system when moving to Shopify Plus, then realized it's really useful to be able to bulk create and assign gift cards in lots of use cases, so decided I'd make a portable tool to share.

## Getting Started

### Create Your Private App (One Time Process)

1. In your Shopify admin navigate to Apps, then scroll to bottom of the page and click "Manage Private Apps"
2. Click on "Create New Private App" at top right of screen
3. Give your new app any name you like
4. Clear any default permissions that are set (usually Read access to Products which is not required here)
5. Scroll down the inactive permissions list to Gift Cards and select Read and Write, this is the only permission your app will require.
6. Hit Save.
7. You will now have an API Key and Password which can be entered in the corresponding cells on the Settings tab of the workbook.

## Using the Tool

#### Settings
All that is required are your Private App's API Key and Password from the step above and your xxxx.myshopify.com domain (just the xxxx bit).

#### Inputting Cards to Be Generated
Five columns are used to form the body of the post to the API:

  - ##### Value
   The tool will create a voucher for every row in the Value column with a something entered. Technically this is the only required column, although it is hard to imagine a use case for just creating batches of Gift Cards with randomly generated codes.
  - ##### Note
   A long text type field, notes are visible when viewing the gift card in the Shopify admin. Unless you have some very specific customization(s) this is for internal use only. Linebreaks and basic formatting are preserved, HTML does not work.
  - ##### Code
   Minimum 8 characters, max 20. Alphanumeric (a-z,0-9) only, must be globally unique.
  - ##### Template
  Like `gift_card.birthday.liquid` Shopify gift card liquid templates are effectively standalone HTML pages, a template with a "Happy Birthday" message is an obvious example.
  - ##### Shopify Customer ID
  Like `` This is the number after the last / in the URL when viewing the customer record in the Shopify admin. Unfortunately Shopify does not include this in their default customer export, so a 3rd party app is needed in order to export Customer IDs to be used to assign gift cards to customers. I use and highly recommend [Excelify.io](https://excelify.io/) but there are [lots of alternatives](https://apps.shopify.com/search?q=csv+export#).
  
  When you assign a gift card to an existing customer it is sent to them via email immediately on creation, if you have their mobile/cell number in Shopify they will also be delivered via SMS.

## Security
You really shouldn't download and run XLSM files from the internet ever. It would be pretty easy for me to have added a line in here that sends me gift card codes for your store. The VBA code that is used is [here in the SRC folder]() and ideally you or someone in your oranization who understands it should check the code and create your own local version based on it.

As a minimal security step the VBA strips the values from the cells containing API credentials on both file open and close, it is slightly annoying to input each them each time, but far less annoying than your disgruntled intern going rogue with the gift card api on your store!

You can download the XLSM file from this repository and it will work fine, but you shouldn't and I told you not you.
