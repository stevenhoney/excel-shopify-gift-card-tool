# Excel Shopify Gift Card Tool
---
<small>Credit for actually getting this to work goes to [this chap on Upwork](https://www.upwork.com/freelancers/~018f69f455ef453569) who is the guy you should contact if you want to extend the functionality of the tool or create something similar for another Shopify API.</small>

### What
A macro enabled Excel workbook (XLSM) which will create a gift voucher per line in a Shopify Plus store using entered details. Connected to Shopify via a private app which the user must first create (see Getting Started below).

### Why
I half finished this to help me migrate gift cards from our legacy system when moving to Shopify Plus, then realized it's really useful to be able to bulk create and assign gift cards in lots of use cases, so decided I'd make a portable tool to share.

# Getting Started

### Create Your Private App (One Time Process)

1. In your Shopify admin navigate to Apps, then scroll to bottom of the page and click "Manage Private Apps"
2. Click on "Create New Private App" at top right of screen
3. Give your new app any name you like
4. Clear any default permissions that are set (usually Read access to Products which is not required here)
5. Scroll down the inactive permissions list to Gift Cards and select Read and Write, this is the only permission your app will require.
6. Hit Save.
7. You will now have an API Key and Password which can be entered in the corresponding cells on the Settings tab of the workbook.

### Using the Tool

##### Settings
All that is required are your Private App's API Key and Password from the step above and your xxxx.myshopify.com domain (just the xxxx bit).

##### Inputting Cards to Be Generated
Five columns are used to form the body of the post to the API:

###### Value
The tool will create a voucher for every row in the Value column with a something entered. Technically this is the only required column, although it is hard to imagine a use case for just creating batches of Gift Cards with randomly generated codes.

###### Note
A long text type field, notes are visible when viewing the gift card in the Shopify admin. Unless you have some very specific customization(s) this is for internal use only. Linebreaks and basic formatting are preserved, HTML does not work.

##### Code
Minimum 8 characters, max 20. Alphanumeric (a-z,0-9) only, must be globally unique.

##### Template





