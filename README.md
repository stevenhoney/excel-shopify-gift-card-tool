# Excel Shopify Gift Card Tool
---

### What
A macro enabled Excel workbook (XLSM) which will create a gift voucher per line in a Shopify Plus store using entered details (you must have a Plus level subscription to use the Gift Card API). Connects to Shopify via a private app which the user must first create (see Getting Started below).

### Why
I half finished this to help me migrate gift cards from our legacy system when moving to Shopify Plus, then realized it's really useful to be able to bulk create and assign gift cards in lots of use cases, so decided I'd make a portable tool to share.

## Getting Started

### Create Your Private App (One Time Process)

1. Follow the [instructions here](https://help.shopify.com/en/manual/apps/app-types/custom-apps) to create a custom app.
2. Under API permissions you will need Gift Cards - select Read and Write, these are the only permissions your app will require.
4. You will now have an API Token to use in the Excel tool - note this down somewhere, it is only shown once (you can uninstall then reinstall your custom app to geenrate a new token any time).

## Using the Tool

#### Settings
All you need is the API Token from the step above and your xxxx.myshopify.com domain (just the xxxx bit).

<img src="https://github.com/stevenhoney/excel-shopify-gift-card-tool/blob/master/Screenshot-1.png" width="400px" />

(No, I haven't uploaded valid credentials in a screenshot here, don't worry!)

#### Inputting Cards to Be Generated
Six columns are used to create each gift card and form the body of each post to the API, all except the Value column are **optional**:

![Screenshot-2]

  - ##### Value
   Decimal number, do not include currency symbols, do no change the cell formatting. The tool will create a voucher for every row in the Value column with a value entered. Technically this is the only required column, although it is hard to imagine a use case for just creating batches of Gift Cards with randomly generated codes.
  - ##### Note
   A long text type field, notes are visible when viewing the gift card in the Shopify admin. Unless you have some very specific customization(s) this is for internal use only. Linebreaks and basic formatting are preserved, HTML does not work.
  - ##### Code
   Minimum 8 characters, max 20. Alphanumeric (a-z,0-9) only, must be globally unique.
  - ##### Template
  Like `gift_card.birthday.liquid` Shopify gift card liquid templates are effectively standalone HTML pages, a template with a "Happy Birthday" message is an obvious example. Note that the template suffix is an available variable in your gift card notification email template, so if you say wanted to run a promotional gift card send to a segment of customers you can set up a corresponding template and then include specific messaging in the notification email, e.g. if you created a template like gift_card.my-cool-promo.liquid then within the email template you can use ```liquid {% if gift_card.template_suffix == "my-cool-promo" %} //Promo specifc content goes here... {% endif %} ```
  - ##### Shopify Customer ID
  Like `3363246964811` This is the number after the last / in the URL when viewing the customer record in the Shopify admin. 
  ![Screenshot-3]
Unfortunately Shopify does not include this in their default customer export, so a 3rd party app is needed in order to export Customer IDs to be used to assign gift cards to customers. I use and highly recommend [Matrixify](https://matrixify.app/) but there are [lots of alternatives](https://apps.shopify.com/search?q=csv+export#).

  - #### Expiry Date (YYYY-MM-DD)
  An expiry date for the created gift card, must be in the YYYY-MM-DD format, do not change the Excel cell format to a Date type, it is intentionally a Text type column and needs to be so.
  
When you assign a gift card to an existing customer it is sent to them via email immediately on creation, if you have their mobile/cell number in Shopify they will also be delivered via SMS.
  
## API Rate Limiting
Plus stores have a 20 request per second limit on REST API requests, to ensure we stay under that the VBA script has a 65 millisecond sleep between requests built in, so makes a maximum of 15.4~ requests per second. In reality becasue there's time taken between the POST and response it's never quite this fast, but you can expect say a 1000 gift cards to take comfortably less than 2 minutes to create.

## Security
You really shouldn't download and run XLSM files from the internet, ever. It would be pretty easy for me to have added a line in here that sends me gift card codes or API tokens for your store (I didn't). The VBA code that is used is uploaded here as a .bas file and ideally you or someone in your organization who understands VBA a little should check the code and create your own local version based upon it.

As a minimal security step the VBA strips the values from the cells containing the API Token on both file open and close, it is slightly annoying to input them each time, but far less annoying than someone going rogue with the gift card api on your store!

You can download the XLSM file from this repository and it will work fine, but you shouldn't and I told you not to :grin:

[Screenshot-1]: https://github.com/stevenhoney/excel-shopify-gift-card-tool/blob/master/Screenshot-1.png
[Screenshot-2]: https://github.com/stevenhoney/excel-shopify-gift-card-tool/blob/master/Screenshot-2.png
[Screenshot-3]: https://github.com/stevenhoney/excel-shopify-gift-card-tool/blob/master/Screenshot-3.png
