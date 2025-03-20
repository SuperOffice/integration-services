EXCEL QUOTE CONNECTOR READ ME
=============================

Testing connector
 _______                  _  
(_______)                | | 
 _____   _   _ ____ _____| | 
|  ___) ( \ / ) ___) ___ | | 
| |_____ ) X ( (___| ____| | 
|_______|_/ \_)____)_____)\_)
                             


SHEET 1 = Capabilities
------------------------
Column 1 = Capability Name
Column 2 = TRUE or FALSE

Capability Names
----------------------
"iproductprovider_provide_cost";
"iproductprovider_provide_minimumprice";
"iproductprovider_provide_stockdata";
"iproductprovider_provide_extradata";
"iproductprovider_provide_picture";
"iorderconsumer_provide_orderstate"
"ilistprovider_provide_productcategorylist"
"ilistprovider_provide_productfamilylist"
"ilistprovider_provide_producttypelist"
"ilistprovider_provide_paymenttermslist"
"ilistprovider_provide_paymenttypelist"
"ilistprovider_provide_deliverytermslist"
"ilistprovider_provide_deliverytypelist"
"iconnector_perform_complexsearch"
"iaddressprovider_provide_addresses"

"send_quote_url" - set the URL when sending quote. Browser should open.
"place_order_url" - set the URL when placing order. Browser should open.

"send_quote_soproto" - set the URL when sending quote. SuperOffice client should open/switch to contact card
"place_order_soproto" - set the URL when placing order. SuperOffice client should open/switch to project card

additional capabilites to provoke errors:
"cannot_start" - throw error during initialize
"fail_start" - return error during initialize
"warn_start" - return warning during initialize

"cannot_create" - throw error when creating quote
"fail_create" - return error when creating quote
"warn_create" - return warning when creating quote

"cannot_delete" - throw error when deleting quote

"cannot_send" - throw error when sending quote. Quote will still be sent.
"fail_send"   - return error when sending quote. Quote will still be sent.
"warn_send"   - error in log. Quote will still be sent.

"cannot_find" - throw error when searching. 

"cannot_product" - throw error when creating quote line from product.
"fail_product" - return error when creating quote line from product.
"warn_product" - return warning when creating quote line from product.

"cannot_save" - throw error when saving quote

"cannot_place_order" - throw error when placing order. Quote does not change state. Error dialog.
"fail_place_order" - return error message when placing order. Quote does not change state. Error dialog.

"cannot_validate_alt" - throw error when validating alternative. Error message
"fail_validate_alt" - return error  when validating alternative. Error message
"warn_validate_alt" - return warning when validating alternative. Warning message.

"cannot_validate_line" - throw error when validating alternative. Error message
"fail_validate_line" - return error  when validating alternative. Error message
"warn_validate_line" - return warning when validating alternative. Warning message.

"cannot_update" - throw error when updating prices. Error dialog.
"fail_update" - return error when updating prices. Error dialog.
"warn_update" - return warning when updating prices. Log message.

"cannot_order_state" - throw error when checking order state.
"fail_order_state" - return error when checking order state.
"warn_order_state" - return warning when checking order state.

"fail_configure_field_keys" - show warnings in admin / configure connection
"fail_configure_field_ranks" - show warnings in admin / configure connection

"remove_first_line_on_recalc" - remove the first quote line when the quote alternative is recalculated.  See bug 25210
"replace_90pct_lines_on_recalc" - replace any lines with more than 90% discount with a number of lines based on the quantity. See bug 25210

SHEET 2 = Addresses
------------------------
Column 1 = Contact id
Column 2 = Type of address (DELIVERY or INVOICE)
Column 3 = Address line 1
Column 4 = Address line 2
Column 5 = Address line 3
Column 6 = Address city
Column 7 = Address zip



SHEET 3 = PriceLists
---------------------------


How to setup the test connector:
--------------------------------
Add files to WebInstallation\bin folder:
SuperOffice.ExcelQuoteConnector.dll
SuperOffice.QuoteConnector.dll (Should already be installed)
SuperOffice.TestQuoteConnector.dll


Web.config
Add to DynamicLoad (TestQuoteConnector probably not needed)
	<!-- <add key="TestQuoteConnector" value="SuperOffice.TestQuoteConnector.dll"/> -->
    <add key="ExcelQuoteConnector" value="SuperOffice.ExcelQuoteConnector.dll"/>
		
	
Make a folder with the Excel document(s) and descriptions, for example C:\ExcelQuoteConnector. 
Must be shared with read/write. 

Make an ERP connection (in Sync tab) with the excel file as file: C:\ExcelQuoteConnector\ExcelConnectorTestxxx.xlsx	
Make an ERP connection in ERP connections tab. Use the above Erp connection and same excel path.

Pricelist must be in correct currency and have a good name
Pricelist are not visible in "SuperOffice products".

Ecxel products are visible only when searching in Quote. 
%% don't work, use actual product name. (Life or Women are good search terms)
