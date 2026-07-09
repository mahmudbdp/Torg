ShipStation -> Bluecherry (MJSO) Converter
============================================

HOW TO USE
-----------
1. Double-click ShipStation_to_Bluecherry.html to open it in your browser
   (no install, no internet needed - everything runs locally in the browser).
2. Upload your ShipStation order export (.csv or .xlsx).
3. Check the Store -> Customer code mapping table (step 2). Add a row for
   any new marketplace/store before converting, if it's not listed yet.
   This table is saved in the browser, so you only add each store once.
4. Click "Convert to MJSO .xls", review the preview, then click Download.
5. Upload the downloaded .xls to Bluecherry.

sample_shipstation_export.csv is a sample input file kept here for reference/testing.

BUSINESS RULES BUILT INTO THE CONVERTER
----------------------------------------
- PO_NUM = "Custom - Field 2" from ShipStation; if that's blank, falls back
  to the Order Number.
- BILL_NUM = Order Number.
- CUST_NAME: "&" is replaced with "and" (e.g. "A & J Smith" -> "A and J Smith").
- SKU is split into STYLE / LBL_CODE / COLOR_CODE:
    - Text after the last "-" = COLOR_CODE.
    - The part before that: STYLE is everything up to (and including) the
      last digit; the letters after the last digit = LBL_CODE.
    - If there's no digit at all (e.g. "CAPWM"), the last 2 letters are
      taken as LBL_CODE and the rest as STYLE.
    - A trailing "-FBA" on a SKU is dropped before splitting (it's an
      Amazon FBA marker, not part of style/color).
    - If nothing matches, the row is flagged as a warning in the preview
      so you can check it manually before uploading.
- ORG_PRICE = (Amount Paid by Customer - Order Tax) / total quantity for
  that order (summed across all line items of the same order). This means
  discounted orders get an averaged net-of-tax unit price rather than the
  list price.
- START_DATE = the order date from ShipStation (date only, no time).
- END_DATE = today's date (the day you run the conversion).
- FRGT_AMT = "Amount - Order Shipping". For orders with multiple line
  items (multiple SKUs => multiple rows), only the FIRST row of that
  order gets the freight amount; the other rows get 0, so it isn't
  double-counted.
- TAX_AMT = "Amount - Order Tax", copied per row as-is.
- Constant values on every row: DIVISION=1, SEASON=ALL, CONF_TYPE=A,
  DIMENSION=-, SIZE_DESC=-, ORD_TYPE=REG, ADDR_TYPE=OT, SHIPPER=F11,
  MISC_AMT=0, TAX_EMPT=P, OVER_FRGT=P. STORE, UPC, PHONE, ADDRESS3,
  ADDRESS4, DEPARTMENT are left blank.

STORE -> CUSTOMER MAPPING (defaults, editable in the tool)
------------------------------------------------------------
  Zentail Amazon              -> TOAMA
  Zentail Outdoor Products    -> TOECO
  New eBay Store              -> TOEBA
  Zentail Okeechobee Fats     -> TOECO
  Zentail Samurai Tactical    -> TOECO

If you see a new store name in a future export that isn't in this list,
the tool will warn you after converting - add it via "+ Add row" in step 2.
