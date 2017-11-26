AVQ (Alpha Vantage [Query/Quote])
=================================

AVQ is an Excel-Add-In in form of a User-Defined Function (UDF) to get stock data from the Alpha Vantage API. 
Alpha Vantage provides a free API for realtime and historical data on stocks and other finance data in JSON or CVS formats. 
AVQ is a simple wrapper to get these data via an Excel function in Excel workbooks.
AVQ currently supports the following Alpha Vantage data:

* TIME_SERIES_DATA: Daily time series (date, daily open, daily high, daily low, daily close, daily volume)

AVQ required a free Alpha Vantage API Key, that can be requested on [www.alphavantage.co](https://www.alphavantage.co/support/#api-key).

For users of the shutdown [xlquote](www.xlquotes.com) Add-In, AVQ provides an UDF with the same function signature "XLQ" as the original one, but supports only selected parameter values.

AVQ is an independent development and has no relationship to Alpha Vantage.
In general, the same Alpha Vantage [term of services](https://www.alphavantage.co/terms_of_service/) apply.

AVQ uses [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) to parse and convert the JSON result returned by the Alpha Vantage API.

## Installation

1. Get free Alpha Vantage API Key [here](https://www.alphavantage.co/support/#api-key) (only Email required).
2. Add Excel-Add-In (file `AVQ.xlam`) to Excel via `File | (Excel) Options | Add-ins | Excel Add-Ins` or Ribbon `Developer | (Excel) Add-ins`.  
   Change your `Trust Center Settings`, if necessary.
3. Set your Alpha Vantage API Key via the `Set API Key` function in the new Ribbon tab `AVQ`.
4. See [Usage Examples](#usage-examples) for AVQ User-Defined Functions and their parameters.

Tested with:
* Excel 2007 (32-bit)
* Excel 2010 (32-bit)
* Excel 2013 (32-bit)
* Excel 2016 (32-bit)
* Excel 2016 (64-bit)

Excel 2003 is currently not supported.

## Usage Examples
Use AVQ in your personal Excel based financial reporting to update your current stock values.
![MyFinance](/MyFinance.png "Example of personal financial report")

Excel Formula | Result
------------- | -------
=AVQD("MSFT") | Returns the recent "close" stock quote of Microsoft Corporation.
=AVQD("MSFT";"close") | Returns the recent "close" stock quote of Microsoft Corporation.
=AVQD("MSFT";"high";-2) | Returns the "high" stock quote of Microsoft Corporation from two days ago.
=AVQD("MSFT";"open";5) | Returns the 5th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation.
=AVQD("MSFT";"volume";;"2017-11-15") | Returns the trading volume of Microsoft Corporation of 2017-11-15.
=XLQ("MSFT")| Returns the recent "close" stock quote of Microsoft Corporation.

## Documentation
```vbnet
''=======================================================
'' PROGRAM: AVQ (Alpha Vantage Query) Excel-Add-In (https://github.com/LelandGrunt/AVQ)
'' VERSION: 1.0.0
'' LICENSE: MIT (https://opensource.org/licenses/MIT)
'' DESCRIPTION: Simulation of the XLQ User-Defined Function of Excel-Add-In xlquotes (Stock Prices for Microsoft Excel©).
''              Also, Excel User-Defined Function wrapper for Alpha Vantage API for financial data.
''              Supported Alpha Vantage functions: Stock Time Series Data on Daily basis (TIME_SERIES_DAILY).
'' PREREQUIREMENT: Free Alpha Vantage API Key (https://www.alphavantage.co/support/#api-key)
'' ARGUMENTS: symbol - Ticker symbol (Mandatory)
''              item - Name of item to return (Optional, Default "close")
''               day - Day/X-th item of the time series (Optional, Default 0)
''              date - Date of the time series (Optional)
'' EXAMPLES: =AVQD("MSFT")                        - Returns the recent "close" stock quote of Microsoft Corporation.
''           =AVQD("MSFT";"close")                - Returns the recent "close" stock quote of Microsoft Corporation.
''           =AVQD("MSFT";"high";-2)              - Returns the "high" stock quote of Microsoft Corporation from two day ago.
''           =AVQD("MSFT";"open";7)               - Returns the 7th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation.
''           =AVQD("MSFT";"volume";;"2017-11-15") - Returns the trading volume of Microsoft Corporation of 2017-11-15.
''           =XLQ("MSFT")                         - Returns the recent "close" stock quote of Microsoft Corporation.
'' ERROR: Returns Excel error #N/A if no value was found, or #NULL! if an error occurred.
'' COMMENTS: (1) Supported item values: open, high, low, close, volume.
''           (2) If day = 0 then most recent data point is selected (Date unspecific).
''               If day < 0 then data point of current date minus day is selected (Date specific).
''               If day > 0 then the x-th (x = day) data point is selected (Date unspecific).
''           (3) If date is given then it overrules day parameter.
'' CHANGES----------------------------------------------
'' Date         Developer      Change
'' 2017-11-26   Leland Grunt   First public release.
'' -----------------------------------------------------
'' Copyright © 2017 Leland Grunt (leland.grunt[at]gmail.com)
'' Stock data are provided by Alpha Vantage (www.alphavantage.co)
'' Alpha Vantage is a copyright of Alpha Vantage Inc.
'' Alpha Vantage Terms of Service https://www.alphavantage.co/terms_of_service/
'' Uses VBA JSON Parser and Converter (https://github.com/VBA-tools/VBA-JSON)
'' xlquotes is a copyright of Dirk Voigtländer (www.xlquotes.com)
''=======================================================

Public Function AVQD(ByVal symbol As String, _
                     Optional ByVal item As String = "close", _
                     Optional ByVal day As Integer = 0, _
                     Optional ByVal quoteDate As Date) As Variant

Public Function XLQ(ticker As String, _
                    Optional side As String = "LAST", _
                    Optional hist As String = "") As Variant
```

## Files
* `AlphaVantageQuery.bas`: AVQ source code
* `AVQ - Examples.xlsx`: AVQ Examples and Tests
* `AVQ.xlam`: AVQ Excel-Add-In
* `LICENSE`: The MIT License text
* `MyFinance.png`: MyFinance Report Usage Example
* `README.md`: This readme text

## License
[MIT](https://opensource.org/licenses/MIT)