# qs-excel
Functions to help build option symbols and  other tools for RTD with Quotestream Professional.

# Overview
The file  quotestream.xlam is a VBA add-in for excel.   This provides functions including:
- build option symbols from expirty dates and root symbols.  Quotestream requires them in an odd format and it is 
a total nuisance to produce that format in an excel formula. 

- provide the RTD Server string so you don't have to retype it.



# Installation
The VBA Add in isin the file 'qs-excel.xlam".  

Read aout [how to install VBA Add Ins for excel](https://www.automateexcel.com/vba/install-add-in).  

# Usage

Currently there are two functions.  

In an excel cel:  '==qs_rtd_server()'  will return the constant string "quotestream.rtdserver".


The function 
```
Public Function QSOptionSymbol(root, country, expiry, optionType, strike)
    part1 = root + Space(6 - Len(root))
    pc = UCase(Trim(optionType))
    strike_str = Format(Str(strike * 1000), "00000000")
    expiryStr = Format(expiry, "yymmdd")
    Symbol = "@" & part1 & expiryStr & Trim(pc) & strike_str & ":" & country
    QSOptionSymbol = Symbol
End Function
```


will produce an option symbol required by Quotestream Professional RTD.  The expiry should be a date in a spreadhseet cel.

The option symbol should look something like this: `@GLD   211217C00183000:US`

I have found you have to run both Excel and Quotestream with elevated privileges (i.e. as Administrator) to use RTD successfuly.

Here is a sample to get the last trade for an option.

`=RTD("quotestream.rtdserver",,"@GLD   211217C00183000:US","Last")`

# Quotestream RTD Info

Please review the [Quotestream RTD Documentation ](http://www.quotestreampro.com/RTDdocs/).

