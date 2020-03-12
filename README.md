# SierraLibrary



[![JHAvLX.jpg](https://iili.io/JHAvLX.jpg)](https://dotnet.microsoft.com/download/dotnet-framework/net452)   [![JHASBn.png](https://iili.io/JHASBn.png)](https://products.office.com/en/excel)   [![JHAU1s.png](https://iili.io/JHAU1s.png)](https://www.iii.com/products/sierra-ils/)  




# About
**SierraFunctions** class implements functions to make [Sierra Apis](https://techdocs.iii.com/sierraapi/Content/titlePage.htm#)
call using [Excel-DNA](https://excel-dna.net/)


# Build
Clone repo in **VS2019** and change **base_url** and **header_basic_auth** in SierraFunctions class according to 
[SIERRA API Documentation](https://techdocs.iii.com/sierraapi/Content/zTutorials/tutAuthenticate.htm) then build.


# Installation
Open **Excel** then click on **Developer** tab and click **Excel Add-ins** button.
Click **Browse...** button. Select **SierraLibrary-AddIn.xll** file in **SierraLibrary\bin\Release** directory.

[![JHGRRe.png](https://iili.io/JHGRRe.png)](https://freeimage.host/i/excel1.JHGRRe)

After install **Excel Add-ins** you can see all implemented functions in **sierra** category.

[![JHGliQ.png](https://iili.io/JHGliQ.png)](https://freeimage.host/i/excel2.JHGliQ)

# Functions
  - [About()](#About)
  - [GetToken()](#GetToken)
  - [Barcode2Id()](#Barcode2Id)
  - [Barcode2Name()](#Barcode2Name)
  - [Barcode2Email()](#Barcode2Email)
  - [Barcode2PatronType()](#Barcode2PatronType)
  - [Barcode2MoneyOwed()](#Barcode2MoneyOwed)
  - [Barcode2CheckoutItems()](#Barcode2CheckoutItems)
  - [Item2BibId()](#Item2BibId)
  - [BibId2Title()](./README.md#BibId2Title)


##### About()
Returns author information.

Sample usage: =About()

[![JHGE0B.png](https://iili.io/JHGE0B.png)](https://freeimage.host/i/about1.JHGE0B)

##### GetToken()
Returns token.

Sample usage: =GetToken()

Sample output in Postman
```
		 {
		   "access_token": "v0Qvd3EscNjMPF9zH606RebLuOaVrTuG6Bs9Vf1_cPFxRKCJPWSbTPOlTOi-bLF17Hcl-8-A2UdTvyMhZfIDATYKLgnh5y_02xNqYq9PGIQ",
		   "token_type": "bearer",
		   "expires_in": 3600
		 }
```



##### Barcode2Id()
Returns patron id by barcode.

Sample usage: = Barcode2Id("1845")

[![JHGwzJ.png](https://iili.io/JHGwzJ.png)](https://freeimage.host/i/excel3.JHGwzJ)

[![JHGNWv.png](https://iili.io/JHGNWv.png)](https://freeimage.host/i/excel4.JHGNWv)

[![JHGkfp.png](https://iili.io/JHGkfp.png)](https://freeimage.host/i/excel6.JHGkfp)




##### Barcode2Name()
Returns patron name by barcode.

Sample usage: =Barcode2Name("1845")

[![JHG8gI.png](https://iili.io/JHG8gI.png)](https://freeimage.host/i/barcode2name.JHG8gI)

[![JHGUJt.png](https://iili.io/JHGUJt.png)](https://freeimage.host/i/barcode2name1.JHGUJt)

[![JHGg5X.png](https://iili.io/JHGg5X.png)](https://freeimage.host/i/barcode2name2.JHGg5X)



##### Barcode2Email()
Returns patron e-mail by barcode.

Sample usage: =Barcode2Email("1845")

[![JHGren.png](https://iili.io/JHGren.png)](https://freeimage.host/i/barcode2email.JHGren)

[![JHG4bs.png](https://iili.io/JHG4bs.png)](https://freeimage.host/i/barcode2email1.JHG4bs)

[![JHGiXf.png](https://iili.io/JHGiXf.png)](https://freeimage.host/i/barcode2email2.JHGiXf)



##### Barcode2PatronType()
Returns patron type by barcode.

Sample usage: =Barcode2PatronType("1845")

[![JHGQql.png](https://iili.io/JHGQql.png)](https://freeimage.host/i/barcode2patrontype.JHGQql)

[![JHGZ12.png](https://iili.io/JHGZ12.png)](https://freeimage.host/i/barcode2patrontype1.JHGZ12)

[![JHGtgS.png](https://iili.io/JHGtgS.png)](https://freeimage.host/i/barcode2patrontype2.JHGtgS)


##### Barcode2MoneyOwed()
Returns patron money owed by barcode.

Sample usage: =Barcode2MoneyOwed("1845")

[![JHGm79.png](https://iili.io/JHGm79.png)](https://freeimage.host/i/money.JHGm79)

[![JHGpee.png](https://iili.io/JHGpee.png)](https://freeimage.host/i/money1.JHGpee)

[![JHGymu.png](https://iili.io/JHGymu.png)](https://freeimage.host/i/money2.JHGymu)

##### Barcode2CheckoutItems()
Returns checkout items by barcode.

Sample usage: =Barcode2CheckoutItems("1845")

[![JHMdLx.png](https://iili.io/JHMdLx.png)](https://freeimage.host/i/checkout.JHMdLx)

[![JHMF1V.png](https://iili.io/JHMF1V.png)](https://freeimage.host/i/checkout1.JHMF1V)

[![JHMnmg.png](https://iili.io/JHMnmg.png)](https://freeimage.host/i/checkout2.JHMnmg)



##### Item2BibId()
Returns bib id by item.

Sample usage: =Item2BibId("1136526")

[![JHMzhJ.png](https://iili.io/JHMzhJ.png)](https://freeimage.host/i/item2bib.JHMzhJ)

[![JHMILv.png](https://iili.io/JHMILv.png)](https://freeimage.host/i/item2bib1.JHMILv)

[![JHMuBR.png](https://iili.io/JHMuBR.png)](https://freeimage.host/i/item2bib2.JHMuBR)



##### BibId2Title()
Returns title by bib id.

Sample usage: =BibId2Title("1159654")

[![JHMA1p.png](https://iili.io/JHMA1p.png)](https://freeimage.host/i/bib2title.JHMA1p)

[![JHM72I.png](https://iili.io/JHM72I.png)](https://freeimage.host/i/bib2title1.JHM72I)

[![JHMY7t.png](https://iili.io/JHMY7t.png)](https://freeimage.host/i/bib2title2.JHMY7t)
----






License
----

Copyright (c) [Gyokay Nurvet Mustafa](https://gyokay.cloud/). All rights reserved.

Licensed under the [MIT](https://github.com/gyokaynurvet/Sierra-Library/blob/master/LICENSE) License.

**Free Software**

Made with ❤ in Turkey

[//]: # (References)
[//]: # (https://dillinger.io/ Online Markdown editor)
[//]: # (https://freeimage.host/ Free image hosting gokay.gursoy@gmail.com Google Login)
[//]: # (https://techdocs.iii.com/sierraapi/Content/titlePage.htm#)
[//]: # (https://excel-dna.net/)
[//]: # (https://techdocs.iii.com/sierraapi/Content/zTutorials/tutAuthenticate.htm)


