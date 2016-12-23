# myWebServer
Програма збільшує можливості FastReport; Команди передаються через HTTP-запити

##Формат команд
"http://localhost:8080/" + GUID

*GUID задається у строці запуску*

Приклади:

http://localhost:8080/123/getPricesFromExcel/help?

http://localhost:8080/123/server/hello?

http://localhost:8080/123/getPricesFromExcel/fileload?filename=D:\price.xls&supplier=soldi


##Команди
+ server/
  + hello?
  + stop?
  + show?
  + hide?
+ getPricesFromExcel/
  + fileload?
    + FileName=
    + Supplier=
      + REHAU (default)
      + Accent Plast
      + Soldi
  + getprices?
    + outputDataType=
      + tabbed strings (default)
      + dictionarylist (JSONEncoded)
  + getrow?
    + searchFor= 
      + *a name of column*
    + value=
      + *a value fo searching*
    + searchOptions=
      + starts with
      + value starts with (default)
      + equal
      + equal of lowered
    + outputDataType=
      + tabbed strings (default)
      + dictionarylist (JSONEncoded)
  + getmarkings?
    + "searchFor= 
      + *a name of column*
    + value= 
      + *a value fo searching*
    + searchOptions=
      + starts with (default)
      + value starts with
      + equal
      + equal of lowered
  + getprice?
    + "searchFor= 
      + *a name of column*
    + value= 
      + *a value for searching*
    + searchOptions=
      + value starts with (default)
      + equal
      + equal of lowered
  + help?
