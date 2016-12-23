# myWebServer
Програма збільшує можливості FastReport; Команди передається через HTTP-запити

##Формат команд
"http://localhost:8080/" + GUID

*GUID задається у строці запуску*


##Команди
+ server
  + hello
  + stop
  + show
  + hide
+ getPricesFromExcel
  + fileload
    + FileName=
    + Supplier=*(REHAU (default), Accent Plast, Soldi)*
  + getprices
  + getrow
  + getmarkings
  + getprice
  + help
