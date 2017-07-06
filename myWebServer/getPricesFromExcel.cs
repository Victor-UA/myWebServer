using System;
using System.Collections.Specialized;
using System.Linq;
using System.IO;
using System.Data;
using Excel;
using System.Collections.Generic;

namespace myWebServer
{
    class GetPricesFromExcel
    {
        public GetPricesFromExcel(string guid)
        {
            GUID = guid;
        }

        static public Dictionary<string, GetPricesFromExcel> Sessions { get; set; } =
            new Dictionary<string, GetPricesFromExcel>();

        public string GUID { get; private set; }

        private DataSet excelDataSet;
        private DataSet pricesDataSet;
        public string Execute(string Command, NameValueCollection QueryString)
        {            
            switch (Command.ToLower())
            {
                case "fileload":
                        return FileLoad(QueryString);
                case "getprices":
                    return getPrices(QueryString);
                case "getrow":
                    return getRow(QueryString);
                case "getmarkings":
                    return getMarkings(QueryString);
                case "getprice":
                    return getPrice(QueryString);
                case "help":
                    return "Info!\n" +
                        "FileLoad\n" +
                        "getPrices\n" +
                        "getRow\n" +
                        "getMarkings\n" +
                        "getPrice";
                    ;
                default:
                    return Command + ": Unknown command";
            }
            
        }

        public void ExcelClose()
        {
            Sessions.Remove(GUID);
            /*
            try
            {
                foreach (Excel.Window window in excelApp.Windows)
                {
                    window.Close(false, Type.Missing, Type.Missing);
                }
                excelApp.Quit();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e.Message);
            }
            */
            
        }

        private string FileLoad(NameValueCollection QueryString)
        {
            if (QueryString.AllKeys.Contains("help", StringComparer.CurrentCultureIgnoreCase))
            {
                return "Info!\n" +
                    "FileName=\n" +
                    "Supplier=\n" +
                    "\tREHAU (default)\n" +
                    "\tAccent Plast\n" + 
                    "\tSoldi"
                ;
            }
            excelDataSet = null;
            if (QueryString.AllKeys.Contains("filename", StringComparer.CurrentCultureIgnoreCase))
            {
                string fileName = QueryString.GetValues("filename").GetValue(0).ToString();
                if (!File.Exists(fileName))
                {
                    return ("Error!\n" + fileName + ": File not found");
                }
                try
                {
                    FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    excelDataSet = excelReader.AsDataSet();
                    excelReader.Close();

                    string preparingResult = Prepare(QueryString);

                    if (preparingResult.StartsWith("Ok!"))
                    {
                        return "Ok!\nFile has been loaded successfully\n" +
                        "Worksheets=" + excelDataSet.Tables.Count + '\n' +
                        preparingResult
                        ;
                    }
                    else
                    {
                        return "Error!\nFile has been loaded successfully\n" +
                        "Worksheets=" + excelDataSet.Tables.Count + '\n' +
                        "But\n" + 
                        preparingResult
                        ;
                    }

                    
                }
                catch (Exception e)
                {
                    return ("Error!\n" + e.Message);
                }
                
            }
            else
            {
                return "Error!\nFileName not found";
            }
        }
        private string Prepare(NameValueCollection QueryString)
        {
            int Count = 0;
            string supplier = "";
            if (QueryString.AllKeys.Contains("supplier", StringComparer.CurrentCultureIgnoreCase))
            {
                supplier = QueryString.GetValues("supplier").GetValue(0).ToString();
            }
            
            if (excelDataSet != null)
            {
                pricesDataSet = new DataSet();
                for (int i = 0; i < excelDataSet.Tables.Count; i++)
                {
                    Count += GetPricesFromSheet(excelDataSet.Tables[i], i, supplier);
                }
            }
            else
            {
                return "Error!\nexcelDataSet is empty.";
            }
            return "Ok!\n" + 
                "Found ("+ Count + ") rows"
                ;
        }

        private string getRow(NameValueCollection QueryString)
        {
            if (QueryString.AllKeys.Contains("help", StringComparer.CurrentCultureIgnoreCase))
            {
                return "Info!\n" +
                    "searchFor= a name of column\n" +
                    "value= a value fo searching\n" +
                    "searchOptions=\n" +
                    "\t[starts with]\n" +
                    "\t[value starts with] (default)\n" +
                    "\t[equal]\n" +
                    "\t[equal of lowered]\n" + 
                    "outputDataType=\n" +
                    "\t[tabbed strings] (default)\n" +
                    "\t[dictionarylist]: JSONEncode(DictionaryList)"
                ;
            }

            string outputDataType = "tabbed strings";
            if (QueryString.AllKeys.Contains("outputdatatype", StringComparer.CurrentCultureIgnoreCase))
            {
                outputDataType = QueryString.GetValues("outputDataType").GetValue(0).ToString().ToLower();
            }

            string columnName = string.Empty;
            string value = string.Empty;
            string searchOptions = string.Empty;

            if (QueryString.AllKeys.Count() > 0)
            {
                columnName = QueryString.AllKeys.Contains("searchfor", StringComparer.CurrentCultureIgnoreCase) ? QueryString.GetValues("searchFor").GetValue(0).ToString() : string.Empty;
                value = QueryString.AllKeys.Contains("value", StringComparer.CurrentCultureIgnoreCase) ? QueryString.GetValues("value").GetValue(0).ToString() : string.Empty;
                searchOptions = QueryString.AllKeys.Contains("searchoptions", StringComparer.CurrentCultureIgnoreCase) ? QueryString.GetValues("searchOptions").GetValue(0).ToString() : string.Empty;
            }

            string Result = string.Empty;
            if (excelDataSet != null)
            {
                DataTable resultTable = new DataTable();
                switch (searchOptions)
                {
                    case "equal":
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                Found = Row[Table.Columns[columnName]].ToString().Equals(value);
                                if (Found)
                                {
                                    Result = Row2outputDataType(Row, Table, outputDataType);
                                    break;
                                }
                            }
                            if (Found)
                            {
                                break;
                            }
                        }
                        break;
                    case "equal of lowered":
                        value = value.ToLower();
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                Found = Row[Table.Columns[columnName]].ToString().ToLower().Equals(value);
                                if (Found)
                                {
                                    Result = Row2outputDataType(Row, Table, outputDataType);
                                    break;
                                }
                            }
                            if (Found)
                            {
                                break;
                            }
                        }
                        break;
                    case "starts with":
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                string source = Row[Table.Columns[columnName]].ToString();
                                Found = !source.Equals(string.Empty) && source.StartsWith(value);
                                if (Found)
                                {
                                    Result = Row2outputDataType(Row, Table, outputDataType);
                                    break;
                                }
                            }
                            if (Found)
                            {
                                break;
                            }
                        }
                        break;
                    default:
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                string source = Row[Table.Columns[columnName]].ToString();
                                Found = !source.Equals(string.Empty) && value.StartsWith(source);
                                if (Found)
                                {
                                    Result = Row2outputDataType(Row, Table, outputDataType);
                                    break;
                                }
                            }
                            if (Found)
                            {
                                break;
                            }
                        }
                        break;
                }
                if (outputDataType == "dictionarylist")
                {
                    if (!Result.Equals(string.Empty))
                    {
                        Result = '[' + Result + ']';
                    }
                }
            }
            else
            {
                return "Error!\nexcelDataSet is empty.";
            }

            return (Result.Equals(string.Empty) ? "Error!Data is not found\n" : "Ok!\n") + Result;
        }

        private string getPrices(NameValueCollection QueryString)
        {
            if (QueryString.AllKeys.Contains("help", StringComparer.CurrentCultureIgnoreCase))
            {
                return "Info!\n" +
                    "outputDataType=\n" +
                    "\t[tabbed strings] (default)\n" +
                    "\t[dictionarylist]: JSONEncode(DictionaryList)";
            }

            string outputDataType = "tabbed strings";
            if (QueryString.AllKeys.Contains("outputdatatype", StringComparer.CurrentCultureIgnoreCase))
            {
                outputDataType = QueryString.GetValues("outputDataType").GetValue(0).ToString().ToLower();
            }

            string Result = string.Empty;
            if (excelDataSet != null) {
                switch (outputDataType)
                {
                    case "dictionarylist":
                        for (int i = 0; i < excelDataSet.Tables.Count; i++)
                        {
                            Result += (Result.Length == 0 ? "" : ",") + GetPricesFormDataSetSheet(pricesDataSet.Tables[i], outputDataType);
                        }
                        Result = '[' + Result + ']';
                        break;
                    default:
                        for (int i = 0; i < excelDataSet.Tables.Count; i++)
                        {
                            Result += GetPricesFormDataSetSheet(pricesDataSet.Tables[i], outputDataType);
                        }
                        break;
                }
            }
            else
            {
                return "Error!\nexcelDataSet is empty.";
            }
            
            return "Ok!\n" + Result;
        }
        private string getPrice(NameValueCollection QueryString)
        {
            if (QueryString.AllKeys.Contains("help", StringComparer.CurrentCultureIgnoreCase))
            {
                return "Info!\n" +
                    "searchFor= a name of column\n" +
                    "value= a value fo searching\n" +
                    "searchOptions=\n" +
                    "\t[value starts with] as default\n" +
                    "\t[equal]\n" +
                    "\t[equal of lowered]"
                    ;
            }

            string columnName = string.Empty;
            string value = string.Empty;
            string searchOptions = string.Empty;

            if (QueryString.AllKeys.Count() > 0)
            {
                columnName = QueryString.AllKeys.Contains("searchfor", StringComparer.CurrentCultureIgnoreCase) ? QueryString.GetValues("searchFor").GetValue(0).ToString()  : string.Empty;
                value = QueryString.AllKeys.Contains("value", StringComparer.CurrentCultureIgnoreCase) ? QueryString.GetValues("value").GetValue(0).ToString() : string.Empty;
                searchOptions = QueryString.AllKeys.Contains("searchoptions", StringComparer.CurrentCultureIgnoreCase) ? QueryString.GetValues("searchOptions").GetValue(0).ToString() : string.Empty;
            }
            
            string Result = string.Empty;
            if (excelDataSet != null)
            {
                switch (searchOptions)
                {
                    case "equal":
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                Found = Row[Table.Columns[columnName]].ToString().Equals(value);
                                if (Found)
                                {
                                    Result = Row[Table.Columns["price1"]].ToString();
                                    break;
                                }
                            }
                            if (Found)
                            {
                                break;
                            }
                        }
                        break;
                    case "equal of lowered":
                        value = value.ToLower();
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                Found = Row[Table.Columns[columnName]].ToString().ToLower().Equals(value);
                                if (Found)
                                {
                                    Result = Row[Table.Columns["price1"]].ToString();
                                    break;
                                }
                            }
                            if (Found)
                            {
                                break;
                            }
                        }
                        break;
                    default:
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                string source = Row[Table.Columns[columnName]].ToString();
                                Found = !source.Equals(string.Empty) && value.StartsWith(source);
                                if (Found)
                                {
                                    Result = Row[Table.Columns["price1"]].ToString();
                                    break;
                                }
                            }
                            if (Found)
                            {
                                break;
                            }
                        }
                        break;
                }
            }
            else
            {
                return "Error!\nexcelDataSet is empty.";
            }

            return (Result == string.Empty ? "Error!Data is not found\n" : "Ok!\n") + Result;
        }
        private string getMarkings(NameValueCollection QueryString)
        {
            if (QueryString.AllKeys.Contains("help", StringComparer.CurrentCultureIgnoreCase))
            {
                return "Info!\n" +
                    "searchFor= a name of column\n" +
                    "value= a value fo searching\n" +
                    "searchOptions=\n" +
                    "\t[starts with] as default\n" +
                    "\t[value starts with]\n" +
                    "\t[equal]\n" +
                    "\t[equal of lowered]"
                    ;
            }

            string columnName = string.Empty;
            string value = string.Empty;
            string searchOptions = string.Empty;

            if (QueryString.AllKeys.Count() > 0)
            {
                columnName = QueryString.AllKeys.Contains("searchfor", StringComparer.CurrentCultureIgnoreCase) ? QueryString.GetValues("searchFor").GetValue(0).ToString() : string.Empty;
                value = QueryString.AllKeys.Contains("value", StringComparer.CurrentCultureIgnoreCase) ? QueryString.GetValues("value").GetValue(0).ToString() : string.Empty;
                searchOptions = QueryString.AllKeys.Contains("searchoptions", StringComparer.CurrentCultureIgnoreCase) ? QueryString.GetValues("searchOptions").GetValue(0).ToString() : string.Empty;
            }

            string Result = string.Empty;
            if (excelDataSet != null)
            {
                switch (searchOptions)
                {
                    case "equal":
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                Found = Row[Table.Columns[columnName]].ToString().Equals(value);
                                if (Found)
                                {
                                    Result += (Result.Length > 0 ? "\t" : "") + Row[Table.Columns[columnName]].ToString();
                                }
                            }
                        }
                        break;
                    case "equal of lowered":
                        value = value.ToLower();
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                Found = Row[Table.Columns[columnName]].ToString().ToLower().Equals(value);
                                if (Found)
                                {
                                    Result += (Result.Length > 0 ? "\t" : "") + Row[Table.Columns[columnName]].ToString();
                                }
                            }
                        }
                        break;
                    case "value starts with":
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                string source = Row[Table.Columns[columnName]].ToString();
                                Found = !source.Equals(string.Empty) && value.StartsWith(source);
                                if (Found)
                                {
                                    Result += (Result.Length > 0 ? "\t" : "") + Row[Table.Columns[columnName]].ToString();
                                }
                            }
                        }
                        break;
                    default:
                        foreach (DataTable Table in pricesDataSet.Tables)
                        {
                            bool Found = false;
                            foreach (DataRow Row in Table.Rows)
                            {
                                Found = !value.Equals(string.Empty) && Row[Table.Columns[columnName]].ToString().StartsWith(value);
                                if (Found)
                                {
                                    Result += (Result.Length > 0 ? "\t" : "") + Row[Table.Columns[columnName]].ToString();
                                }
                            }
                        }
                        break;
                }
            }
            else
            {
                return "Error!\nexcelDataSet is empty.";
            }

            return (Result == string.Empty ? "Error!Data is not found\n" : "Ok!\n") + Result;
        }


        private int RowIndexOf(DataRow row, object obj)
        {
            for (int i = 0; i < row.ItemArray.Length; i++)
            {
                if (obj.Equals(row.ItemArray[i]))
                {
                    return i;
                }
            }
            return -1;
        }

        private int RowIndexOfStartsWith(DataRow row, string search)
        {
            for (int i = 0; i < row.ItemArray.Length; i++)
            {
                if (row.ItemArray[i].ToString().StartsWith(search))
                {
                    return i;
                }
            }
            return -1;
        }

        private int GetPricesFromSheet(DataTable Prices, int sheetNum)
        {
            return GetPricesFromSheet(Prices, sheetNum, "REHAU");
        }
        private int GetPricesFromSheet(DataTable Prices, int sheetNum, string Supplier)
        {
            int Marking2Column = -1;
            int MarkingColumn = -1;
            int TitleColumn = -1;
            int PriceColumn = -1;
            int PriceUSDColumn = -1;
            int PriceUAHColumn = -1;
            int FirstPriceRow = -1;
            int LastPriceRow = -1;
            bool EndOfPrice = false;
            bool isOk = false;

            int Result = 0;

            DataTable Table = new DataTable(Prices.TableName);
            Table.Columns.Add("sheetName", typeof(string));
            Table.Columns.Add("sheetNum", typeof(int));
            Table.Columns.Add("rowNum", typeof(int));
            Table.Columns.Add("marking", typeof(string));
            Table.Columns.Add("marking2", typeof(string));
            Table.Columns.Add("ggname", typeof(string));
            Table.Columns.Add("currency", typeof(string));
            Table.Columns.Add("price1", typeof(double));
            Table.Columns.Add("price1USD", typeof(double));
            Table.Columns.Add("price1UAH", typeof(double));
            pricesDataSet.Tables.Add(Table);

            string currency = "mainCurrency";

            switch (Supplier.ToLower())
            {
                case "accent plast":
                    #region Accent Plast
                    Marking2Column = 0;
                    MarkingColumn = 1;
                    TitleColumn = 2;
                    PriceColumn = 3;
                    PriceUSDColumn = 4;
                    PriceUAHColumn = 5;
                    FirstPriceRow = -1;
                    LastPriceRow = -1;
                    EndOfPrice = false;
                    isOk = false;

                    Result = 0;

                    for (int i = 0; i < Prices.Rows.Count && !isOk; i++)
                    {
                        DataRow row = Prices.Rows[i];
                        #region Look4Columns
                        if (Marking2Column < 0)
                        {
                            Marking2Column = RowIndexOf(row, "Позиция");
                        }
                        if (TitleColumn < 0)
                        {
                            TitleColumn = RowIndexOfStartsWith(row, "Описание позиции");
                        }
                        else
                        {
                            if (FirstPriceRow < 0)
                            {
                                double tmp;
                                if (double.TryParse(row.ItemArray[Marking2Column].ToString(), out tmp) || (!string.IsNullOrEmpty(row.ItemArray[MarkingColumn].ToString())))
                                {
                                    FirstPriceRow = i;
                                    LastPriceRow = i;
                                }
                            }
                            else
                            {
                                double tmp;
                                if (!EndOfPrice && (double.TryParse(row.ItemArray[Marking2Column].ToString(), out tmp) || !string.IsNullOrEmpty(row.ItemArray[MarkingColumn].ToString())))
                                {
                                    LastPriceRow = i;
                                }
                                else
                                {
                                    //EndOfPrice = true;
                                }
                            }
                        }
                        if (PriceColumn < 0)
                        {
                            PriceColumn = RowIndexOfStartsWith(row, "евро с пдв");
                        }
                        #endregion

                        isOk = (EndOfPrice || i + 1 == Prices.Rows.Count) &&
                            Marking2Column >= 0 && PriceColumn >= 0 && FirstPriceRow >= 0 && TitleColumn >= 0 &&
                            PriceUSDColumn >= 0 && PriceUAHColumn >= 0
                            ;

                    }
                    if (isOk)
                    {
                        for (int i = FirstPriceRow; i <= LastPriceRow; i++)
                        {
                            DataRow row = Prices.Rows[i];
                            string marking = row.ItemArray[MarkingColumn].ToString();
                            string marking2 = row.ItemArray[Marking2Column].ToString();
                            string ggname = row.ItemArray[TitleColumn].ToString();

                            
                            if (marking.Equals(string.Empty))
                            {
                                #region Marking & ggName from TitleColumn
                                string title = row.ItemArray[TitleColumn].ToString();

                                string[] titleArr = title.Split('_');
                                if (titleArr.Length > 2)
                                {
                                    marking = titleArr[2];
                                    if (marking.Contains(' ') || titleArr.Length == 3)
                                    {
                                        string[] markingArr = marking.Split(' ');
                                        try
                                        {
                                            marking = markingArr[0];
                                        }
                                        catch { }
                                        for (int j = 1; j < markingArr.Length; j++)
                                        {
                                            try
                                            {
                                                ggname += ggname.Equals(string.Empty) ? markingArr[j] : " " + markingArr[j];
                                            }
                                            catch { }
                                        }

                                    }

                                    for (int j = 3; j < titleArr.Length; j++)
                                    {
                                        try
                                        {
                                            ggname += ggname.Equals(string.Empty) ? titleArr[j] : "_" + titleArr[j];
                                        }
                                        catch { }
                                    }
                                }
                                else
                                {
                                    marking = "";
                                }
                                #endregion
                            }

                            double price1 = 0;
                            double price1USD = 0;
                            double price1UAH = 0;

                            try
                            {
                                //price1 = Convert.ToDouble(row.ItemArray[PriceColumn].ToString().Replace(",", "."));
                                double.TryParse(row.ItemArray[PriceColumn].ToString().Replace(",", "."), out price1);
                            }
                            catch (Exception) { }
                            try
                            {
                                //price1USD = Convert.ToDouble(row.ItemArray[PriceUSDColumn].ToString().Replace(",", "."));
                                double.TryParse(row.ItemArray[PriceUSDColumn].ToString().Replace(",", "."), out price1USD);
                            }
                            catch (Exception) { }
                            try
                            {
                                //price1UAH = Convert.ToDouble(row.ItemArray[PriceUAHColumn].ToString().Replace(",", "."));
                                double.TryParse(row.ItemArray[PriceUAHColumn].ToString().Replace(",", "."), out price1UAH);
                            }
                            catch (Exception) { }
                            currency = price1 > 0 ? "mainCurrency" : price1USD > 0 ? "USD" : price1UAH > 0 ? "UAH" : "mainCurrency";
                            if (price1 != 0 || price1USD != 0 || price1UAH != 0)
                            {
                                Table.Rows.Add(new object[] { Prices.TableName, sheetNum, i, marking, marking2, ggname, currency, price1, price1USD, price1UAH });
                                Result += 1;
                            }
                        }
                    }
                    return Result;
                #endregion
                case "soldi":
                    #region Soldi
                    MarkingColumn = 0;
                    TitleColumn = 1;
                    PriceColumn = 2;
                    PriceUSDColumn = 3;
                    PriceUAHColumn = 4;
                    FirstPriceRow = 1;
                    LastPriceRow = -1;
                    EndOfPrice = false;
                    isOk = false;

                    Result = 0;

                    for (int i = 0; i < Prices.Rows.Count && !isOk; i++)
                    {
                        DataRow row = Prices.Rows[i];
                        #region Look4Columns
                        /*
                        if (Marking2Column < 0)
                        {
                            Marking2Column = RowIndexOf(row, "Позиция");
                        }
                        */
                        if (TitleColumn < 0)
                        {
                            TitleColumn = RowIndexOfStartsWith(row, "Описание позиции");
                        }
                        else
                        {
                            /*
                            if (FirstPriceRow < 0)
                            {
                                double tmp;
                                if (double.TryParse(row.ItemArray[Marking2Column].ToString(), out tmp))
                                {
                                    FirstPriceRow = i;
                                    LastPriceRow = i;
                                }
                            }
                            
                            else*/
                            {
                                //double tmp;
                                if (!EndOfPrice /*&& double.TryParse(row.ItemArray[Marking2Column].ToString(), out tmp)*/)
                                {
                                    LastPriceRow = i;
                                }
                                else
                                {
                                    //EndOfPrice = true;
                                }
                            }
                        }
                        if (PriceColumn < 0)
                        {
                            PriceColumn = RowIndexOfStartsWith(row, "евро с пдв");
                        }
                        #endregion

                        isOk = (EndOfPrice || i + 1 == Prices.Rows.Count) &&
                            //Marking2Column >= 0 && PriceColumn >= 0 && FirstPriceRow >= 0 && TitleColumn >= 0 &&
                            PriceUSDColumn >= 0 && PriceUAHColumn >= 0
                            ;

                    }
                    if (isOk)
                    {
                        for (int i = FirstPriceRow; i <= LastPriceRow; i++)
                        {
                            DataRow row = Prices.Rows[i];
                            string marking = row.ItemArray[MarkingColumn].ToString();
                            //string marking2 = row.ItemArray[Marking2Column].ToString();
                            string ggname = row.ItemArray[TitleColumn].ToString();


                            if (marking.Equals(string.Empty))
                            {
                                #region Marking & ggName from TitleColumn
                                string title = row.ItemArray[TitleColumn].ToString();

                                string[] titleArr = title.Split('_');
                                if (titleArr.Length > 2)
                                {
                                    marking = titleArr[2];
                                    if (marking.Contains(' ') || titleArr.Length == 3)
                                    {
                                        string[] markingArr = marking.Split(' ');
                                        try
                                        {
                                            marking = markingArr[0];
                                        }
                                        catch { }
                                        for (int j = 1; j < markingArr.Length; j++)
                                        {
                                            try
                                            {
                                                ggname += ggname.Equals(string.Empty) ? markingArr[j] : " " + markingArr[j];
                                            }
                                            catch { }
                                        }

                                    }

                                    for (int j = 3; j < titleArr.Length; j++)
                                    {
                                        try
                                        {
                                            ggname += ggname.Equals(string.Empty) ? titleArr[j] : "_" + titleArr[j];
                                        }
                                        catch { }
                                    }
                                }
                                else
                                {
                                    marking = "";
                                }
                                #endregion
                            }

                            double price1 = 0;
                            double price1USD = 0;
                            double price1UAH = 0;

                            try
                            {
                                //price1 = Convert.ToDouble(row.ItemArray[PriceColumn].ToString().Replace(",", "."));
                                double.TryParse(row.ItemArray[PriceColumn].ToString().Replace(",", "."), out price1);
                            }
                            catch (Exception) { }
                            try
                            {
                                //price1USD = Convert.ToDouble(row.ItemArray[PriceUSDColumn].ToString().Replace(",", "."));
                                double.TryParse(row.ItemArray[PriceUSDColumn].ToString().Replace(",", "."), out price1USD);
                            }
                            catch (Exception) { }
                            try
                            {
                                //price1UAH = Convert.ToDouble(row.ItemArray[PriceUAHColumn].ToString().Replace(",", "."));
                                double.TryParse(row.ItemArray[PriceUAHColumn].ToString().Replace(",", "."), out price1UAH);
                            }
                            catch (Exception) { }
                            currency = price1 > 0 ? "mainCurrency" : price1USD > 0 ? "USD" : price1UAH > 0 ? "UAH" : "mainCurrency";
                            if (price1 != 0 || price1USD != 0 || price1UAH != 0)
                            {
                                Table.Rows.Add(new object[] { Prices.TableName, sheetNum, i, marking, "", ggname, currency, price1, price1USD, price1UAH });
                                Result += 1;
                            }
                        }
                    }
                    return Result;
                #endregion
                default:
                    #region REHAU
                    MarkingColumn = 1;//-1;
                    TitleColumn = 6;
                    PriceColumn = 7;
                    FirstPriceRow = -1;
                    LastPriceRow = -1;
                    EndOfPrice = false;
                    isOk = false;

                    Result = 0;

                    for (int i = 0; i < Prices.Rows.Count && !isOk; i++)
                    {
                        DataRow row = Prices.Rows[i];
                        if (MarkingColumn < 0)
                        {
                            MarkingColumn = RowIndexOf(row, "Артикул");
                        }
                        else
                        {
                            if (FirstPriceRow < 0)
                            {
                                double tmp;
                                if (double.TryParse(row.ItemArray[MarkingColumn].ToString(), out tmp))
                                {
                                    FirstPriceRow = i;
                                    LastPriceRow = i;
                                }
                            }
                            else
                            {
                                double tmp;
                                if (!EndOfPrice && double.TryParse(row.ItemArray[MarkingColumn].ToString(), out tmp))
                                {
                                    LastPriceRow = i;
                                }
                                else
                                {
                                    //EndOfPrice = true;
                                }
                            }
                        }
                        if (PriceColumn < 0)
                        {
                            PriceColumn = RowIndexOfStartsWith(row, "Ціна зі складу");
                        }
                        if (TitleColumn < 0)
                        {
                            TitleColumn = RowIndexOfStartsWith(row, "Найменування");
                        }

                        isOk = (EndOfPrice || i + 1 == Prices.Rows.Count) &&
                            MarkingColumn >= 0 && PriceColumn >= 0 && FirstPriceRow >= 0 && TitleColumn >= 0;

                    }
                    if (isOk)
                    {
                        for (int i = FirstPriceRow; i <= LastPriceRow; i++)
                        {
                            DataRow row = Prices.Rows[i];
                            string marking = row.ItemArray[MarkingColumn].ToString() + '.' + row.ItemArray[MarkingColumn + 1].ToString();
                            string ggname = row.ItemArray[TitleColumn].ToString();
                            double price1 = 0;

                            try
                            {
                                price1 = Convert.ToDouble(row.ItemArray[PriceColumn].ToString().Replace(",", "."));
                                price1 = price1 * 1.2;

                                Table.Rows.Add(new object[] { Prices.TableName, sheetNum, i, marking, "", ggname, currency, price1, 0, 0 });
                                Result += 1;
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                    return Result;
                    #endregion
            }

        }

        private string Row2outputDataType(DataRow Row, DataTable Table, string outputDataType)
        {
            string Result = string.Empty;

            string sheetName = Row[Table.Columns["sheetName"]].ToString();
            int sheetNum = (int)Row[Table.Columns["sheetNum"]];
            int rowNum = (int)Row[Table.Columns["rowNum"]];
            string marking = Row[Table.Columns["marking"]].ToString();
            string marking2 = Row[Table.Columns["marking2"]].ToString();
            string ggname = Row[Table.Columns["ggname"]].ToString();
            string currency = Row[Table.Columns["currency"]].ToString();
            double price1 = (double)Row[Table.Columns["price1"]];
            double price1USD = (double)Row[Table.Columns["price1USD"]];
            double price1UAH = (double)Row[Table.Columns["price1UAH"]];

            switch (outputDataType)
            {
                case "dictionarylist":
                    Result += (Result.Length == 0 ? "" : ",") +
                        "{" +
                        "\"sheetName\":\"" + sheetName + "\"," +
                        "\"sheetNum\":" + sheetNum + "," +
                        "\"rowNum\":" + rowNum + "," +
                        "\"marking\":\"" + marking + "\"," +
                        "\"marking2\":\"" + marking2 + "\"," +
                        "\"ggname\":\"" + ggname + "\"," +
                        "\"currency\":\"" + currency + "\"," +
                        "\"price1\":" + price1 + "," + 
                        "\"price1USD\":" + price1USD + "," +
                        "\"price1UAH\":" + price1UAH +
                        '}';
                    break;
                default:
                    Result += (Result.Length == 0 ? "" : "\n") +
                        sheetName + '\t' +
                        sheetNum + '\t' +
                        rowNum + '\t' +
                        marking + '\t' +
                        marking2 + '\t' +
                        ggname + '\t' +
                        currency + '\t' +
                        price1 + '\t' +
                        price1USD + '\t' +
                        price1UAH
                        ;
                    break;
            }
            return Result;
        }
        private string Rows2outputDataType(DataRowCollection Rows, DataTable Table, string outputDataType)
        {
            string Result = string.Empty;
            foreach (DataRow Row in Rows)
            {
                switch (outputDataType)
                {
                    case "dictionarylist":
                        Result += (Result.Length == 0 ? "" : ",") +
                            Row2outputDataType(Row, Table, outputDataType);
                        break;
                    default:
                        Result += (Result.Length == 0 ? "" : "\n") +
                            Row2outputDataType(Row, Table, outputDataType);
                        break;
                }
            }
            return Result;
        }
        private string GetPricesFormDataSetSheet(DataTable Prices, string outputDataType)
        {
            string Result = Rows2outputDataType(Prices.Rows, Prices, outputDataType);
            /*
            foreach (DataRow Row in Prices.Rows)
            {
                switch (outputDataType)
                {
                    case "dictionarylist":
                        Result += (Result.Length == 0 ? "" : ",") +
                            Rows2outputDataType(Row, Prices, outputDataType);
                        break;
                    default:
                        Result += (Result.Length == 0 ? "" : "\n") +
                            Rows2outputDataType(Row, Prices, outputDataType);
                        break;
                }
              
                string sheetName = Row[Prices.Columns["sheetName"]].ToString();
                int sheetNum = (int) Row[Prices.Columns["sheetNum"]];
                int rowNum = (int) Row[Prices.Columns["rowNum"]];
                string marking = Row[Prices.Columns["marking"]].ToString();
                string marking2 = Row[Prices.Columns["marking2"]].ToString();
                string ggname = Row[Prices.Columns["ggname"]].ToString();
                double price1 = (double) Row[Prices.Columns["price1"]];
                switch (outputDataType)
                {
                    case "dictionarylist":
                        Result += (Result.Length == 0 ? "" : ",") + 
                            "{" +
                            "\"sheetName\":\"" + sheetName + "\"," +
                            "\"sheetNum\":" + sheetNum + "," +
                            "\"rowNum\":" + rowNum + "," +
                            "\"marking\":\"" + marking + "\"," +
                            "\"marking2\":\"" + marking2 + "\"," +
                            "\"ggname\":\"" + ggname + "\"," + 
                            "\"price1\":" + price1 +  
                            '}'
                            ;
                        break;
                    default:
                        Result +=
                            sheetName + '\t' +
                            sheetNum + '\t' +
                            rowNum + '\t' +
                            marking + '\t' +
                            marking2 + '\t' +
                            ggname + '\t' + 
                            price1 + '\n';
                        break;
                }
                
            }
            */
            return Result;
        }
    }
}