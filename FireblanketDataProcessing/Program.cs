using ClosedXML.Excel;
using FireblanketDataProcessing;

/*
 * Directions:
 * Take the big CSV file and open a terminal into the folder it lives in
 * Give it this: split -l 1000000 data.csv data --additional-suffix=.csv
 * Then manually format all the new CSV files to a general format
 * Rename them and save as xlsx
 * Use that directory as the path here
 * 
 * Caution:
 * This thing requires a metric ton ram, like 10gb maybe (maybe more)
 */

List<string> boro = new List<string>();
List<string> block = new List<string>();
List<string> lot = new List<string>();
List<string> zip = new List<string>();
List<string> address = new List<string>();
List<string> easement = new List<string>();
List<string> value = new List<string>();
List<string> year = new List<string>();
List<string> neighborhood = new List<string>();
List<string> latitude = new List<string>();
List<string> longitude = new List<string>();

int lowestYear = 2100;
int highestYear = 1900;

List<Property> finalPropList = new List<Property>();

Console.WriteLine("Enter path:");
string folderPath = "C:\\Users\\kcdod\\Downloads\\files";
string outputPath = "C:\\Users\\kcdod\\Downloads\\dataFile.xlsx";
startTask(folderPath);
makeProperties();
finalPropList.Sort();
printToSheet(outputPath);

void startTask(string path)
{
    //Using async to read in the files might be time efficient if the system has 128ish gb of ram available
    bool useAsync = false;

    string[] files = Directory.GetFiles(path);
    List<List<Property>> propertyLists = new List<List<Property>>();

    if (useAsync)
    {
        List<Task<List<Property>>> taskList = new List<Task<List<Property>>>();

        for (int i = 0; i < files.Length; i++)
        {
            Task<List<Property>> t = Task.Run(() =>
            {
                return gatherData(files[i]);
            });
            taskList.Add(t);
        }

        Task tk = Task.WhenAll(taskList);
        tk.Wait();

        for (int i = 0; i < taskList.Count(); i++)
        {
            propertyLists.Add(taskList[i].Result);
        }
    }
    else
    {
        for (int i = 0; i < files.Length; i++)
        {
            propertyLists.Add(gatherData(files[i]));
        }
    }

    List<Property> propertyList = new List<Property>();

    for (int i = 0; i < propertyLists.Count(); i++)
    {
        propertyList.AddRange(propertyLists[i]);
    }
}

List<Property> gatherData(string path)
{
    SheetProcessor sp = new SheetProcessor();

    List<Property> propertyList = new List<Property>();

    //These are 1 indexed
    //Referenced from the sheet
    int col_boro = 2;
    int col_block = 3;
    int col_lot = 4;
    int col_zip = 20;
    int col_address = 19;
    int col_easement = 5;
    int col_value = 13;
    int col_year = 30;
    int col_neighborhood = 39;
    int col_latitude = 33;
    int col_longitude = 34;

    var wb = new XLWorkbook(path, XLEventTracking.Disabled);
    IXLWorksheet ws = wb.Worksheet(1);

    int cells = sp.getAmountCellsInColumn(col_boro, ref ws);
    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_boro, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        boro.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_block, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        block.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_lot, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        lot.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_zip, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        zip.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_address, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        address.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_easement, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        easement.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_value, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        value.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_year, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        year.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_neighborhood, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        neighborhood.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_latitude, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        latitude.Add(data);
    }

    for (int j = 1; j < cells; j++)
    {
        string data = sp.readSheetData(col_longitude, j, ref ws);
        if (data == "")
        {
            data = "N/A";
        }
        longitude.Add(data);
    }

    //Critically vital to not running out of ram
    wb.Dispose();
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();

    return propertyList;
}

void makeProperties()
{
    List<Property> props = new List<Property>();

    for (int i = 0; i < boro.Count(); i++)
    {
        Property prop = new Property();

        if (!isEasement(easement[i]))
        {
            prop.BoroughNumber = boro[i];
            prop.Block = block[i];
            prop.Lot = lot[i];
            prop.ZIP = zip[i];
            prop.Address = address[i];
            prop.Latitude = latitude[i];
            prop.Longitude = longitude[i];
            prop.Neighborhood = neighborhood[i];
            prop.Value = value[i];
            prop.Year = convertYear(year[i]);

            prop.PARID = makePARID(boro[i], block[i], lot[i]);
            prop.BoroughName = makeBorough(boro[i]);

            props.Add(prop);
        }
    }

    finalPropList.AddRange(combineProperties(props));
}

bool isEasement(string value)
{
    if (value != "N/A")
    {
        return true;
    }
    else
    {
        return false;
    }
}

string makePARID(string boro, string block, string lot)
{
    string parid = boro;

    if (block.Length == 1)
    {
        parid += ("0000" + block);
    }
    else if (block.Length == 2)
    {
        parid += ("000" + block);
    }
    else if (block.Length == 3)
    {
        parid += ("00" + block);
    }
    else if (block.Length == 4)
    {
        parid += ("0" + block);
    }
    else if (block.Length == 5)
    {
        parid += block;
    }

    if (lot.Length == 1)
    {
        parid += ("000" + lot);
    }
    else if (lot.Length == 2)
    {
        parid += ("00" + lot);
    }
    else if (lot.Length == 3)
    {
        parid += ("0" + lot);
    }
    else if (lot.Length == 4)
    {
        parid += lot;
    }

    return parid;
}

string makeBorough(string boro)
{
    string borough = "";

    if (boro == "1")
    {
        borough = "Manhattan";
    }
    else if (boro == "2")
    {
        borough = "Bronx";
    }
    else if (boro == "3")
    {
        borough = "Brooklyn";
    }
    else if (boro == "4")
    {
        borough = "Queens";
    }
    else if (boro == "5")
    {
        borough = "Staten Island";
    }

    return borough;
}

string convertYear(string year)
{
    string yr = year.Substring(0, 4);

    int yearInt = Convert.ToInt16(yr);

    if (yearInt > highestYear)
    {
        highestYear = yearInt;
    }

    if (yearInt < lowestYear)
    {
        lowestYear = yearInt;
    }

    return yr;
}

List<Property> combineProperties(List<Property> properties)
{
    properties.Sort();
    List<Property> finishedList = new List<Property>();

    List<List<Property>> subLists = new List<List<Property>>();

    Property lastProperty = properties[0];
    int start = 0;
    int count = 1;

    for (int i = 1; i < properties.Count(); i++)
    {
        if (properties[i].PARID == lastProperty.PARID)
        {
            count++;
        }
        else
        {
            List<Property> propList = new List<Property>(properties.GetRange(start, count));
            subLists.Add(propList);
            start = i;
            count = 1;
            lastProperty = properties[i];
        }
    }

    List<Task<Property>> taskList = new List<Task<Property>>();

    for (int i = 0; i < subLists.Count(); i++)
    {
        int iterator = i;
        List<Property> subList = subLists[iterator];

        Task<Property> t = Task.Run(() =>
        {
            return doCombine(subList);
        });
        taskList.Add(t);
    }

    Task tk = Task.WhenAll(taskList);
    tk.Wait();

    for (int i = 0; i < taskList.Count(); i++)
    {
        finishedList.Add(taskList[i].Result);
    }

    for (int i = 0; i < finishedList.Count(); i++)
    {
        finishedList[i].YearsAndValues = finishedList[i].YearsAndValues.OrderBy(years => years.Item1).ToList();
    }

    finishedList.Sort();

    return finishedList;
}

Property doCombine(List<Property> properties)
{
    Property prop = properties[0];

    for (int i = 1; i < properties.Count(); i++)
    {
        Tuple<string, string> yearValue = new Tuple<string, string>(properties[i].Year, properties[i].Value);
        prop.YearsAndValues.Add(yearValue);

        if ((prop.ZIP == "" || prop.ZIP == "N/A") && (properties[i].ZIP != "" || properties[i].ZIP != "N/A"))
        {
            prop.ZIP = properties[i].ZIP;
        }

        if ((prop.Latitude == "" || prop.Latitude == "N/A") && (properties[i].Latitude != "" || properties[i].Latitude != "N/A"))
        {
            prop.Latitude = properties[i].Latitude;
        }

        if ((prop.Longitude == "" || prop.Longitude == "N/A") && (properties[i].Longitude != "" || properties[i].Longitude != "N/A"))
        {
            prop.Longitude = properties[i].Longitude;
        }

        if ((prop.Neighborhood == "" || prop.Neighborhood == "N/A") && (properties[i].Neighborhood != "" || properties[i].Neighborhood != "N/A"))
        {
            prop.Neighborhood = properties[i].Neighborhood;
        }
    }

    return prop;
}

void printToSheet(string path)
{
    using (File.Create(path)) { }
    SheetProcessor sp = new SheetProcessor();
    XLWorkbook wb = sp.createSheet("Locations");
    wb.Worksheets.Add("Locations 2");
    IXLWorksheet ws = wb.Worksheet(1);

    sp.writeSheetData(1, 1, "PARID:", ref ws);
    sp.writeSheetData(2, 1, "BoroughName:", ref ws);
    sp.writeSheetData(3, 1, "BoroughNumber:", ref ws);
    sp.writeSheetData(4, 1, "Block:", ref ws);
    sp.writeSheetData(5, 1, "Lot:", ref ws);
    sp.writeSheetData(6, 1, "Address:", ref ws);
    sp.writeSheetData(7, 1, "ZIP:", ref ws);
    sp.writeSheetData(8, 1, "Neighborhood:", ref ws);
    sp.writeSheetData(9, 1, "Latitude:", ref ws);
    sp.writeSheetData(10, 1, "Longitude:", ref ws);

    for (int i = 0; i < ((highestYear - lowestYear) + 1); i++)
    {
        sp.writeSheetData((i + 11), 1, ((lowestYear + i).ToString() + " Value:"), ref ws);
    }

    for (int i = 0; i < finalPropList.Count(); i++)
    {
        int iterator;
        if (i < 1048574)
        {
            iterator = i;
        }
        else
        {
            ws = wb.Worksheet(2);
            iterator = i - 1048574;
        }

        sp.writeSheetData(1, iterator + 2, finalPropList[i].PARID, ref ws);
        sp.writeSheetData(2, iterator + 2, finalPropList[i].BoroughName, ref ws);
        sp.writeSheetData(3, iterator + 2, finalPropList[i].BoroughNumber, ref ws);
        sp.writeSheetData(4, iterator + 2, finalPropList[i].Block, ref ws);
        sp.writeSheetData(5, iterator + 2, finalPropList[i].Lot, ref ws);
        sp.writeSheetData(6, iterator + 2, finalPropList[i].Address, ref ws);
        sp.writeSheetData(7, iterator + 2, finalPropList[i].ZIP, ref ws);
        sp.writeSheetData(8, iterator + 2, finalPropList[i].Neighborhood, ref ws);
        sp.writeSheetData(9, iterator + 2, finalPropList[i].Latitude, ref ws);
        sp.writeSheetData(10, iterator + 2, finalPropList[i].Longitude, ref ws);

        for (int j = 0; j < ((highestYear - lowestYear) + 1); j++)
        {
            string year = (lowestYear + j).ToString();

            for (int k = 0; k < finalPropList[i].YearsAndValues.Count(); k++)
            {
                if (finalPropList[i].YearsAndValues[k].Item1 == year)
                {
                        sp.writeSheetData((j + 11), iterator + 2, finalPropList[i].YearsAndValues[k].Item2, ref ws);
                }
            }
        }
    }

    sp.saveSheet(path, ref wb);
}