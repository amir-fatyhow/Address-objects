using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection.Emit;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using System.Net;
using static System.Runtime.InteropServices.JavaScript.JSType;

public class XmlParser
{
    public static void Main(string[] args)
    {
        const string URL = "https://fias.nalog.ru/Public/Downloads/Actual/gar_delta_xml.zip";
        const string SAVE_PATH = "D:\\data.zip";

        const string LEVEL_NAME = "AS_OBJECT_LEVELS_"; 
        const string FOLDERS_OF_OBJECT_PATH = "D:\\data"; 
        const string TARGET_EXCEL_PATH = "D:\\tables.xlsx"; // путь, по которому записываем данные в Excel таблицу

        const string ZIP_PATH = "D:\\data.zip"; 
        const string EXTRACT_PATH = "D:\\data";

        const string TARGET_PATH_READING = "AS_ADDR_OBJ_"; // паттерн названий файлов, которые необходимо считать "AS_ADDR_OBJ_20241007"

        try
        {
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(URL, SAVE_PATH);
            }


            System.IO.Compression.ZipFile.ExtractToDirectory(ZIP_PATH, EXTRACT_PATH);



            List<ObjectLevelData> levels = getLevelTypes(FOLDERS_OF_OBJECT_PATH, LEVEL_NAME);
            List<ObjectData> objects = getObjects(FOLDERS_OF_OBJECT_PATH, TARGET_PATH_READING);

            // получаем списки разбитые по группам по дате изменения (параметр updatedate у OBJECT)
            List<List<ObjectData>> objectsByUpdatetime = objects.Select((x, i) => new { Index = i, Value = x })
                .GroupBy(x => x.Value.UPDATEDATE)
                .Select(x => x.Select(v => v.Value).ToList())
                .ToList();

            writeToExcel(objectsByUpdatetime, levels, TARGET_EXCEL_PATH);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
    }

    public static List<ObjectLevelData> getLevelTypes(string pathToFolder, string pathToLevelFile)
    {
        Regex regex = new Regex($"{pathToLevelFile}*");  
        IEnumerable<string> levelType = Directory.EnumerateFiles(pathToFolder).Where(f => regex.IsMatch(f));
        XDocument docLevels = XDocument.Load(levelType.First());

        List<ObjectLevelData> levels = new List<ObjectLevelData>();
        foreach (var levelElement in docLevels.Descendants("OBJECTLEVEL"))
        {
            // Извлечение данных из атрибутов
            int level = int.Parse(levelElement.Attribute("LEVEL").Value);
            string name = levelElement.Attribute("NAME").Value;
            ObjectLevelData obj = new ObjectLevelData(level, name);
            levels.Add(obj);
        }
        return levels;
    }

    public static List<ObjectData> getObjects(string pathToAllFoldersOfObject, string nameOfFilePattern)
    {
        Regex regex = new Regex($"{nameOfFilePattern}[0-9]*");
        string[] allFolders = Directory.GetDirectories(pathToAllFoldersOfObject);
        IEnumerable<string> matches = Enumerable.Empty<string>();
        foreach (string folder in allFolders)
        {
            IEnumerable<string> match = Directory.EnumerateFiles(folder).Where(f => regex.IsMatch(f));
            matches = matches.Concat(match);
        }

        List<ObjectData> objects = new List<ObjectData>();

        if (matches.Count() != 0)
        {
            foreach (var m in matches)
            {
                XDocument doc = XDocument.Load(m);
                foreach (var objectElement in doc.Descendants("OBJECT"))
                {
                    int isActive = int.Parse(objectElement.Attribute("ISACTIVE").Value);
                    if (isActive == 1)
                    {
                        string name = objectElement.Attribute("NAME").Value;
                        string typeName = objectElement.Attribute("TYPENAME").Value;
                        int level = int.Parse(objectElement.Attribute("LEVEL").Value);
                        DateTime updateDate = DateTime.Parse(objectElement.Attribute("UPDATEDATE").Value);
                        ObjectData obj = new ObjectData(name, typeName, level, updateDate, isActive);
                        objects.Add(obj);
                    }
                }
            }
        }
        return objects;
    }

    public static void writeToExcel(List<List<ObjectData>> objectsByUpdatetime, List<ObjectLevelData> levels, string targetPath)
    {
        Application excelApp = new Application();
        Workbook workbook = excelApp.Workbooks.Add();
        foreach (var obj in objectsByUpdatetime)
        {

            List<List<ObjectData>> objectsByName = obj.Select((x, i) => new { Index = i, Value = x })
                                                        .GroupBy(x => x.Value.LEVEL)
                                                        .Select(x => x.Select(v => v.Value).ToList())
                                                        .ToList();

            Worksheet worksheet = workbook.Worksheets.Add();
            worksheet.Name = $"{objectsByName.First().First().UPDATEDATE.ToShortDateString()}";

            worksheet.Cells[1, 1] = $"Отчет по добавленным адресным объектам за {objectsByName.First().First().UPDATEDATE.ToShortDateString()}";
            worksheet.Cells[3, 1] = "Тип объекта";
            worksheet.Cells[3, 2] = "Наименование";

            foreach (var objectByName in objectsByName)
            {
                int row = 4;
                List<ObjectData> sortedObjects = objectByName.OrderBy(o => o.NAME).ToList();
                ObjectData firstObject = sortedObjects.First();
                string level = "";
                foreach (var l in levels)
                {
                    if (l.LEVEL == firstObject.LEVEL)
                    {
                        level = $"{l.NAME}";
                        worksheet.Cells[2, 1] = level;
                        break;
                    }
                }

                foreach (var o in sortedObjects)
                {
                    worksheet.Cells[row, 1] = $"{o.TYPENAME}";
                    worksheet.Cells[row, 2] = $"{o.NAME}";
                    row++;
                }

            }
        }
        workbook.SaveAs(targetPath);
        excelApp.Quit();
    }
}

// Класс для хранения данных объекта
public class ObjectData
{
    public string NAME { get; set; }
    public string TYPENAME { get; set; }
    public int LEVEL { get; set; }
    public DateTime UPDATEDATE { get; set; }
    public int ISACTIVE { get; set; }

    public ObjectData(string name, string typeName, int level, DateTime updateDate, int isActive)
    {
        NAME = name;
        TYPENAME = typeName;
        LEVEL = level;
        UPDATEDATE = updateDate;
        ISACTIVE = isActive;
    }
}

// Класс для уровней адресообразующих объектов
public class ObjectLevelData
{
    public int LEVEL { get; set; }
    public string NAME { get; set; }

    public ObjectLevelData(int level, string name)
    {
        LEVEL = level;
        NAME = name;
    }
}