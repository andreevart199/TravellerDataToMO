using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.HSSF.Util;

namespace TravellerDataToMO
{
    class Program
    {
        static async Task Main(string[] args)
        {
            List<Traveller> listTraveller = await GetListTravellersFromFileXLSAsync();
            var codesMo = listTraveller.Select(x => x.КодМедОрганизации).Distinct().ToList();
            Parallel.ForEach(codesMo, (codeMo) => 
                { 
                    Task.Run(() => CreateFileWithTravellerAsync(codeMo, listTraveller.Where(x => x.КодМедОрганизации == codeMo).ToList()));
                });
            Console.ReadLine();
        }

        /// <summary>
        ///  Читаем файл и заполняем listTraveller данными из каждой строки
        /// </summary>
        /// <returns>
        /// </returns>
        public static async Task<List<Traveller>> GetListTravellersFromFileXLSAsync()
        {
            List<Traveller> listTraveller = new List<Traveller>();
            Console.WriteLine("Путь к файлу: ");
            string filenameRead = Console.ReadLine().Replace("\"", "");
            using (FileStream fs = new FileStream(filenameRead, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                NPOI.SS.UserModel.IWorkbook wb = null;
                NPOI.SS.UserModel.ISheet sh = null;
                try
                {
                    wb = new XSSFWorkbook(fs);
                    sh = (XSSFSheet)wb.GetSheetAt(0);
                }
                catch
                {
                    Console.WriteLine("Can't open file as a Workbook");
                    return null;
                }
                //первая строка с данными в текущем файле МВД 23 
                for (int row = 24; row <= sh.LastRowNum; row++)
                {
                    //Получаем фамилию из файла
                    string family = sh.GetRow(row).GetCell(1).StringCellValue;
                    //Получаем имя с отчеством из файла
                    string[] nameAndPatronymic = sh.GetRow(row).GetCell(2).StringCellValue.Split(' ');
                    //Получаем имя из массива nameAndPatronymic
                    string name = nameAndPatronymic[0];
                    //Получаем отчество из массива nameAndPatronymic                                    //на случай если кто то без отчества
                    string patronymic = nameAndPatronymic.Length < 3 ? nameAndPatronymic[1] : nameAndPatronymic[1] + " " + nameAndPatronymic[2];
                    //Получаем дату рождения из файла
                    string birthDate = sh.GetRow(row).GetCell(3).NumericCellValue.ToString();
                    //в файле источнике  дата представлена строкой 7 или 8 символов, вставляем точки в зависимости от одной или второй ситуации
                    string resBirth = birthDate.Length > 7 ? birthDate.Insert(2, ".").Insert(5, ".") : birthDate.Insert(1, ".").Insert(4, ".");
                    //Получаем Код Мед организации
                    string codeMO = await GetCodeMOFromFOMSAsync(family, name, patronymic, resBirth);
                    Traveller traveller = new Traveller()
                    {
                        //название полей идут в файл, названиями столбцов. поэтому на кириллице
                        //присаиваем фамилию
                        Фамилия = family,
                        //присваиваем имя
                        Имя = name,
                        //присваиваем отчество
                        Отчество = patronymic,
                        //присваиваем ДР
                        ДатаРождения = birthDate,
                        //Присваиваем НР(стоблец 8)
                        НР = sh.GetRow(row).GetCell(8).CellType == NPOI.SS.UserModel.CellType.Numeric ? sh.GetRow(row).GetCell(8).NumericCellValue.ToString() : sh.GetRow(row).GetCell(8).StringCellValue,
                        //приваиваем код медорганизации
                        КодМедОрганизации = String.IsNullOrEmpty(codeMO) ? "0" : codeMO
                    };
                    //добавляем заполненный объект в список 
                    listTraveller.Add(traveller);
                    Console.WriteLine(traveller.Фамилия + " " + traveller.КодМедОрганизации);
                }
            }
            return listTraveller;
        }

        /// <summary>
        /// Получаем код МО
        /// </summary>
        /// <param Фамилия="family"></param>
        /// <param Имя="name"></param>
        /// <param Отчество="patronymic"></param>
        /// <param ДР="birthDate"></param>
        /// <returns></returns>
        public static async Task<string> GetCodeMOFromFOMSAsync(string family, string name, string patronymic, string birthDate)
        {
            var client = new PatiVer.WcfServiceClient();
            var fomsData = await client.GetPersonInfo_FIOAsync("---", family, name, patronymic, birthDate, "---", "---", false, 11);
            string codeMO = (fomsData.SearchResult == "0" || fomsData.SearchResult == "3") ? "0" : fomsData.AttachmentData.CodeMO;
            return codeMO;
        }
        /// <summary>
        /// Создаем файл с путешественниками 
        /// </summary>
        /// <param name="codeMO"></param>
        /// <param name="listTraveller"></param>
        public static async Task CreateFileWithTravellerAsync(string codeMO, List<Traveller> listTraveller)
        {
            string nameMO = await GetNameMOAsync(codeMO);
            PropertyInfo[] properties = new Traveller().ReturnType();
            //выкидываем невалидные для имени файла символы
            nameMO = string.Join("_", nameMO.Split(Path.GetInvalidFileNameChars()));
            string path = $@"C:\Working\{nameMO}.xlsx";
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet("Sheet1");
                ICreationHelper cH = wb.GetCreationHelper();
                IRow hederRow = sheet.CreateRow(0);
                //заполняем первую строку с названиями полей
                for (int i = 0; i < properties.Length; i++)
                {
                    ICell cell = hederRow.CreateCell(i);
                    cell.CellStyle.FillBackgroundColor = 1;
                    cell.SetCellValue(properties[i].Name);
                }
                //заполняем строки с данными о прибывших
                for (int i = 0; i < listTraveller.Count; i++)
                {
                    IRow row = sheet.CreateRow(i);
                    for (int j = 0; j < properties.Length; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        cell.SetCellValue(listTraveller[i].GetType().GetProperty(properties[j].Name).GetValue(listTraveller[i], null).ToString());
                    }
                }
                wb.Write(stream);
            }
        }
        /// <summary>
        /// Получаем имя МО по коду МО
        /// </summary>
        /// <param name="codeMO"></param>
        /// <returns></returns>
        public static async Task<string> GetNameMOAsync(string codeMO)
        {
            string nameMO = "";
            string connectionString = @"connectionString";
            string sqlLoadNameMOexpression = String.Format("select top 1 NAM_MOK from NSI_F003 WHERE MCOD = {0}", codeMO);
            using (SqlConnection connection =  new SqlConnection(connectionString))
            {
                await connection.OpenAsync();
                using (SqlCommand command = new SqlCommand(sqlLoadNameMOexpression, connection))
                using (SqlDataReader reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        nameMO = reader.GetValue(0).ToString();
                    }
                }
            }
            return nameMO;
        }
    }
}
