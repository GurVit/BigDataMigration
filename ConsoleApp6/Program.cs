using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;



namespace ConsoleApp6
{
    class Program
    {
        static string filepath = @"D:\File.xlsx";
        static readonly Random random = new Random();
        static ExcelPackage excelPackage = new ExcelPackage();
        static ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
        
        static void Main(string[] args)
        {
            Console.WriteLine("Запись в Excel началась");
            FirstService();
            Console.WriteLine("Запись в Excel закончена. Чтобы начать запись в базу данных нажмите любую кнопку");
            Console.ReadKey();
            SecondService();
            Console.WriteLine("Запись в базу данных закончена");
        }

        static void FirstService()
        {
            excelPackage.Workbook.Properties.Author = "VitG";
            excelPackage.Workbook.Properties.Created = DateTime.Now;

            worksheet.Cells[1, 1].Value = "Фамилия";
            worksheet.Cells[1, 2].Value = "Имя";
            worksheet.Cells[1, 3].Value = "Отчество";
            worksheet.Cells[1, 4].Value = "Телефон";
            worksheet.Cells[1, 5].Value = "Адрес";

            for (int i = 2; i < 200002; i++)
            {
                ToExcel(i);
            }

            FileInfo fin = new FileInfo(filepath);
            excelPackage.SaveAs(fin);
        }

        static void SecondService()
        {
            SqlConnection myConn = new SqlConnection("Server=localhost;Integrated security=SSPI;database=master");
            Workbook workbook = new Workbook(filepath);
            workbook.Save(@"D:\File.csv", SaveFormat.CSV);
            Console.WriteLine("CSV is ready");
            string strDB;
            strDB = "CREATE DATABASE VSKDatabase ON PRIMARY " +
             "(NAME = VSKDatabase_Data, " +
             "FILENAME = 'D:\\VSKDatabaseData.mdf', " +
             "SIZE = 50MB, MAXSIZE = 50MB, FILEGROWTH = 10%)" +
             "LOG ON (NAME = VSKDatabase_Log, " +
             "FILENAME = 'D:\\VSKDatabaseLog.ldf', " +
             "SIZE = 50MB, " +
             "MAXSIZE = 50MB, " +
             "FILEGROWTH = 10%)";

            SqlCommand CommandCreateDB = new SqlCommand(strDB, myConn);
            try
            {
                myConn.Open();
                CommandCreateDB.ExecuteNonQuery();
                string strTable = "USE VSKDatabase CREATE TABLE People (Фамилия VARCHAR(20), Имя VARCHAR(20), Отчество VARCHAR(20), Телефон VARCHAR(20), Адрес VARCHAR(20))";
                SqlCommand CommandCreateTable = new SqlCommand();
                CommandCreateTable.Connection = myConn;
                CommandCreateTable.CommandText = strTable;
                CommandCreateTable.ExecuteNonQuery();
                Console.WriteLine("DataBase is Created Successfully");
                string strBULK = "USE VSKDatabase BULK INSERT People FROM 'D:\\File.csv' WITH(FIRSTROW=2, LASTROW=200001, BATCHSIZE=40000, FIELDTERMINATOR = ',', ROWTERMINATOR = '0x0A'); ";
                SqlCommand CommandBULK = new SqlCommand(strBULK, myConn);
                CommandBULK.ExecuteNonQuery();
                Console.WriteLine("DataSet is recorded");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if (myConn.State == ConnectionState.Open)
                {
                    myConn.Close();
                }
                Console.ReadKey();
            }
        }

        static async void ToExcel(int i)
        {
            worksheet.Cells[i, 1].Value = await GetWord();
            worksheet.Cells[i, 2].Value = await GetWord();
            worksheet.Cells[i, 3].Value = await GetWord();
            worksheet.Cells[i, 4].Value = random.Next(100000, 1000000);
            worksheet.Cells[i, 5].Value = await GetWord();
        }

        static async Task<string> GetWord()
        {
            string word = "";
            var r = new Random();
            while (word.Length < 10)
            {
                char c = (char)r.Next(33, 125);
                if (char.IsLetterOrDigit(c))
                    word += c;
            }
            return word;
        }
    }
}
