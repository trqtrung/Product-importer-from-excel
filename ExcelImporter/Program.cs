using System;
using ExcelDataReader;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data;
using System.Text;
using MySql.Data.MySqlClient;

namespace ExcelImporter
{
    

    public class Program
    {
        static string version = "v20180408";
        static string sourcePath = @"C:\\Excel\";
        static string ttaMariaDBConnectionString = @"Database=tta;Data Source=127.0.0.1;User Id=root;Password=trung1992;SslMode=none";

        static void Main(string[] args)
        {
            Console.WriteLine("Excel Importer - " + version);
            Console.WriteLine("this thing is made by .net core");
            Console.Title = "Excel Importer - " + version;

            //TestConnection();

            if (!Directory.Exists(sourcePath))
                Directory.CreateDirectory(sourcePath);

            var extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    ".xls",
    ".xlsx",
};

            IEnumerable<string> files = Directory.EnumerateFiles(sourcePath).Where(filename => extensions.Contains(Path.GetExtension(filename)));

            if (!files.Any())
            {
                Console.WriteLine("No excel files found in " + sourcePath);
            }
            else
            {
                foreach (string f in files)
                {
                    if (File.Exists(f))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (var stream = File.Open(f, FileMode.Open, FileAccess.Read))
                        {
                            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                            //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                            //...
                            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                            //IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                            //IExcelDataReader excelReader;

                            //1. Reading Excel file
                            //if (Path.GetExtension(f).ToUpper() == ".XLS")
                            //{
                            //    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                            //    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                            //}
                            //else
                            //{
                            //    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                            //    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                            //}

                            ////2. DataSet - The result of each spreadsheet will be created in the result.Tables
                            //DataSet result = excelReader.AsDataSet();

                            ////3. DataSet - Create column names from first row
                            //excelReader.IsFirstRowAsColumnNames = false;

                            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                            {

                                //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                                //DataSet result = reader.AsDataSet();
                                ////...
                                ////4. DataSet - Create column names from first row
                                //excelReader.IsFirstRowAsColumnNames = true;
                                //DataSet result = excelReader.AsDataSet();
                                int i = 0;

                                Console.OutputEncoding = Encoding.UTF8;

                                while (reader.Read())
                                {

                                    // 1. Use the reader methods
                                    //do
                                    //{
                                    Console.WriteLine(i);
                                    ColumnIndexItem col = new ColumnIndexItem();
                                    col = null;

                                    for (int j = 0; j < 18; j++)
                                    {
                                        Console.Write(reader.GetValue(j) + " - ");

                                        while (col == null)
                                        {
                                            switch (reader.GetValue(j).ToString().ToLower())
                                            {
                                                case "product name":
                                                    col.ProductName = j;
                                                    break;
                                                case "brand":
                                                    col.Brand = j;
                                                    break;
                                                case "type":
                                                    col.Type = j;
                                                    break;
                                                case "colour":
                                                    col.Colour = j;
                                                    break;
                                                case "size inch":
                                                    col.SizeInch = j;
                                                    break;
                                                case "size mm":
                                                    col.SizeMM = j;
                                                    break;
                                                case "price":
                                                    col.Price = j;
                                                    break;
                                                case "unit":
                                                    col.Unit = j;
                                                    break;
                                                case "origin":
                                                    col.Origin = j;
                                                    break;
                                                case "supplier":
                                                    col.Supplier = j;
                                                    break;
                                                case "height":
                                                    col.Height = j;
                                                    break;
                                                case "width":
                                                    col.Width = j;
                                                    break;
                                                case "length":
                                                    col.Length = j;
                                                    break;
                                                case "weight":
                                                    col.Weight = j;
                                                    break;
                                                case "code":
                                                    col.Code = j;
                                                    break;
                                            }
                                        
                                        }
                                        
                                    }
                                    //if(i > 3 && i < 6)
                                    //{
                                    //    ProductItem item = new ProductItem();
                                    //    item.Name = reader.GetString(0);
                                    //    item.Price = reader.GetDouble(5);
                                    //    item.Type = 7;

                                    //    if (InsertProduct(item) > 0)
                                    //        Console.WriteLine("inserted");
                                    //    else
                                    //        Console.WriteLine("oops");
                                    //}
                                    //Console.WriteLine(i + " - " +reader.GetString(0) + " - " + reader.GetString(1) + " - "+ reader.GetString(2) +" - "+ reader.GetString(3));
                                    i++;

                                    //} while (reader.NextResult());
                                }
                                // 2. Use the AsDataSet extension method
                                //var result = reader.AsDataSet();

                                // The result of each spreadsheet is in result.Tables
                            }

                        }
                    }
                }
            }
            Console.WriteLine("Finished Importing. Press any key to exit program.");
            Console.ReadLine();
        }

        static long InsertProduct(ProductItem item)
        {
            try
            {
                using (MySqlConnection connection = new MySqlConnection(ttaMariaDBConnectionString))
                {
                    connection.Open();
            //        Console.WriteLine("ServerVersion: " + connection.ServerVersion +
            //"\nState: " + connection.State.ToString());
                    MySqlCommand cmd = new MySqlCommand("INSERT INTO products (name, price, type) VALUES (@name,@price,@type)", connection);
                    cmd.Parameters.Add("@name", item.Name);
                    cmd.Parameters.Add("@type", item.Type);
                    cmd.Parameters.Add("@price", item.Price);
                    cmd.Prepare();

                    cmd.ExecuteNonQuery();
                    long id = cmd.LastInsertedId;
                    connection.Close();
                    return id;
                }
            }
            catch(Exception ex)
            {

            }
            return 0;
        }

        public static int GetValueID(string name, string key)
        {
            string sql = "SELECT id FROM options_lists WHERE name = @name AND key =@key";
            using (MySqlConnection connection = new MySqlConnection(ttaMariaDBConnectionString))
            {
                connection.Open();

                MySqlCommand cmd = new MySqlCommand(sql, connection);
                cmd.Parameters.AddWithValue("@name", name);
                cmd.Parameters.AddWithValue("@key", key);
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        reader.Read();

                        return Convert.ToInt32(reader["id"]);
                    }
                }
                connection.Close();
            }
            return 0;
        }

        static long InsertOption(OptionItem item)
        {
            try
            {
                using (MySqlConnection connection = new MySqlConnection(ttaMariaDBConnectionString))
                {
                    connection.Open();

                    MySqlCommand cmd = new MySqlCommand("INSERT INTO options_lists (name, key, value) VALUES (@name,@key,@value)", connection);

                    cmd.Parameters.Add("@name", item.Name);
                    cmd.Parameters.Add("@key", item.Key);
                    cmd.Parameters.Add("@value", item.Value);
                    cmd.Prepare();

                    cmd.ExecuteNonQuery();
                    long id = cmd.LastInsertedId;
                    connection.Close();
                    return id;
                }
            }
            catch (Exception ex)
            {

            }
            return 0;
        }

        public static bool CheckProductExists(ProductItem item)
        {
            string sql = "SELECT id FROM products WHERE brand = @brand AND type=@type AND name LIKE @name";

            return false;
        }
    }

    public class ProductItem
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public double Price { get; set; }
        public string SKU { get; set; }
        public string Colour { get; set; }
        public double Weight { get; set; }
        public double Size { get; set; }
        public int Type { get; set; }
    }
    public class ColumnIndexItem
    {
        public int ProductName { get; set; }
        public int Code { get; set; }
        public int Type { get; set; }
        public int Brand { get; set; }
        public int SizeInch { get; set; }
        public int SizeMM { get; set; }
        public int Price { get; set; }
        public int Unit { get; set; }
        public int Origin { get; set; }
        public int Supplier { get; set; }
        public int Colour { get; set; }
        public int Height { get; set; }
        public int Width { get; set; }
        public int Length { get; set; }
        public int Weight { get; set; }
    }
    public class OptionItem
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Key { get; set; }
        public string Value { get; set; }
    }
}