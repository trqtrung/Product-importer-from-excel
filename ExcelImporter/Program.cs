using ExcelDataReader;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelImporter
{


    public class Program
    {
        static string version = "v20180418";
        static string sourcePath = @"C:\\Excel\";
        static string ttaMariaDBConnectionString = @"Database=tta;Data Source=127.0.0.1;User Id=trung;Password=123456;SslMode=none";

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
                                int i = -1;

                                Console.OutputEncoding = Encoding.UTF8;
                                ColumnIndexItem col = new ColumnIndexItem();
                                col.ProductName = -1;
                                
                                while (reader.Read())
                                {
                                    i++;
                                    // 1. Use the reader methods
                                    //do
                                    //{


                                    if (col.ProductName < 0)
                                    {
                                        Console.WriteLine(i);
                                        int fields = reader.FieldCount;
                                        for (int j = 0; j < fields; j++)
                                        {
                                            string colName = "";

                                            if (reader.GetValue(j) == null)
                                                continue;
                                            else
                                                colName = reader.GetValue(j).ToString();

                                            if (j == 0 && colName.ToLower() != "product name")
                                                break;

                                            switch (colName.ToLower())
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
                                                case "sub type":
                                                    col.SubType = j;
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
                                    else
                                    {
                                        ProductItem p = new ProductItem();

                                        p.Name = reader.GetString(col.ProductName);
                                        p.TypeName = reader.GetString(col.Type);
                                        p.BrandName = reader.GetString(col.Brand);

                                        if (string.IsNullOrEmpty(p.Name) || string.IsNullOrEmpty(p.TypeName) || string.IsNullOrEmpty(p.BrandName))
                                        {
                                            Console.WriteLine(i + " - Missing Name or Type or Brand. Skip record.");
                                            continue;
                                        }

                                        //p.Name = p.Name.Replace("  ", " ").Trim();
                                        p.Name = Regex.Replace(p.Name, @"\s+", " ").Trim();

                                        int typeID = GetValueID(p.TypeName.Trim(), "product.type","");

                                        if (typeID == 0)
                                        {
                                            OptionItem o = new OptionItem();
                                            o.Key = "product.type";
                                            o.Name = p.TypeName;

                                            typeID = (int)InsertOption(o);
                                        }
                                        p.TypeID = typeID;

                                        if(!reader.IsDBNull(col.SubType))
                                        {
                                            string subType = reader.GetString(col.SubType);

                                            int subTypeID = GetValueID(subType.Trim(), "product.subtype", typeID.ToString());

                                            if (subTypeID == 0)
                                            {
                                                OptionItem o = new OptionItem();
                                                o.Key = "product.subtype";
                                                o.Name = subType;
                                                o.Value = typeID.ToString();

                                                subTypeID = (int)InsertOption(o);
                                            }

                                            p.SubTypeID = subTypeID;
                                        }

                                        int brandID = GetValueID(p.BrandName.Trim(), "product.brand","");
                                        if (brandID == 0)
                                        {
                                            OptionItem o = new OptionItem();
                                            o.Key = "product.brand";
                                            o.Name = p.BrandName;

                                            brandID = (int)InsertOption(o);
                                        }
                                        p.BrandID = brandID;

                                        bool exists = CheckProductExists(p);

                                        if (exists)
                                            continue;

                                        string supplier = "";

                                        if (reader.GetValue(col.Supplier) == null)
                                        {
                                            supplier = "Unknown";
                                        }
                                        else
                                        {

                                            supplier = reader.GetValue(col.Supplier).ToString();
                                        }

                                        int supID = FindSupplierID(supplier);

                                        if (supID == 0)
                                        {
                                            SupplierItem s = new SupplierItem();
                                            s.Name = "";
                                            s.Phone = "";

                                            if (IsEnglishLetter(supplier[0]))
                                                s.Name = supplier;
                                            else if (!IsEnglishLetter(supplier[supplier.Length - 1]))
                                                s.Phone = supplier;

                                            supID = (int)InsertSupplier(s);
                                        }

                                        p.SupplierID = supID;

                                        if (!reader.IsDBNull(col.Length))
                                            p.Length = reader.GetDouble(col.Length);

                                        if (!reader.IsDBNull(col.Height))
                                            p.Height = reader.GetDouble(col.Height);

                                        if (!reader.IsDBNull(col.Weight))
                                            p.Weight = reader.GetDouble(col.Weight);

                                        if (!reader.IsDBNull(col.Width))
                                            p.Width = reader.GetDouble(col.Width);

                                        if (!reader.IsDBNull(col.SizeMM))
                                            p.Size = reader.GetDouble(col.SizeMM);

                                        if (!reader.IsDBNull(col.Colour))
                                            p.Colour = reader.GetString(col.Colour);

                                        if (!reader.IsDBNull(col.Code))
                                            p.Code = reader.GetString(col.Code);

                                        if(!reader.IsDBNull(col.Price))
                                        {
                                            p.Price = reader.GetDouble(col.Price) * 1000;
                                            p.PriceDate = DateTime.Now;
                                        }

                                        long recordID = InsertProduct(p);

                                        if (recordID > 0)
                                        {
                                            p.ID = (int)recordID;
                                            

                                            InsertPrice(p);

                                            Console.WriteLine(i + " - Inserted " + p.Name);
                                        }
                                        else
                                        {
                                            Console.WriteLine(i+ " - Error - " + p.Name);
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
                    MySqlCommand cmd = new MySqlCommand("INSERT INTO products (name, price, type, sub_type, sku, details, colours, weight, height, width, length, size, code, brand, created) VALUES (@name,@price,@type,@sub_type, @sku, @details, @colours, @weight, @height, @width, @length, @size, @code, @brand, NOW())", connection);
                    cmd.Parameters.Add("@name", item.Name);
                    cmd.Parameters.Add("@type", item.TypeID);
                    cmd.Parameters.Add("sub_type", item.SubTypeID);
                    cmd.Parameters.Add("@price", item.Price);

                    cmd.Parameters.Add("@sku", item.SKU);
                    cmd.Parameters.Add("@details", item.Details);
                    cmd.Parameters.Add("@colours", item.Colour);

                    cmd.Parameters.Add("@weight", item.Weight);
                    cmd.Parameters.Add("@height", item.Height);
                    cmd.Parameters.Add("@width", item.Weight);

                    cmd.Parameters.Add("@length", item.Length);
                    cmd.Parameters.Add("@size", item.Size);
                    cmd.Parameters.Add("@code", item.Code);

                    cmd.Parameters.AddWithValue("@brand", item.BrandID);


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

        static long InsertPrice(ProductItem item)
        {
            try
            {
                using (MySqlConnection connection = new MySqlConnection(ttaMariaDBConnectionString))
                {
                    connection.Open();
                    MySqlCommand cmd = new MySqlCommand("INSERT INTO prices (product_id, price, supplier_id, price_date, created) VALUES (@product_id,@price,@supplier_id, @price_date, NOW())", connection);
                    cmd.Parameters.Add("@product_id", item.ID);
                    cmd.Parameters.Add("@price", item.Price);
                    cmd.Parameters.Add("@supplier_id", item.SupplierID);
                    cmd.Parameters.Add("@price_date", item.PriceDate);
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

        public static int GetValueID(string name, string key, string value)
        {
            string sql = "SELECT id FROM options_lists WHERE name = @name AND `key` = @key" ;
            if(!string.IsNullOrEmpty(value))
            {
                sql += " AND `value` = @value";
            }
            using (MySqlConnection connection = new MySqlConnection(ttaMariaDBConnectionString))
            {
                connection.Open();

                MySqlCommand cmd = new MySqlCommand(sql, connection);
                cmd.Parameters.AddWithValue("@name", name);
                cmd.Parameters.AddWithValue("@key", key);
                if (!string.IsNullOrEmpty(value))
                {
                    cmd.Parameters.AddWithValue("@value", value);
                }
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

                    MySqlCommand cmd = new MySqlCommand("INSERT INTO options_lists (`name`, `key`, `value`) VALUES (@name,@key,@value)", connection);

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
            string sql = "SELECT id FROM products WHERE brand = @brand AND type = @type AND name LIKE @name";

            using (MySqlConnection connection = new MySqlConnection(ttaMariaDBConnectionString))
            {
                connection.Open();

                MySqlCommand cmd = new MySqlCommand(sql, connection);
                cmd.Parameters.AddWithValue("@name", item.Name);
                cmd.Parameters.AddWithValue("@brand", item.BrandID);
                cmd.Parameters.AddWithValue("@type", item.TypeID);
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        reader.Read();

                        return true;
                    }
                }
                connection.Close();
            }
            return false;
        }

        public static int FindSupplierID(string something)
        {
            string sql = "SELECT id FROM suppliers WHERE `name` LIKE @caigi OR `phone` LIKE @caigi OR `email` LIKE @caigi";
            using (MySqlConnection connection = new MySqlConnection(ttaMariaDBConnectionString))
            {
                connection.Open();

                MySqlCommand cmd = new MySqlCommand(sql, connection);
                cmd.Parameters.AddWithValue("@caigi", "'%"+something+"%'");

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

        public static long InsertSupplier(SupplierItem item)
        {
            string sql = "INSERT INTO suppliers (name, phone) VALUES (@name, @phone)";

            try
            {
                using (MySqlConnection connection = new MySqlConnection(ttaMariaDBConnectionString))
                {
                    connection.Open();

                    MySqlCommand cmd = new MySqlCommand(sql, connection);

                    cmd.Parameters.Add("@name", item.Name);
                    cmd.Parameters.Add("@phone", item.Phone);

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

        public static bool IsEnglishLetter(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }
    }

    public class ProductItem
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public double Price { get; set; }
        public DateTime PriceDate { get; set; }
        public string SKU { get; set; }
        public string Colour { get; set; }
        public double Weight { get; set; }//gram
        public double Size { get; set; }
        public int TypeID { get; set; }
        public string TypeName { get; set; }
        public int SupplierID { get; set; }
        public int SupplierName { get; set; }
        public int BrandID { get; set; }
        public string BrandName { get; set; }
        public string Code { get; set; }
        public string Details { get; set; }
        public double Height { get; set; }//cm
        public double Length { get; set; }//cm
        public double Width { get; set; }//cm
        public int SubTypeID { get; set; }
    }
    public class ColumnIndexItem
    {
        public int ProductName { get; set; }
        public int Code { get; set; }
        public int Type { get; set; }
        public int SubType { get; set; }
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
    public class SupplierItem
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Phone { get; set; }
        public string Address { get; set; }
        public string Email { get; set; }
        public string Website { get; set; }
    }
}