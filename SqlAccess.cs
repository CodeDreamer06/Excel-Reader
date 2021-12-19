using System;
using System.IO;
using System.Data;
using System.Data.SQLite;
using OfficeOpenXml;

namespace DotNet_Excel_Reader
{
  class SqlAccess
  {
    public string path = "./data.db";
    protected static void Execute(string query) {
      // try {
        using(SQLiteConnection con = new SQLiteConnection("Data Source=./data.db;Version=3;")){
          con.Open();
          using var cmd = new SQLiteCommand(query, con);
          cmd.ExecuteNonQuery();
        }
      // }

      // catch {
      //   Console.WriteLine("An unknown error occured.");
      // }
    }

    public static void CreateTable() {
      if(File.Exists(@"./data.db")) {
          File.Delete(@"./data.db");
      }
         try {
           using(SQLiteConnection con = new SQLiteConnection("Data Source=./data.db;Version=3;")){
             con.Open();
             var cmd = new SQLiteCommand(@"SELECT name FROM sqlite_master WHERE type='table' AND name='passwords'", con);
              if(cmd.ExecuteScalar() == null)
                Execute(@"CREATE TABLE passwords(id INTEGER PRIMARY KEY AUTOINCREMENT, password VARCHAR(20) NOT NULL, length INT, num_chars INT, num_digits INT)");
           }
         }
         catch {
           Console.WriteLine("Unable to check if table exists.");
         }
     }

     public static void GetPasswords(string query = @"select * from passwords") {
      using(SQLiteConnection con = new SQLiteConnection("Data Source=./data.db;Version=3;")){
        con.Open();
        using var cmd = new SQLiteCommand(query, con);
        SQLiteDataReader reader = cmd.ExecuteReader();
        while (reader.Read())
          Console.WriteLine($"{reader["id"]} {reader["password"]} {reader["length"]} {reader["num_chars"]} {reader["num_digits"]}");
      }
    }

    protected static void AddPassword(string password, string length, string numChars, string numDigits) {
      int convertedLength;
      int.TryParse(length, out convertedLength);
      int convertedNumChars;
      int.TryParse(numChars, out convertedNumChars);
      int convertedNumDigits;
      int.TryParse(numDigits, out convertedNumDigits);
      Execute($"INSERT INTO passwords(password, length, num_chars, num_digits) VALUES('{password}', {convertedLength}, {convertedNumChars}, {convertedNumDigits});");
   }

   public static void FetchData() {
     Console.WriteLine("Loading the file...");
     FileInfo existingFile = new FileInfo("common_passwords.xlsx");
         using (ExcelPackage package = new ExcelPackage(existingFile))
         {
             ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
             Console.WriteLine("Adding logs into the database...");
             for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
             {
                 AddPassword(Convert.ToString(worksheet.Cells[row, 1].Value), Convert.ToString(worksheet.Cells[row, 2].Value), Convert.ToString(worksheet.Cells[row, 3].Value), Convert.ToString(worksheet.Cells[row, 4].Value));
             }
         }
   }
  }
}
