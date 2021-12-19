using System;

namespace DotNet_Excel_Reader
{
    class Program
    {
      static void Main(string[] args)
        {
            SqlAccess.CreateTable();
            SqlAccess.FetchData();
            SqlAccess.GetPasswords();
        }
    }
}
