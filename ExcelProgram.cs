using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace IRIS_Test
{
    public class GetExcelCoulmnNumberFromRowName
    {
        public static void Main(string[] args)
        {
            //Logger path
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Logger.txt");
            try
            {
                Console.WriteLine("Enter your input :" + "\n" + "a - If you want corresponding column number of column name");
                Console.WriteLine("b - If you want corresponding column name of column number");
                string? userInput = Console.ReadLine();

                switch (userInput)
                {
                    case "a":
                        Console.WriteLine(GetColumnNumberbyName());
                        break;

                    case "b":
                        Console.WriteLine(GetColumnNamebyNumber());
                        break;

                    default:
                        Console.WriteLine(" You did not type a or b");
                        Console.WriteLine();
                        Console.ReadLine();
                        break;
                }
            } 
            catch (Exception ex)
            {
               File.AppendAllText(path,"\n" +"Exception: " + "LoggerDate: " + DateTime.Now + " - " + ex.ToString() +"\n");
            }           

        }

        /// <summary>
        /// Method to get ColumnNumber in Excel from given ColumnName
        /// </summary>
        /// <returns></returns>
        public static string GetColumnNumberbyName()
        {
            Console.WriteLine("Enter Excel Column Name:");
            string? inputString;
            inputString = Console.ReadLine();
            string result = string.Empty;
            if (!string.IsNullOrEmpty(inputString))
            {
                int columnNumber = 0;
                for (int i = 0; i < inputString.Length; i++)
                {
                    columnNumber *= 26;
                    columnNumber += inputString.ToUpper()[i] - 'A' + 1;
                    result = columnNumber.ToString();
                }
                if (columnNumber < 0)
                    result = "Please enter valid Excel column name";
            }
            return result;

        }
        /// <summary>
        /// Method to get ColumnName in Excel from given ColumnNumber
        /// </summary>
        /// <returns></returns>
        public static string GetColumnNamebyNumber()
        {
            Console.WriteLine("Enter Excel Column Number:");
            string? inputColumnNumber;
            string result = string.Empty;
            inputColumnNumber = Console.ReadLine();
            if (!string.IsNullOrEmpty(inputColumnNumber))
            {

                if (int.TryParse(inputColumnNumber, out _))
                {
                    int columnNumber = Convert.ToInt32(inputColumnNumber);

                    while (columnNumber > 0)
                    {
                        int rem = columnNumber % 26;
                        if (rem == 0)
                        {
                            result += "Z";
                            columnNumber = (columnNumber / 26) - 1;
                        }
                        else
                        {
                            result += (char)((rem - 1) + 'A');
                            columnNumber = columnNumber / 26;
                        }
                    }
                    result = reverse(result);
                }
                else
                {
                    result = "Please input valid number";
                }
            }            
            return result;
        }

        static string reverse(string input)
        {
            char[] reversedString = input.ToCharArray();
            Array.Reverse(reversedString);
            return new string(reversedString);
        }
    }

}
