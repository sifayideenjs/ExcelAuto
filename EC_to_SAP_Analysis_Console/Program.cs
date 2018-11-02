using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EC_to_SAP_Analysis_Console
{
    class Program
    {
        static void Main(string[] arguments)
        {
            if (arguments == null)
            {
                AddLogDetails("Command line arguments are null");
            }
            else if (arguments.Length == 0)
            {
                try
                {
                    DisplayHeader();
#if DEBUG
                    string iCIMSFilePath = @"C:\Users\h149041\Downloads\Analytics\Analytics\nwp\iCIMS.xlsx";
#else
                    System.Console.Write("Enter iCIMS file path:>");
                    string iCIMSFilePath = System.Console.ReadLine();
#endif
                    iCIMSFilePath = ValidatePath(iCIMSFilePath);

                    //System.Console.WriteLine(Environment.NewLine);

                    Console.WriteLine(string.Empty);
#if DEBUG
                    string onbFilePath = @"C:\Users\h149041\Downloads\Analytics\Analytics\nwp\ONB.xlsx";
#else
                    System.Console.Write("Enter ONB file path:>");
                    string onbFilePath = System.Console.ReadLine();
#endif
                    onbFilePath = ValidatePath(onbFilePath);

                    //System.Console.WriteLine(Environment.NewLine);

                    Console.WriteLine(string.Empty);
#if DEBUG
                    string sapHireFilePath = @"C:\Users\h149041\Downloads\Analytics\Analytics\nwp\SAP Hire Completion Date Data.xlsx";
#else
                    System.Console.Write("Enter SAP Hire Completion Date Data file path:>");
                    string sapHireFilePath = System.Console.ReadLine();
#endif
                    sapHireFilePath = ValidatePath(sapHireFilePath);

                    //System.Console.WriteLine(Environment.NewLine);

                    Console.WriteLine(string.Empty);
#if DEBUG
                    string sapNameValidationFilePath = @"C:\Users\h149041\Downloads\Analytics\Analytics\nwp\SAP Name Validation.xlsx";
#else
                    System.Console.Write("Enter SAP Name Validation file path:>");
                    string sapNameValidationFilePath = System.Console.ReadLine();
#endif
                    sapNameValidationFilePath = ValidatePath(sapNameValidationFilePath);

                    //System.Console.WriteLine(Environment.NewLine);

                    Console.WriteLine(string.Empty);
#if DEBUG
                    string aliasCreationFilePath = @"C:\Users\h149041\Downloads\Analytics\Analytics\nwp\Alias Creation.xlsx";
#else
                    System.Console.Write("Enter Alias Creation file path:>");
                    string aliasCreationFilePath = System.Console.ReadLine();
#endif
                    aliasCreationFilePath = ValidatePath(aliasCreationFilePath);

                    //System.Console.WriteLine(Environment.NewLine);

                    Console.WriteLine(string.Empty);
#if DEBUG
                    string outputDirectory = @"C:\Users\h149041\Downloads\Analytics\Analytics\nwp\OUT";
#else
                    System.Console.Write("Enter Output directory:>");
                    string outputDirectory = System.Console.ReadLine();
#endif
                    outputDirectory = ValidatePath(outputDirectory);

                    System.Console.WriteLine(Environment.NewLine);

                    Console.WriteLine(string.Empty);

                    Converter converter = new Converter(iCIMSFilePath, onbFilePath, sapHireFilePath, sapNameValidationFilePath, aliasCreationFilePath);
                    bool result = converter.Convert(outputDirectory);
                    Console.WriteLine(string.Empty);
                    AddLogDetails(result ? "Sucessfully Converted" : "Failed to Convert");

                    Console.ReadLine();

                    //Console.WriteLine(string.Empty);
                }
                catch (Exception e)
                {
                    AddLogDetails(e.Message);
                }
            }
            else if (arguments.Length == 1 && HelpRequired(arguments[0]))
            {
                DisplayHelp();
            }
        }

        private static void AddLogDetails(string message)
        {
            System.Console.WriteLine("{0}", message);
        }

        private static void DisplayHeader()
        {
            System.Console.WriteLine(string.Format("{0} | version 1.0 | Hanief Abdullah | 2018", "EC to SAP Analysis"));
            System.Console.WriteLine(Environment.NewLine);
        }

        private static string ValidatePath(string filePath)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                if (filePath.First() == '"')
                {
                    filePath = filePath.TrimStart('"');
                }

                if (filePath.Last() == '"')
                {
                    filePath = filePath.TrimEnd('"');
                }
            }
            return filePath;
        }

        private static bool HelpRequired(string param)
        {
            param = param.ToLower();
            return param == "-help" || param == "--help" || param == "/help" || param == "help" || param == "/?";
        }

        private static void DisplayHelp()
        {
            //DisplayHeader();
            System.Console.WriteLine("Options:");
            System.Console.WriteLine(string.Format("Usage: -{0}", "help"));
        }
    }
}
