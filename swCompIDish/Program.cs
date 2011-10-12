using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using SwDocumentMgr;
using DSOFile;

namespace swCompIDish
{
    class Program
    {
        static void Main(string[] args)
        {
            //string[] args = { @"C:\PracticeFiles3\Version_1008.SLDDRW" };
            //Takes Care of input checking and input parsing
            string docPath;
            bool quietMode;
            switch (args.Length)
            {
                case 1:
                    quietMode = false;
                    if (args[0].Contains("*") || args[0].Contains("?"))
                    {
                        inputError(quietMode);
                        return;
                    }
                    docPath = Path.GetFullPath(args[0]);
                    break;
                case 2:
                    if (args[0] != "/q")
                    {
                        quietMode = false;
                        inputError(quietMode);
                        return;
                    }
                    quietMode = true;
                    if (args[1].Contains("*") || args[1].Contains("?"))
                    {
                        inputError(quietMode);
                        return;
                    }
                    docPath = Path.GetFullPath(args[1]);
                    break;
                default:
                    quietMode = false;
                    inputError(quietMode);
                    return;
            }
            
            //Get Document Type
            SwDmDocumentType docType = setDocType(docPath);
            if (docType == SwDmDocumentType.swDmDocumentUnknown)
            {
                inputError(quietMode);
                return;
            }
            try
            {
                OleDocumentProperties dsoProperties = new OleDocumentProperties();
                dsoProperties.Open(docPath, false);
                Console.WriteLine("\"" + docPath + "\"\t\t\"" + dsoProperties.SummaryProperties.DateCreated + "\"");
                dsoProperties.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("\"" + docPath + "\"\t\"" + "File is internally damaged, .NET error occurred, or swCompIDish.exe has a Bug. " + "Error Message: " + e.Message.Replace(Environment.NewLine, " ") + ". Stack Trace: " + e.StackTrace.Replace(Environment.NewLine, " ") + "\"");
                inputError(quietMode);
            }
        }




        static void inputError(bool quietMode)
        {
            if (quietMode)
                return;

            Console.WriteLine(@"
Syntax 
    [option] [ParentFilePath]

Output
    ""FilePath""        ""dsoFileCreationDate"" ""CLSID""

Output if Error Occurs
    ""ParentFilePath""  ""ErrorMessageAndStackTrace""


Only one Filename is accepted.  No wildcars allowed. If the path has spaces 
use quotes around it.  Note that the file must have one of the following 
file extensions: .sldasm, .slddrw, .sldprt, .asm, .drw, or .prt.  The output
is tab delimited.  This makes it easy to redirect the output to a text file
that can be opened as spreadsheet.

Options
    /q      Quiet mode.  Suppresses the current message.  It does
            not suppress the one line error messages related to problems
            opening SolidWorks Files.  Quiet mode is useful for batch files
            when you are directing the output to a file.  The main error 
            message is suppressed but you are still informed about problems 
            opening files.

Version 2011-Oct-12 12:06
Written and Maintained by Jason Nicholson
http://github.com/jasonnicholson/swCompIDish");
        }






        static SwDmDocumentType setDocType(string docPath)
        {
            string fileExtension = System.IO.Path.GetExtension(docPath);

            //Notice no break statement is needed because I used return to get out of the switch statement.
            switch (fileExtension.ToUpper())
            {
                case ".SLDASM":
                case ".ASM":
                    return SwDmDocumentType.swDmDocumentAssembly;
                case ".SLDDRW":
                case ".DRW":
                    return SwDmDocumentType.swDmDocumentDrawing;
                case ".SLDPRT":
                case ".PRT":
                    return SwDmDocumentType.swDmDocumentPart;
                default:
                    return SwDmDocumentType.swDmDocumentUnknown;
            }

        }




    }
}
