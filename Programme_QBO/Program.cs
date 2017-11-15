using System;
using System.Net;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Programme_QBO
{
    class Program
    {
        private static Excel.Workbook Mybook = null;
        private static Excel.Application Myapp = null;
        private static Excel.Worksheet Mysheet = null;
        private static Process opned;
        private static Process Opned
        {
            get { return opned; }
            set { opned = value; }
        }
        private static bool connected = false;
        private static bool Connected
        {
            get { return connected; }
            set { connected = value; }
        }
        static void Main(string[] args)
        {
            Myapp = new Excel.Application();
            Myapp.Visible = false;
            int disconnectedCounter=0;
            Mybook = Myapp.Workbooks.Open(@"C:\Data\Random\Code QBO\Code QBO\Codes.xlsx");
            Mysheet = (Excel.Worksheet)Mybook.Sheets[1];
            int lastRow = Mysheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            BindingList<Code> Codelist = new BindingList<Code>();
            for (int index =1;index<=lastRow;index++)
            {
                Array MyValues = (Array)Mysheet.get_Range("A" + index.ToString(),"B"+index.ToString()).Cells.Value;
                Codelist.Add(new Code { Id = MyValues.GetValue(1, 1).ToString(), Path = MyValues.GetValue(1, 2).ToString() });
            }
            
            Console.WriteLine("Numéro du port?");
            
            string portnumber = Console.ReadLine();
            string port = ("COM"+ portnumber);
            SerialPort Port = new SerialPort(port);
            string lastcode="";
            bool isUrl = false;
            Port.Open();
            
            


            while (true){
                string message = Port.ReadLine();
                string decoded = message.Substring(0, 4);
                string pattern = @"\b"+decoded;
                string urlpattern = "\bhttp://";
                

                foreach (Code code in Codelist) { 

                    string id = code.Id;
                    Match match = Regex.Match(id, pattern);
                    Match verif = Regex.Match(lastcode, pattern);
                    Match close = Regex.Match("0000", pattern);
                    

                    if (match.Success && verif.Success==false)
                    {

                        lastcode = id;
                        try
                        {
                            Console.WriteLine("Opening : " + code.Path);

                            Opned =Process.Start(code.Path);
                            Uri uriResult;
                            bool result = Uri.TryCreate(code.Path, UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                            if (result)
                            {
                                isUrl = true;
                            }
                            else
                            {
                                isUrl = false;
                            }

                            Connected = true;
                            
                        }
                        catch (System.ComponentModel.Win32Exception)
                        {
                            Console.WriteLine("Wrong Path please check the Codes spreadsheet (remember : no spaces in the path allowed");
                        }
                    }
                    if (match.Success && verif.Success)
                    {
                        Console.WriteLine("already launched");
                    }
                    if (close.Success && Connected==true)
                    {
                        disconnectedCounter += 1;
                        if (disconnectedCounter >= 2)
                        {
                            Console.WriteLine("Closing last opened program");
                            if (isUrl == false)
                            {
                                Opned.Kill();
                            }
                            Connected = false;
                            lastcode = "";
                            disconnectedCounter = 0;
                        }
                        
                    }

                }
            }


            
        }
    }
}
