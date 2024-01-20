using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tekla.Structures;
using Tekla.Structures.ModelInternal;
using TSM = Tekla.Structures.Model;

namespace ReportGen
{

    internal class Program
    {
       public static List<string> modelslist = new List<string> { };

        static void Main(string[] args)
        {
            
            start_dummytekla();
            Models_list_handling();
            
        }
        
        private static void start_dummytekla()
        {
            Process p = new Process();
            p.StartInfo.FileName="C:\\Users\\DELL\\Desktop\\dummy tekla.lnk";
            p.StartInfo.UseShellExecute = true;
            p.Start();

        }
        public static void Models_list_handling()
        {
            //models_list_creation

            string[] lines = File.ReadAllLines("D:\\Engineering\\projects\\4D\\5D\\Workfile\\Models_List.txt");
            foreach (string line in lines)
            {
                if (line.Contains("REM"))
                {
                    
                }
                else
                {
                    modelslist.Add(line);
                    
                }
               
                
            }
            modelslist.ForEach(Console.WriteLine);

        //models_list_processing
            Repeat1:
            TSM.Model model = new TSM.Model();
            TSM.ModelHandler handler = new TSM.ModelHandler();
            try
            {
                model.GetConnectionStatus();
            }
            catch (Exception e) {
                Console.WriteLine("Establishing connection..");
                    }
            finally
            {
                foreach (string line in modelslist)
                {
                    handler.Open(line);
                Repeat2:
                    TSM.ModelInfo info = model.GetInfo();
                    if (info.ModelPath == line)
                    {
                        string modelname = info.ModelName;
                        handler.Save();
                        TSM.Operations.Operation.CreateReportFromAll("C:\\ProgramData\\Trimble\\Tekla Structures\\2020.0\\Environments\\Middle_East\\General\\template\\STEEL_MTO_AEK.pdf.rpt", "D:\\Engineering\\projects\\4D\\5D\\Workfile\\Reports\\" + "Quantities_" + modelname + ".pdf", "", "", "");
                    }
                    else
                    {
                        goto Repeat2;
                    }
                }
            }
               
                
            
            
            Console.Read();
        }
       
        private static void Send_Email()
        {
            

        }
    }
}
