using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;

namespace FindAccessViolation
{
    class Program
    {
        static bool notInFile(Excel._Worksheet ws, string form, string cia)
        {
            int i = 0;
            string valor;
            do
            {
                i++;
                if (ws.Cells[i, "B"].Text == form && ws.Cells[i, "C"].Text == cia)
                    return false;
                valor = ws.Cells[i, "A"].Text;
            } while (valor != "");
            return true;
        }
        static void Main(string[] args)
        {
            
            Console.WriteLine("Digite o caminho do Excel:");
            string path = Console.ReadLine().Trim();

            var excelApp = new Excel.Application();
            excelApp.Visible = true;
           
            var excelBook = excelApp.Workbooks.Open(path, Editable: true);
            Excel._Worksheet ws = excelApp.ActiveSheet; // ws recebe a planilha

            int nextLine;
            for (nextLine = 1; (ws.Cells[nextLine, "A"].Text) != ""; nextLine++);

            //int ID = Convert.ToInt32(ws.Cells[nextLine, "A"].Text);

            String StrID = ws.Cells[nextLine-1, "A"].Text;
            int ID = Convert.ToInt32(StrID);

            Console.WriteLine("Digite o diretório para pesquisar o erro: ");
            String sourceDirectory = Console.ReadLine();
            Console.WriteLine("Digite a mensagem de erro procurada (Case Sensitive): ");
            String message = Console.ReadLine();
            var txtFiles = Directory.EnumerateFiles(sourceDirectory, "*.txt");
            Console.WriteLine("\nResultados:\n");
            bool hasError;
            string form = "", cia = "";
            double total = txtFiles.Count(), progress = 0;
            foreach (string currentFile in txtFiles)
            {
                progress += 1;
                hasError = false;
                string[] Linhas = File.ReadAllLines(currentFile);
                foreach (string line in Linhas)
                {
                    if (line.Contains(message))
                        hasError = true;
                    if (hasError && line.Contains("Formu"))
                        form = line.Substring(18);
                    else if (hasError && line.Contains("Empresa.........:"))
                    {
                        cia = line.Substring(18);
                        if (notInFile(ws, form, cia))
                        {
                            ws.Cells[nextLine, "A"] = ID+=1;
                            ws.Cells[nextLine, "B"] = form;
                            ws.Cells[nextLine, "C"] = cia;
                            ws.Cells[nextLine, "D"] = Path.GetFullPath(currentFile);
                            nextLine += 1;
                        }
                    }                    
                    string porcentagem = ((int)((progress / total) * 100)).ToString();               
                    Console.Write("\rVerificando arquivo: {0} de: {1} ({2}% Completo)  ", new string[] { progress.ToString(), total.ToString(), porcentagem });
                }
            }
            excelBook.Save();
            excelBook.Close(true);
            excelApp.Quit();
            Console.WriteLine("\n\n*-- Fim de Execução --*");
            Console.ReadLine();
        }

    }
}
