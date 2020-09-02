using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Threading;
using System.Windows.Forms;
using System.IO;

namespace _3D_ASM
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string filename;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.InitialDirectory = "D:\\1";
            fileDialog.Filter = "brd files (*.brd)|*.brd";
            fileDialog.RestoreDirectory = true;
            fileDialog.Title = "Открыть файл";
            fileDialog.ShowDialog();
            filename = fileDialog.FileName;
            Console.WriteLine(filename);

            string[] line = File.ReadAllLines(filename, System.Text.Encoding.GetEncoding(1251));
            Console.WriteLine(line.Length);

            //StreamReader f = new StreamReader(filename, System.Text.Encoding.GetEncoding(1251));
            //while (!f.EndOfStream)
            //{
            //    Console.WriteLine(f.ReadLine());
            //    //string s = f.ReadLine();
            //}
            Console.ReadKey();
            //f.Close();
        }
    }
}
