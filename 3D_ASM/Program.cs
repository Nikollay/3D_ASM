using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Linq;
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
            List<string> list = new List<string>(line);
            Console.WriteLine(line.Length);
            Board board;
            board = new Board();
            int board_outline_start=0, drilled_holes_start=0, placement_start=0, board_outline_end=0, drilled_holes_end=0, placement_end=0;

            List<string> board_outline;
            List<string> drilled_holes;
            List<string> placement;
            string[] strSplit;
            Console.WriteLine(list.Count);

            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].ToUpper().Contains(".BOARD_OUTLINE UNOWNED")) { board_outline_start=i; }
                if (list[i].ToUpper().Contains(".END_BOARD_OUTLINE")) { board_outline_end = i; }
                if (list[i].ToUpper().Contains(".DRILLED_HOLES")) { drilled_holes_start = i; }
                if (list[i].ToUpper().Contains(".END_DRILLED_HOLES")) { drilled_holes_end = i; }
                if (list[i].ToUpper().Contains(".PLACEMENT")) { placement_start = i; }
                if (list[i].ToUpper().Contains(".END_PLACEMENT")) { placement_end = i; }
            }

            board_outline = list.GetRange(board_outline_start + 1, board_outline_end-board_outline_start-1);
            drilled_holes= list.GetRange(drilled_holes_start + 1, drilled_holes_end - drilled_holes_start-1);
            placement= list.GetRange(placement_start + 1, placement_end - placement_start-1);

            board.thickness = double.Parse(board_outline[0]);

            for (int i = 1; i < board_outline.Count; i++)
            {
                strSplit = board_outline[i].Split((char)32);
                if (strSplit[3] == "0.0000")
                {
                    Line l = new Line();
                    l.x1 = double.Parse(strSplit[1]);
                    l.x2 = double.Parse(strSplit[2]);
                    board.lines.Add(l);
                }
                else { };
            }    



                //for (int i = 0; i < board_outline.Count; i++) { Console.WriteLine(board_outline[i]); }
                //for (int i = 0; i < drilled_holes.Count; i++) { Console.WriteLine(drilled_holes[i]); }
                //for (int i = 0; i < placement.Count; i++) { Console.WriteLine(placement[i]); }

                Console.ReadKey();
            
        }
    }
}
