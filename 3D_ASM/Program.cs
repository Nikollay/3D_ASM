using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.IO;

namespace ASM_3D
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
            Board board;
            Point point;
            board = new Board();
            int board_outline_start=0, drilled_holes_start=0, placement_start=0, board_outline_end=0, drilled_holes_end=0, placement_end=0;

            List<string> board_outline;
            List<string> drilled_holes;
            List<string> placement;
            string[] strSplit;
            
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
            
            board.thickness = float.Parse(board_outline[0].Replace(".", ","));
            board_outline.RemoveAt(0);


            board.point = new List<Point>();
            for (int i = 0; i < board_outline.Count; i++)
            {
                point = new Point();
                strSplit = board_outline[i].Split((char)32);
                point.x = float.Parse(strSplit[1].Replace(".", ","))/1000;
                point.y = float.Parse(strSplit[2].Replace(".", ","))/1000;
                point.angle = float.Parse(strSplit[3].Replace(".", ","));
                board.point.Add(point);
            }

            //SolidWorks
            Console.WriteLine("Подключение к SldWorks.Application");
            var progId = "SldWorks.Application.27";
            var progType = System.Type.GetTypeFromProgID(progId);
            var swApp = System.Activator.CreateInstance(progType) as ISldWorks;
            swApp.Visible = true;
            Console.WriteLine("Успешное подключение к версии SldWorks.Application " + swApp.RevisionNumber());
            Console.WriteLine(DateTime.Now.ToString());
            Console.CursorSize = 100;
            ModelDoc2 swModel;
            AssemblyDoc swAssy;


            //Новая сборка платы
            double swSheetWidth=0, swSheetHeight=0;
            string boardName;
            int Errors=0, Warnings=0;
            swAssy = (AssemblyDoc)swApp.NewDocument("D:\\PDM\\EPDM_LIBRARY\\EPDM_SolidWorks\\EPDM_SWR_Templates\\Модуль_печатной_платы.asmdot", (int)swDwgPaperSizes_e.swDwgPaperA2size, swSheetWidth, swSheetHeight);
            swModel = (ModelDoc2)swAssy;
            //Сохранение
            boardName = filename.Remove(filename.Length-3) + "SLDASM";
            Console.WriteLine(boardName);
            swModel.Extension.SaveAs(boardName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews, null, ref Errors, ref Warnings);
            //**********

            //Доска
            Component2 board_body;
            PartDoc part;
            ModelDoc2 swCompModel;
            Feature swRefPlaneFeat, plane;
            swAssy.InsertNewVirtualPart(null, out board_body);
            board_body.Select4(false, null, false);
            swAssy.EditPart();
            swCompModel = (ModelDoc2)board_body.GetModelDoc2();
            part = (PartDoc)swCompModel;
            part.SetMaterialPropertyName2("-00", "гост материалы.sldmat", "Rogers 4003C");

            int j = 1;
            do
                {
                swRefPlaneFeat = (Feature)swCompModel.FeatureByPositionReverse(j);
                j = j + 1;
                }
            while (swRefPlaneFeat.Name != "Спереди");

            plane = (Feature)board_body.GetCorresponding(swRefPlaneFeat);
            plane.Select2(false, -1);

            swModel.SketchManager.InsertSketch(false);
            swModel.SketchManager.AddToDB = true;

            foreach (Point p in board.point)
            {
                //Console.WriteLine(p.x +(char)32+ p.y);
                Console.WriteLine(p.x);
                Console.WriteLine(p.y);
                swModel.SketchManager.CreatePoint(p.x, p.y, 0);
            }
            swModel.ClearSelection2(true);
            swAssy.HideComponent();
            swAssy.ShowComponent();
            swAssy.EditAssembly();



                //for (int i = 0; i < board_outline.Count; i++) { Console.WriteLine(board_outline[i]); }
                //for (int i = 0; i < drilled_holes.Count; i++) { Console.WriteLine(drilled_holes[i]); }
                //for (int i = 0; i < placement.Count; i++) { Console.WriteLine(placement[i]); }

            Console.ReadKey();
            
        }
    }
}
