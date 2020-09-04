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
            Circle circle;
            Component component;
            board = new Board();
            int board_outline_start = 0, drilled_holes_start = 0, placement_start = 0, board_outline_end = 0, drilled_holes_end = 0, placement_end = 0;

            List<string> board_outline;
            List<string> drilled_holes;
            List<string> placement;
            string[] strSplit;

            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].ToUpper().Contains(".BOARD_OUTLINE UNOWNED")) { board_outline_start = i; }
                if (list[i].ToUpper().Contains(".END_BOARD_OUTLINE")) { board_outline_end = i; }
                if (list[i].ToUpper().Contains(".DRILLED_HOLES")) { drilled_holes_start = i; }
                if (list[i].ToUpper().Contains(".END_DRILLED_HOLES")) { drilled_holes_end = i; }
                if (list[i].ToUpper().Contains(".PLACEMENT")) { placement_start = i; }
                if (list[i].ToUpper().Contains(".END_PLACEMENT")) { placement_end = i; }
            }

            board_outline = list.GetRange(board_outline_start + 1, board_outline_end - board_outline_start - 1);
            drilled_holes = list.GetRange(drilled_holes_start + 1, drilled_holes_end - drilled_holes_start - 1);
            placement = list.GetRange(placement_start + 1, placement_end - placement_start - 1);
            //for (int i = 0; i < placement.Count; i++) { Console.WriteLine(placement[i]); }
            board.thickness = double.Parse(board_outline[0].Replace(".", ",")) / 1000;
            board_outline.RemoveAt(0);


            board.point = new List<Point>();
            for (int i = 0; i < board_outline.Count; i++)
            {
                point = new Point();
                strSplit = board_outline[i].Split((char)32);
                point.x = float.Parse(strSplit[1].Replace(".", ",")) / 1000;
                point.y = float.Parse(strSplit[2].Replace(".", ",")) / 1000;
                point.angle = float.Parse(strSplit[3].Replace(".", ","));
                board.point.Add(point);
            }

            board.circles = new List<Circle>();
            for (int i = 0; i < drilled_holes.Count; i++)
            {
                circle = new Circle();
                strSplit = drilled_holes[i].Split((char)32);
                circle.xc = float.Parse(strSplit[1].Replace(".", ",")) / 1000;
                circle.yc = float.Parse(strSplit[2].Replace(".", ",")) / 1000;
                circle.radius = float.Parse(strSplit[0].Replace(".", ",")) / 2000;
                if (!strSplit[5].Contains("VIA")) { board.circles.Add(circle); }
            }

            board.components = new List<Component>();
            for (int i = 0; i < placement.Count; i++)
            {
                if (i % 2 == 0)
                {
                    component = new Component();
                    strSplit = placement[i].Split((char)32);
                    component.footprint = strSplit[0];
                    component.physicalDesignator = strSplit[strSplit.Length - 1];
                    component.part_Number = placement[i].Replace(component.footprint, "");
                    component.part_Number = component.part_Number.Replace(component.physicalDesignator, "");
                    component.part_Number = component.part_Number.Trim().Trim('\"');

                    strSplit = placement[i + 1].Split((char)32);
                    component.x = float.Parse(strSplit[0].Replace(".", ",")) / 1000;
                    component.y = float.Parse(strSplit[1].Replace(".", ",")) / 1000;
                    component.z = board.thickness;
                    component.standOff = float.Parse(strSplit[2].Replace(".", ",")) / 1000;
                    component.rotation = float.Parse(strSplit[3].Replace(".", ","));
                    switch (strSplit[4])
                    {
                        case "TOP":
                            component.layer = 1;
                            break;
                        case "BOTTOM":
                            component.layer = 0;
                            break;
                    }
                    board.components.Add(component);
                    //Console.WriteLine(component.part_Number);
                    //Console.WriteLine(component.x);
                    //Console.WriteLine(component.y);
                    //Console.WriteLine(component.rotation);
                }
            }
            //Console.ReadKey();

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
            double swSheetWidth = 0, swSheetHeight = 0;
            string boardName;
            int Errors = 0, Warnings = 0;
            swAssy = (AssemblyDoc)swApp.NewDocument("D:\\PDM\\EPDM_LIBRARY\\EPDM_SolidWorks\\EPDM_SWR_Templates\\Модуль_печатной_платы.asmdot", (int)swDwgPaperSizes_e.swDwgPaperA2size, swSheetWidth, swSheetHeight);
            swModel = (ModelDoc2)swAssy;
            //Сохранение
            boardName = filename.Remove(filename.Length - 3) + "SLDASM";
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


            for (int i = 1; i < board.point.Count; i++)
            {
                swModel.SketchManager.CreateLine(board.point[i - 1].x, board.point[i - 1].y, 0, board.point[i].x, board.point[i].y, 0);
            }
            swModel.FeatureManager.FeatureExtrusion3(true, false, false, 0, 0, board.thickness, board.thickness, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);

            foreach (Circle c in board.circles)
            {
                swModel.SketchManager.CreateCircleByRadius(c.xc, c.yc, 0, c.radius);
            }
            swModel.FeatureManager.FeatureCut3(true, false, true, 1, 0, board.thickness, board.thickness, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false);
            foreach (Point p in board.point)
            {
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

            // Console.ReadKey();
            List<string> allFoundFiles = new List<string>(Directory.GetFiles("D:\\PDM\\Прочие изделия\\ЭРИ", "*.*", SearchOption.AllDirectories));

            foreach (Component comp in board.components)
            {
                comp.fileName = allFoundFiles.Find(item => item.Contains(comp.part_Number));
                if (String.IsNullOrWhiteSpace(comp.fileName)) { comp.fileName = "D:\\PDM\\Прочие изделия\\ЭРИ\\Zero.SLDPRT"; }
            }
          
            double[] transforms, dMatrix;
            string[] coordSys, names;
            double alfa, beta, gamma, x, y, z;
            names = new string[board.components.Count];
            coordSys = new string[board.components.Count];
            dMatrix = new double[16];
            transforms = new double[board.components.Count*16];

            for (int i = 0; i < board.components.Count; i++)
            {
                names[i] = board.components[i].fileName;
            }
            int n = 0;
            foreach (Component comp in board.components)
            {
                alfa = 0;
                x = comp.x;
                y = comp.y;
                if (comp.layer == 1) //Если Top
                {
                    z = (comp.z + comp.standOff);
                    beta = -Math.PI / 2;
                }
                else             //Иначе Bottom
                {
                    z = (comp.z - comp.standOff) / 1000;
                    beta = Math.PI / 2;
                }
                gamma = -(comp.rotation / 180) * Math.PI;

                dMatrix[0] = Math.Cos(alfa) * Math.Cos(gamma) - Math.Sin(alfa) * Math.Cos(beta) * Math.Sin(gamma);
                dMatrix[1] = -Math.Cos(alfa) * Math.Sin(gamma) - Math.Sin(alfa) * Math.Cos(beta) * Math.Cos(gamma);
                dMatrix[2] = Math.Sin(alfa) * Math.Sin(beta); //1 строка матрицы вращения
                dMatrix[3] = Math.Sin(alfa) * Math.Cos(gamma) + Math.Cos(alfa) * Math.Cos(beta) * Math.Sin(gamma);
                dMatrix[4] = -Math.Sin(alfa) * Math.Sin(gamma) + Math.Cos(alfa) * Math.Cos(beta) * Math.Cos(gamma);
                dMatrix[5] = -Math.Cos(alfa) * Math.Sin(beta); //2 строка матрицы вращения
                dMatrix[6] = Math.Sin(beta) * Math.Sin(gamma);
                dMatrix[7] = Math.Sin(beta) * Math.Cos(gamma);
                dMatrix[8] = Math.Cos(beta); //3 строка матрицы вращения
                dMatrix[9] = x; dMatrix[10] = y; dMatrix[11] = z; //Координаты
                dMatrix[12] = 1; //Масштаб
                dMatrix[13] = 0; dMatrix[14] = 0; dMatrix[15] = 0; //Ничего

                for (int k = 0; k < dMatrix.Length; k++) { transforms[n * 16 + k] = dMatrix[k]; }
                n++;
            }
            
            //Вставка
            swAssy.AddComponents3((names), (transforms), (coordSys));

            //Фиксация
            swModel.Extension.SelectAll();
            swAssy.FixComponent();
            swModel.ClearSelection2(true);
            //****************************
            //Заполнение поз. обозначений
            List<Component2> compsColl = new List<Component2>(); //Коллекция из компонентов сборки платы
            object[] tree = (object[])swModel.FeatureManager.GetFeatures(true); //Получаем дерево
            Feature featureTemp;
            Component2 compTemp;

            for (int i = 0; i < tree.Length; i++)
            {
                featureTemp = (Feature)tree[i];
                if (featureTemp.GetTypeName() == "Reference") //Заполняем коллекцию изделиями
                {
                    compTemp = (Component2)featureTemp.GetSpecificFeature2();
                    compsColl.Add(compTemp);
                }
            }

            compsColl[0].Name2 = "Плата"; //Пререименовываем деталь

            if (compsColl.Count == board.components.Count) //Проверка чтобы не сбились поз. обозначения, если появятся значит все правильно иначе они не нужны
            {
                for (int i = 0; i < board.components.Count; i++)
                    compsColl[i].ComponentReference = board.components[i].physicalDesignator; //Заполняем поз. обозначениями
            }      
            //**************
        }
    }
}
