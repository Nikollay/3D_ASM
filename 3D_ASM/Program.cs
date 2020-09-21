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
        static void Main(string[] args)
        {
              
            Board board;
            string filename;
            //Console.WriteLine(Board.GetXML("d:\\Домашняя работа\\ПАКБ.468179.801_2.xml"));
            //Console.ReadKey();

            //SolidWorks
            Console.WriteLine("Подключение к SldWorks.Application");
            var progId = "SldWorks.Application.27";
            var progType = System.Type.GetTypeFromProgID(progId);
            Console.WriteLine(progType);
            Console.ReadKey();
            var swApp = System.Activator.CreateInstance(progType) as ISldWorks;
            swApp.Visible = true;
            Console.WriteLine("Успешное подключение к версии SldWorks.Application " + swApp.RevisionNumber());
            Console.WriteLine(DateTime.Now.ToString());
            Console.CursorSize = 100;

            filename = swApp.GetOpenFileName("Открыть файл", "", "xml Files (*.xml)|*.xml|", out _, out _, out _); //Board.GetFilename();
            if (String.IsNullOrWhiteSpace(filename)) { return; }
            board = Board.GetfromXML(filename);

            ModelDoc2 swModel;
            AssemblyDoc swAssy;
            ModelView activeModelView;
                       
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
                j ++;
            }
            while (swRefPlaneFeat.Name != "Спереди");

            plane = (Feature)board_body.GetCorresponding(swRefPlaneFeat);
            plane.Select2(false, -1);

            swModel.SketchManager.InsertSketch(false);
            swModel.SketchManager.AddToDB = true;
            
            //Эскизы
            foreach (Object skt in board.sketh)
            {
                if (skt.GetType().FullName== "ASM_3D.Line") { Line sk = (Line)skt;  swModel.SketchManager.CreateLine(sk.x1, sk.y1, 0, sk.x2, sk.y2, 0); }
                if (skt.GetType().FullName == "ASM_3D.Arc") { Arc sk = (Arc)skt; swModel.SketchManager.CreateArc(sk.xc, sk.yc, 0, sk.x1, sk.y1, 0, sk.x2, sk.y2, 0, sk.direction); }
            }
            swModel.FeatureManager.FeatureExtrusion3(true, false, false, 0, 0, board.thickness, board.thickness, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);
            swModel.ClearSelection2(true);

            plane.Select2(false, -1);
            swModel.SketchManager.InsertSketch(false);
            swModel.SketchManager.AddToDB = true;
            swModel.SketchManager.DisplayWhenAdded = false;

            foreach (Circle c in board.circles)
            {
                swModel.SketchManager.CreateCircleByRadius(c.xc, c.yc, 0, c.radius);
            }
            swModel.FeatureManager.FeatureCut3(true, false, true, 1, 0, board.thickness, board.thickness, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false);

            swAssy.HideComponent();
            swAssy.ShowComponent();
            swModel.ClearSelection2(true);
            swAssy.EditAssembly();

            string path;
            path = "D:\\PDM\\Прочие изделия\\ЭРИ";
            List<string> allFoundFiles = new List<string>(Directory.GetFiles(path, "*.SLD*", SearchOption.AllDirectories));
            Dictionary<string, string> empty = new Dictionary<string, string>();

            string sample;
            foreach (Component comp in board.components)
            {
                sample = comp.title;
                comp.fileName = allFoundFiles.Find(item => item.Contains(sample));
                if (String.IsNullOrWhiteSpace(comp.fileName))
                { 
                    comp.fileName = "D:\\PDM\\Прочие изделия\\ЭРИ\\Zero.SLDPRT"; 
                   if( !empty.ContainsKey(sample)) { empty.Add(sample, sample); }
                }
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
            swAssy.AddComponents3(names, transforms, coordSys);

            //Фиксация
            swModel.Extension.SelectAll();
            swAssy.FixComponent();
            swModel.ClearSelection2(true);

            activeModelView = (ModelView)swModel.ActiveView;
            activeModelView.DisplayMode = (int)swViewDisplayMode_e.swViewDisplayMode_ShadedWithEdges;
            //****************************

            UserProgressBar pb;
            swApp.GetUserProgressBar(out pb);
            
            //Заполнение поз. обозначений
            List<Component2> compsColl = new List<Component2>(); //Коллекция из компонентов сборки платы
            Feature swFeat;
            Component2 compTemp;
            pb.Start(0, board.components.Count, "Поиск");
            int itm = 0;
            swFeat = (Feature)swModel.FirstFeature();
            while (swFeat != null)
            {
                pb.UpdateProgress(itm);
                //pb.UpdateTitle(itm);
                if (swFeat.GetTypeName().Equals("Reference")) //Заполняем коллекцию изделиями
                {
                    compTemp = (Component2)swFeat.GetSpecificFeature2();
                    compsColl.Add(compTemp);
                }
                swFeat = (Feature)swFeat.GetNextFeature();
                itm++;
            }
            pb.End();

            compsColl[0].Name2 = "Плата"; //Пререименовываем деталь      
            if (compsColl.Count - 1 == board.components.Count) //Проверка чтобы не сбились поз. обозначения, если появятся значит все правильно иначе они не нужны
            {
                for (int i = 0; i < board.components.Count; i++)
                    compsColl[i + 1].ComponentReference = board.components[i].physicalDesignator; //Заполняем поз. обозначениями
            }

            string estr="";
            if (empty.Count!=0)
            {
                foreach (KeyValuePair<string,string> str in empty) { estr = estr + str.Value + System.Environment.NewLine; }
                MessageBox.Show(estr, "Не найдены");
                //swApp.SendMsgToUser2("Не найдены" + estr, 2, 2);
            }
            //**************
            //swApp.ExitApp();
            //swApp = null;


        }
   
    }
}
