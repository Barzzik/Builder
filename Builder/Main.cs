using System;
using Autodesk.AutoCAD.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using ACADApp = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;

namespace Builder
{
    public class Main
    {
        [CommandMethod("LineBuilder")]
        public void LineBuilder()
        {
            //создаем WinForm для получения адреса файла Excel
            Form1 form = new Form1();
            string pathFile = form.GetPath();
            if (pathFile == "") return;

            //создаем экземпляр приложения Excel
            Excel.Application excelApp = new Excel.Application();
            //объявляем переменные с книгой и листом файла Excel
            Excel.Workbook excelBook;
            Excel.Worksheet excelSheet;
            //объявляем переменную с границами данных на данном листе Excel
            Excel.Range excelRange;

            int rowsCount;
            int colsCount;

            //инициализируем переменные для доступа к книге (файлу excel), первому листу книги, определяем границы данных на листе
            excelBook = excelApp.Workbooks.Open(pathFile);
            excelSheet = excelBook.Sheets[1];
            excelRange = excelSheet.UsedRange;


            //считаем количество строк и столбцов с данными
            rowsCount = excelRange.Rows.Count;
            colsCount = excelRange.Columns.Count;

            //получаем ссылку на документ акад и его БД
            Document acadDoc = ACADApp.Application.DocumentManager.MdiActiveDocument;
            Database acadDb = acadDoc.Database;
            //создаем поле документа "Editor" для вывода сообщений в окно консоли AutoCAD
            var editor = ACADApp.Application.DocumentManager.MdiActiveDocument.Editor;

            //заправшиваем начальную координату (точка вхождения первого блока)
            Point3d ptStart = UserTalker.GetPointsFromUser().Value;
            try
            {
                //начинаем транзакцию
                using (Transaction transaction = acadDb.TransactionManager.StartTransaction())
                {

                    BlockTable blockTable = (BlockTable)transaction.GetObject(acadDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                    //создаем координаты для построения
                    double xInsert = ptStart.X;
                    double xPrevous = ptStart.X;
                    double yStart = ptStart.Y;

                    //открываем таблицу блоков на чтение
                    blockTable = (BlockTable)transaction.GetObject(acadDb.BlockTableId, OpenMode.ForRead);

                    //открываем пространство модели на запись
                    BlockTableRecord modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    Builder builder = new Builder(acadDb, transaction, modelSpace);

                    //перебираем все блоки из файла Excel
                    for (int i = 2; i <= rowsCount; i++)
                    {
                        //добавляем в модель блоки
                        builder.AddBlock(blockTable[excelRange.Cells[i, 1].Value2.ToString()],ptStart, xInsert);

                        //добавляем подпись под каждый блок
                        builder.AddText(ptStart, xInsert, excelRange.Cells[i, 1].Value2.ToString());

                        //создаем размеры для каждого блока после первого
                        if (i > 2)
                        {
                            builder.AddDimension(xInsert, xPrevous, yStart);
                            xPrevous = xInsert;
                        }
                        if (double.TryParse(excelRange.Cells[i, 2].Value2.ToString(), out double myD))
                        {
                            xInsert += myD;
                        }
                        //editor.WriteMessage(xStart.ToString() + "  " + xPrevous.ToString()+"\n ");

                    }
                    // фиксируем транзакцию
                    transaction.Commit();
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception)
            {
                editor.WriteMessage("\nОшибка загружаемых данных. Проверьте названия блоков в стобце 1, наличие блоков на чертеже");
                return;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                editor.WriteMessage("\nОшибка загружаемых данных. Проверьте файл на отсутствие пустых ячеек в столбцах 1 и 2");
                return;
            }
            finally
            {
                //освобождаем память, закрывая доступ
                excelBook.Close(false, Type.Missing, Type.Missing);
                excelApp.Quit();
            }
            //выводим в консоль количество строк и столбцов
            editor.WriteMessage("\nЗадача выполнена");
        }
    }

    //класс для создания диалогового окна
    public class Form1 : Form
    {
        public string GetPath()
        {
            var filePath = "";
            //создаем диалоговое окно для выбора файла
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //задаем начальную директорию, фильтр файлов
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                //если в диалоговом окне происходит нажатие кнопки подтверждения...
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //получаем полный адрес выбранного файла в переменную
                    filePath = openFileDialog.FileName;
                }
            }
            return filePath;
        }
    }

    //класс для общения с пользователем
    public class UserTalker
    {
        public static PromptPointResult GetPointsFromUser()
        {
            //получаем базу данных чертежа
            ACADApp.Document acDoc = ACADApp.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            PromptPointResult startPoint;
            PromptPointOptions pPtOpts = new PromptPointOptions("");

            //заправшиваем начальную точку
            pPtOpts.Message = "\nУкажите начальную точку: ";
            startPoint = acDoc.Editor.GetPoint(pPtOpts);
            return startPoint;
            ////выход, если пользователь отменил команду
            //if (startPoint.Status == PromptStatus.Cancel) return startPoint;
        }
    }

    //класс для добавления сущностей в моделе
    public class Builder
    {
        Database db;
        Transaction tr;
        BlockTableRecord ms;
        public Builder(Database dataBase, Transaction transaction, BlockTableRecord modelSpace)
        {
            db = dataBase;
            tr = transaction;
            ms = modelSpace;
        }

        public void AddDimension(double xInsert, double xPrevous, double yStart)
        {
            //создаем новый объект размера
            using (var dim = new RotatedDimension(0.0, new Point3d(xPrevous, yStart, 0.0), new Point3d(xInsert, yStart, 0.0), new Point3d(0.0, yStart + 5.0, 0.0), string.Empty, db.Dimstyle))
            {
                dim.Annotative = AnnotativeStates.True;
                //добавляем созданный объект в пространство модели и транзакцию
                ms.AppendEntity(dim);
                tr.AddNewlyCreatedDBObject(dim, true);
            }
        }

        public void AddBlock(ObjectId blockId, Point3d ptStart, double xInsert)
        {
            //используем ID блока с именем из таблицы
            ObjectId btrId = blockId;
            //создаем новое вхождение блока, используя ранее сохраненный ID определения блока
            using (BlockReference br = new BlockReference(ptStart, btrId))
            {
                //задаем позицию вставки блока
                br.Position = new Point3d(xInsert, ptStart.Y, ptStart.Z);
                //br.Annotative = AnnotativeStates.True;

                // добавляем вхождение блока на пространство модели и в транзакцию
                ms.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);
            }

        }

        public void AddText(Point3d ptStart, double xStart, string text)
        {
            using (DBText acText = new DBText())
            {
                acText.SetDatabaseDefaults();
                acText.Position = new Point3d(xStart, ptStart.Y - 5.0, ptStart.Z);
                acText.TextString = text;
                acText.HorizontalMode = (TextHorizontalMode)1;
                acText.AlignmentPoint = new Point3d(xStart, ptStart.Y - 5.0, ptStart.Z);
                acText.Annotative = AnnotativeStates.True;

                ms.AppendEntity(acText);
                tr.AddNewlyCreatedDBObject(acText, true);
            }
        }

    }
}