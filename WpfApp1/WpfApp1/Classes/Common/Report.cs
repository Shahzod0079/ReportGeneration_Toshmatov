using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using ReportGeneration_Toshmatov.Classes;
using ReportGeneration_Toshmatov.Pages;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace ReportGeneration_Toshmatov.Classes.Common
{
    public class Report
    {
        /// <summary> Метод создания отчёта о группе </summary>
        public static void Group(int IdGroup, Main main)
        {
            SaveFileDialog SFD = new SaveFileDialog()
            {
                InitialDirectory = @"C:\",

                Filter = "Excel (*.xlsx)|*.xlsx"
            };

            SFD.ShowDialog();

            if (SFD.FileName != "")
            {

                GroupContext Group = main.Allgroups.Find(x => x.Id == IdGroup);

                var ExcelApp = new Excel.Application();

                try
                {
                    ExcelApp.Visible = false;

                    Excel.Workbook Workbook = ExcelApp.Workbooks.Add(Type.Missing);

                    Excel.Worksheet Worksheet = Workbook.ActiveSheet;

                    (Worksheet.Cells[1, 1] as Excel.Range).Value = $"Отчёт о группе {Group.Name}";

                    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 5]].Merge();

                    Styles(Worksheet.Cells[1, 1], 18);

                    (Worksheet.Cells[3, 1] as Excel.Range).Value = $"Список группы";

                    Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 5]].Merge();

                    Styles(Worksheet.Cells[3, 1], 12, Excel.XlHAlign.xlHAlignLeft);

                    (Worksheet.Cells[4, 1] as Excel.Range).Value = $"ФИО";

                    Styles(Worksheet.Cells[4, 1], 12, Excel.XlHAlign.xlHAlignCenter, true);

                    (Worksheet.Cells[4, 1] as Excel.Range).ColumnWidth = 35.0f;

                    (Worksheet.Cells[4, 2] as Excel.Range).Value = $"Кол-во не сданных практических";

                    Styles(Worksheet.Cells[4, 2], 12, Excel.XlHAlign.xlHAlignCenter, true);
                    (Worksheet.Cells[4, 3] as Excel.Range).Value = $"Кол-во не сданных теоретических";

                    Styles(Worksheet.Cells[4, 3], 12, Excel.XlHAlign.xlHAlignCenter, true);

                    (Worksheet.Cells[4, 4] as Excel.Range).Value = $"Отсутствовал на паре";

                    Styles(Worksheet.Cells[4, 4], 12, Excel.XlHAlign.xlHAlignCenter, true);

                    (Worksheet.Cells[4, 5] as Excel.Range).Value = $"Опоздал";

                    Styles(Worksheet.Cells[4, 5], 12, Excel.XlHAlign.xlHAlignCenter, true);

                    int Height = 5;

                    List<StudentContext> Students = main.AllStudents.FindAll(x => x.IdGroup == IdGroup);

                    foreach (StudentContext Student in Students)
                    {
                        List<DisciplineContext> StudentDisciplines = main.AllDisciplines.FindAll(
                            x => x.IdGroup == Student.IdGroup);


                        int PracticeCount = 0;

                        int TheoryCount = 0;

                        int AbsenteeismCount = 0;

                        int LateCount = 0;


                        foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
                        {

                            List<WorkContext> StudentWorks = main.AllWorks.FindAll(x => x.IdDiscipline == StudentDiscipline.Id);

                            foreach (WorkContext StudentWork in StudentWorks)
                            {
                                // Получаем оценку за работу
                                EvaluationContext Evaluation = main.AllEvaluation.Find(x =>
                                    x.IdWork == StudentWork.Id &&
                                    x.IdStudent == Student.Id);

                                // Если оценки нет, или она пустая, или равно 2
                                if ((Evaluation != null && (Evaluation.Value.Trim() == "" || Evaluation.Value.Trim() == "2"))
                                    || Evaluation == null)
                                {
                                    // Если практика
                                    if (StudentWork.IdType == 1)
                                        // Считаем не сданную работу
                                        PracticeCount++;
                                    // Если теория
                                    else if (StudentWork.IdType == 2)
                                        // Считаем не сданную работу
                                        TheoryCount++;
                                }

                                // Проверяем что оценка не отсутствует и стоит пропуск
                                if (Evaluation != null && !string.IsNullOrWhiteSpace(Evaluation.Lateness))
                                {
                                    // Если пропуск 90 минут
                                    if (Convert.ToInt32(Evaluation.Lateness) == 90)
                                        // Считаем как пропущенную пару
                                        AbsenteeismCount++;
                                    else
                                        // Считаем как опоздание
                                        LateCount++;
                                }
                            }
                        }

                        // Обращаемся к ячейке, указываем текст
                        (Worksheet.Cells[Height, 1] as Excel.Range).Value = $"({Student.Lastname}) {Student.Firstname}";
                        // Присваиваем стили
                        Styles(Worksheet.Cells[Height, 1], 12, Excel.XlHAlign.xlHAlignLeft, true);

                        // Обращаемся к ячейке, указываем текст
                        (Worksheet.Cells[Height, 2] as Excel.Range).Value = PracticeCount.ToString();
                        // Присваиваем стили
                        Styles(Worksheet.Cells[Height, 2], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        // Обращаемся к ячейке, указываем текст
                        (Worksheet.Cells[Height, 3] as Excel.Range).Value = TheoryCount.ToString();
                        // Присваиваем стили
                        Styles(Worksheet.Cells[Height, 3], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        // Обращаемся к ячейке, указываем текст
                        (Worksheet.Cells[Height, 4] as Excel.Range).Value = AbsenteeismCount.ToString();
                        // Присваиваем стили
                        Styles(Worksheet.Cells[Height, 4], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        // Обращаемся к ячейке, указываем текст
                        (Worksheet.Cells[Height, 5] as Excel.Range).Value = LateCount.ToString();
                        // Присваиваем стили
                        Styles(Worksheet.Cells[Height, 5], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        Height++;
                    }

                    // Сохраняем документ
                    Workbook.SaveAs(SFD.FileName);
                    // Добавьте ЭТОТ КОД после заполнения всех студентов (перед Workbook.SaveAs)

                    // Находим лучшего студента
                    int bestStudentRow = -1;
                    double bestScore = -1;

                    for (int row = 5; row < Height; row++)
                    {
                        // Получаем данные из колонок
                        int practiceCount = Convert.ToInt32((Worksheet.Cells[row, 2] as Excel.Range).Value ?? 0);
                        int theoryCount = Convert.ToInt32((Worksheet.Cells[row, 3] as Excel.Range).Value ?? 0);
                        int absenteeismCount = Convert.ToInt32((Worksheet.Cells[row, 4] as Excel.Range).Value ?? 0);
                        int lateCount = Convert.ToInt32((Worksheet.Cells[row, 5] as Excel.Range).Value ?? 0);

                        // Формула успешности (чем меньше пропусков и не сданных работ, тем лучше)
                        double studentScore = (practiceCount + theoryCount) * -1 - (absenteeismCount * 2 + lateCount * 0.5);

                        if (studentScore > bestScore)
                        {
                            bestScore = studentScore;
                            bestStudentRow = row;
                        }
                    }

                    // Выделяем лучшего студента желтым цветом
                    if (bestStudentRow != -1)
                    {
                        Excel.Range bestRange = Worksheet.Rows[bestStudentRow];
                        bestRange.Interior.Color = 255; // Желтый цвет
                        bestRange.Font.Bold = true;
                    }

                    // Закрываем книгу
                    Workbook.Close();

                    // Закрываем Excel
                    ExcelApp.Quit();
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Ошибка при создании отчёта: {ex.Message}");
                }
            }
        }

        /// <summary> Применение стилей </summary>
        public static void Styles(Excel.Range Cell,
            int FontSize,
            Excel.XlHAlign Position = Excel.XlHAlign.xlHAlignCenter,
            bool Border = false)
        {
            // Присваиваем шрифт
            Cell.Font.Name = "Bahnschrift Light Condensed";

            // Присваиваем размер
            Cell.Font.Size = FontSize;

            // Указываем горизонтальное центрирование
            Cell.HorizontalAlignment = Position;

            // Указываем вертикальное центрирование
            Cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;


            if (Border)
            {

                Excel.Borders border = Cell.Borders;

 
                border.LineStyle = Excel.XlLineStyle.xlContinuous;

                border.Weight = Excel.XlBorderWeight.xlThin;

                Cell.WrapText = true;
            }
        }

    }
}