using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using ReportGeneration_Toshmatov.Classes;
using ReportGeneration_Toshmatov.Pages;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGeneration_Toshmatov.Classes.Common
{
    public class Report
    {
        public static void Group(int IdGroup, Main main)
        {
            SaveFileDialog SFD = new SaveFileDialog()
            {
                InitialDirectory = @"C:\",
                Filter = "Excel (*.xlsx)|*.xlsx"
            };

            if (SFD.ShowDialog() == true && SFD.FileName != "")
            {
                GroupContext Group = main.Allgroups.Find(x => x.Id == IdGroup);
                Excel.Application ExcelApp = null;
                Excel.Workbook Workbook = null;

                try
                {
                    ExcelApp = new Excel.Application();
                    ExcelApp.Visible = false;
                    ExcelApp.DisplayAlerts = false;

                    Workbook = ExcelApp.Workbooks.Add();
                    Excel.Worksheet Worksheet = Workbook.ActiveSheet;

                    // Заголовки
                    ((Excel.Range)Worksheet.Cells[1, 1]).Value = $"Отчёт о группе {Group.Name}";
                    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 5]].Merge();
                    Styles((Excel.Range)Worksheet.Cells[1, 1], 18);

                    ((Excel.Range)Worksheet.Cells[3, 1]).Value = $"Список группы";
                    Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 5]].Merge();
                    Styles((Excel.Range)Worksheet.Cells[3, 1], 12, Excel.XlHAlign.xlHAlignLeft);

                    string[] headers = { "ФИО", "Кол-во не сданных практических", "Кол-во не сданных теоретических", "Отсутствовал на паре", "Опоздал" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        ((Excel.Range)Worksheet.Cells[4, i + 1]).Value = headers[i];
                        Styles((Excel.Range)Worksheet.Cells[4, i + 1], 12, Excel.XlHAlign.xlHAlignCenter, true);
                    }
                    ((Excel.Range)Worksheet.Cells[4, 1]).ColumnWidth = 35;

                    int Height = 5;
                    List<StudentContext> Students = main.AllStudents.FindAll(x => x.IdGroup == IdGroup);

                    foreach (StudentContext Student in Students)
                    {
                        List<DisciplineContext> StudentDisciplines = main.AllDisciplines.FindAll(x => x.IdGroup == Student.IdGroup);

                        int PracticeCount = 0, TheoryCount = 0, AbsenteeismCount = 0, LateCount = 0;

                        foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
                        {
                            List<WorkContext> StudentWorks = main.AllWorks.FindAll(x => x.IdDiscipline == StudentDiscipline.Id);

                            foreach (WorkContext StudentWork in StudentWorks)
                            {
                                EvaluationContext Evaluation = main.AllEvaluation.Find(x =>
                                    x.IdWork == StudentWork.Id && x.IdStudent == Student.Id);

                                if ((Evaluation != null && (Evaluation.Value.Trim() == "" || Evaluation.Value.Trim() == "2")) || Evaluation == null)
                                {
                                    if (StudentWork.IdType == 1) PracticeCount++;
                                    else if (StudentWork.IdType == 2) TheoryCount++;
                                }

                                if (Evaluation != null && !string.IsNullOrWhiteSpace(Evaluation.Lateness))
                                {
                                    if (Convert.ToInt32(Evaluation.Lateness) == 90)
                                        AbsenteeismCount++;
                                    else
                                        LateCount++;
                                }
                            }
                        }

                        ((Excel.Range)Worksheet.Cells[Height, 1]).Value = $"({Student.Lastname}) {Student.Firstname}";
                        Styles((Excel.Range)Worksheet.Cells[Height, 1], 12, Excel.XlHAlign.xlHAlignLeft, true);
                        ((Excel.Range)Worksheet.Cells[Height, 2]).Value = PracticeCount.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 2], 12, Excel.XlHAlign.xlHAlignCenter, true);
                        ((Excel.Range)Worksheet.Cells[Height, 3]).Value = TheoryCount.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 3], 12, Excel.XlHAlign.xlHAlignCenter, true);
                        ((Excel.Range)Worksheet.Cells[Height, 4]).Value = AbsenteeismCount.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 4], 12, Excel.XlHAlign.xlHAlignCenter, true);
                        ((Excel.Range)Worksheet.Cells[Height, 5]).Value = LateCount.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 5], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        Height++;
                    }

                    // ОЦЕНКА ХОРОШО
                    int bestStudentRow = -1;
                    double bestScore = -1;

                    for (int row = 5; row < Height; row++)
                    {
                        int practiceCount = Convert.ToInt32(((Excel.Range)Worksheet.Cells[row, 2]).Value ?? 0);
                        int theoryCount = Convert.ToInt32(((Excel.Range)Worksheet.Cells[row, 3]).Value ?? 0);
                        int absenteeismCount = Convert.ToInt32(((Excel.Range)Worksheet.Cells[row, 4]).Value ?? 0);
                        int lateCount = Convert.ToInt32(((Excel.Range)Worksheet.Cells[row, 5]).Value ?? 0);

                        double studentScore = (practiceCount + theoryCount) * -1 - (absenteeismCount * 2 + lateCount * 0.5);

                        if (studentScore > bestScore)
                        {
                            bestScore = studentScore;
                            bestStudentRow = row;
                        }
                    }

                    if (bestStudentRow != -1)
                    {
                        Excel.Range bestRange = Worksheet.Rows[bestStudentRow];
                        bestRange.Interior.Color = 65535;
                        bestRange.Font.Bold = true;
                    }

                    // ОЦЕНКА ОТЛИЧНО - упрощенная версия
                    int sheetCounter = 1;
                    foreach (StudentContext Student in Students)
                    {
                        try
                        {
                            Excel.Worksheet StudentSheet = Workbook.Worksheets.Add();
                            StudentSheet.Name = Student.Lastname;

                            ((Excel.Range)StudentSheet.Cells[1, 1]).Value = $"Успеваемость студента {Student.Lastname} {Student.Firstname}";
                            StudentSheet.Range[StudentSheet.Cells[1, 1], StudentSheet.Cells[1, 4]].Merge();
                            Styles((Excel.Range)StudentSheet.Cells[1, 1], 16);

                            string[] studentHeaders = { "Дисциплина", "Работа", "Тип", "Оценка" };
                            for (int col = 0; col < studentHeaders.Length; col++)
                            {
                                ((Excel.Range)StudentSheet.Cells[3, col + 1]).Value = studentHeaders[col];
                                Styles((Excel.Range)StudentSheet.Cells[3, col + 1], 12, Excel.XlHAlign.xlHAlignCenter, true);
                                ((Excel.Range)StudentSheet.Cells[3, col + 1]).ColumnWidth = 25;
                            }

                            int row = 4;
                            List<DisciplineContext> StudentDisciplines = main.AllDisciplines.FindAll(x => x.IdGroup == Student.IdGroup);

                            foreach (DisciplineContext Discipline in StudentDisciplines)
                            {
                                List<WorkContext> DisciplineWorks = main.AllWorks.FindAll(x => x.IdDiscipline == Discipline.Id);

                                foreach (WorkContext Work in DisciplineWorks)
                                {
                                    EvaluationContext Evaluation = main.AllEvaluation.Find(x =>
                                        x.IdWork == Work.Id && x.IdStudent == Student.Id);

                                    string оценка = "не сдано";
                                    if (Evaluation != null && !string.IsNullOrWhiteSpace(Evaluation.Value))
                                        оценка = Evaluation.Value;

                                    ((Excel.Range)StudentSheet.Cells[row, 1]).Value = Discipline.Name;
                                    ((Excel.Range)StudentSheet.Cells[row, 2]).Value = Work.Name;

                                    string тип = Work.IdType == 1 ? "Практика" : Work.IdType == 2 ? "Теория" : "Другое";
                                    ((Excel.Range)StudentSheet.Cells[row, 3]).Value = тип;
                                    ((Excel.Range)StudentSheet.Cells[row, 4]).Value = оценка;

                                    row++;
                                }
                            }
                            sheetCounter++;
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show($"Ошибка при создании листа для {Student.Lastname}: {ex.Message}");
                        }
                    }

                    Workbook.SaveAs(SFD.FileName);
                    Workbook.Close();
                    ExcelApp.Quit();

                    System.Windows.MessageBox.Show($"Отчёт успешно сохранён!\nЛучший студент выделен жёлтым цветом.");
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Ошибка при создании отчёта: {ex.Message}");
                }
                finally
                {
                    if (Workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook);
                    if (ExcelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                }
            }
        }

        public static void Styles(Excel.Range Cell, int FontSize,
            Excel.XlHAlign Position = Excel.XlHAlign.xlHAlignCenter, bool Border = false)
        {
            try
            {
                Cell.Font.Name = "Bahnschrift Light Condensed";
                Cell.Font.Size = FontSize;
                Cell.HorizontalAlignment = Position;
                Cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                if (Border)
                {
                    Cell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    Cell.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    Cell.WrapText = true;
                }
            }
            catch { }
        }
    }
}