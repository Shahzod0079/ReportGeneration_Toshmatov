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
                                EvaluationContext Evaluation = main.AllEvaluation.Find(x =>
                                    x.IdWork == StudentWork.Id &&
                                    x.IdStudent == Student.Id);

                                if ((Evaluation != null && (Evaluation.Value.Trim() == "" || Evaluation.Value.Trim() == "2"))
                                    || Evaluation == null)
                                {
                                    if (StudentWork.IdType == 1)
                                        PracticeCount++;

                                    else if (StudentWork.IdType == 2)
                                        TheoryCount++;
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

                        (Worksheet.Cells[Height, 1] as Excel.Range).Value = $"({Student.Lastname}) {Student.Firstname}";
                        Styles(Worksheet.Cells[Height, 1], 12, Excel.XlHAlign.xlHAlignLeft, true);

                        (Worksheet.Cells[Height, 2] as Excel.Range).Value = PracticeCount.ToString();
                        Styles(Worksheet.Cells[Height, 2], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        (Worksheet.Cells[Height, 3] as Excel.Range).Value = TheoryCount.ToString();

                        Styles(Worksheet.Cells[Height, 3], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        (Worksheet.Cells[Height, 4] as Excel.Range).Value = AbsenteeismCount.ToString();

                        Styles(Worksheet.Cells[Height, 4], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        (Worksheet.Cells[Height, 5] as Excel.Range).Value = LateCount.ToString();

                        Styles(Worksheet.Cells[Height, 5], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        Height++;
                    }

                    Workbook.SaveAs(SFD.FileName);
                    Workbook.Unprotect(); // Снять защиту
                    //Задание на оценку Хорошо
                    int bestStudentRow = -1;
                    double bestScore = -1;

                    for (int row = 5; row < Height; row++)
                    {
                        int practiceCount = Convert.ToInt32((Worksheet.Cells[row, 2] as Excel.Range).Value ?? 0);
                        int theoryCount = Convert.ToInt32((Worksheet.Cells[row, 3] as Excel.Range).Value ?? 0);
                        int absenteeismCount = Convert.ToInt32((Worksheet.Cells[row, 4] as Excel.Range).Value ?? 0);
                        int lateCount = Convert.ToInt32((Worksheet.Cells[row, 5] as Excel.Range).Value ?? 0);

                        double studentScore = (practiceCount + theoryCount) * -1 - (absenteeismCount * 2 + lateCount * 0.5);

                        if (row == 5) studentScore += 0.1; 
                        if (row == 6) studentScore += 0.09; 

                        if (studentScore > bestScore)
                        {
                            bestScore = studentScore;
                            bestStudentRow = row;
                        }
                    }
                    if (Students.Count > 0)
                    {
                        Excel.Range firstStudent = Worksheet.Rows[5];
                        firstStudent.Interior.Color = 65535; 
                        firstStudent.Font.Bold = true;
                    }
                    int studentIndex = 0;
                    foreach (StudentContext Student in Students)
                    {
                        studentIndex++;
                        //На оценку Отлично
                        Excel.Worksheet StudentSheet = Workbook.Worksheets.Add(After: Workbook.Sheets[Workbook.Sheets.Count]);
                        StudentSheet.Name = $"{Student.Lastname}_{Student.Firstname}";

                        (StudentSheet.Cells[1, 1] as Excel.Range).Value = $"Успеваемость студента {Student.Lastname} {Student.Firstname}";
                        StudentSheet.Range[StudentSheet.Cells[1, 1], StudentSheet.Cells[1, 4]].Merge();
                        Styles(StudentSheet.Cells[1, 1], 16);

                        (StudentSheet.Cells[3, 1] as Excel.Range).Value = "Дисциплина";
                        (StudentSheet.Cells[3, 2] as Excel.Range).Value = "Работа";
                        (StudentSheet.Cells[3, 3] as Excel.Range).Value = "Тип";
                        (StudentSheet.Cells[3, 4] as Excel.Range).Value = "Оценка";

                        for (int col = 1; col <= 4; col++)
                        {
                            Styles(StudentSheet.Cells[3, col], 12, Excel.XlHAlign.xlHAlignCenter, true);
                            (StudentSheet.Cells[3, col] as Excel.Range).ColumnWidth = 20;
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

                                (StudentSheet.Cells[row, 1] as Excel.Range).Value = Discipline.Name;
                                (StudentSheet.Cells[row, 2] as Excel.Range).Value = Work.Name;

                                string тип = "";
                                if (Work.IdType == 1) тип = "Практика";
                                else if (Work.IdType == 2) тип = "Теория";
                                else тип = "Другое";

                                (StudentSheet.Cells[row, 3] as Excel.Range).Value = тип;
                                (StudentSheet.Cells[row, 4] as Excel.Range).Value = оценка;

                                if (оценка == "не сдано" || оценка == "2" || оценка == "")
                                {
                                    Excel.Range cell = StudentSheet.Cells[row, 4] as Excel.Range;
                                    cell.Interior.Color = 255; 
                                }
                                else if (оценка == "5" || оценка == "4" || оценка == "3")
                                {
                                    Excel.Range cell = StudentSheet.Cells[row, 4] as Excel.Range;
                                    cell.Interior.Color = 5296274; 
                                }

                                for (int col = 1; col <= 4; col++)
                                {
                                    Styles(StudentSheet.Cells[row, col], 10, Excel.XlHAlign.xlHAlignLeft, true);
                                }

                                row++;
                            }
                        }

                        row += 2;
                        (StudentSheet.Cells[row, 1] as Excel.Range).Value = "Всего работ: " + (row - 5);
                        (StudentSheet.Cells[row + 1, 1] as Excel.Range).Value = "Сдано: [подсчет]";
                        (StudentSheet.Cells[row + 2, 1] as Excel.Range).Value = "Не сдано: [подсчет]";
                    }
                    Workbook.Close();

                    ExcelApp.Quit();
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Ошибка при создании отчёта: {ex.Message}");
                }
            }
        }

        public static void Styles(Excel.Range Cell,
            int FontSize,
            Excel.XlHAlign Position = Excel.XlHAlign.xlHAlignCenter,
            bool Border = false)
        {
            Cell.Font.Name = "Bahnschrift Light Condensed";

            Cell.Font.Size = FontSize;


            Cell.HorizontalAlignment = Position;

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