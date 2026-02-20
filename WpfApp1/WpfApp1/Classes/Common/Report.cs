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

                    string[] headers = { "ФИО", "Кол-во пятерок", "Кол-во четверок", "Кол-во троек", "Кол-во двоек", "Несдано", "Пропуски", "Опоздал" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        ((Excel.Range)Worksheet.Cells[4, i + 1]).Value = headers[i];
                        Styles((Excel.Range)Worksheet.Cells[4, i + 1], 12, Excel.XlHAlign.xlHAlignCenter, true);
                    }
                    ((Excel.Range)Worksheet.Cells[4, 1]).ColumnWidth = 35;

                    int Height = 5;
                    List<StudentContext> Students = main.AllStudents.FindAll(x => x.IdGroup == IdGroup);

                    int[,] studentStats = new int[Students.Count, 8]; 

                    for (int s = 0; s < Students.Count; s++)
                    {
                        StudentContext Student = Students[s];
                        List<DisciplineContext> StudentDisciplines = main.AllDisciplines.FindAll(x => x.IdGroup == Student.IdGroup);

                        int count5 = 0, count4 = 0, count3 = 0, count2 = 0, count0 = 0;
                        int пропуски = 0, опоздания = 0;

                        foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
                        {
                            List<WorkContext> StudentWorks = main.AllWorks.FindAll(x => x.IdDiscipline == StudentDiscipline.Id);

                            foreach (WorkContext StudentWork in StudentWorks)
                            {
                                EvaluationContext Evaluation = main.AllEvaluation.Find(x =>
                                    x.IdWork == StudentWork.Id && x.IdStudent == Student.Id);

                                if (Evaluation != null)
                                {
                                    string val = Evaluation.Value.Trim();
                                    if (val == "5") count5++;
                                    else if (val == "4") count4++;
                                    else if (val == "3") count3++;
                                    else if (val == "2") count2++;
                                    else if (val == "") count0++;
                                }
                                else
                                {
                                    count0++;
                                }

                                if (Evaluation != null && !string.IsNullOrWhiteSpace(Evaluation.Lateness))
                                {
                                    if (Convert.ToInt32(Evaluation.Lateness) == 90)
                                        пропуски++;
                                    else
                                        опоздания++;
                                }
                            }
                        }

                        studentStats[s, 0] = s;
                        studentStats[s, 1] = count5;
                        studentStats[s, 2] = count4;
                        studentStats[s, 3] = count3;
                        studentStats[s, 4] = count2;
                        studentStats[s, 5] = count0;
                        studentStats[s, 6] = пропуски;
                        studentStats[s, 7] = опоздания;

                        ((Excel.Range)Worksheet.Cells[Height, 1]).Value = $"({Student.Lastname}) {Student.Firstname}";
                        Styles((Excel.Range)Worksheet.Cells[Height, 1], 12, Excel.XlHAlign.xlHAlignLeft, true);
                        ((Excel.Range)Worksheet.Cells[Height, 2]).Value = count5.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 2], 12, Excel.XlHAlign.xlHAlignCenter, true);
                        ((Excel.Range)Worksheet.Cells[Height, 3]).Value = count4.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 3], 12, Excel.XlHAlign.xlHAlignCenter, true);
                        ((Excel.Range)Worksheet.Cells[Height, 4]).Value = count3.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 4], 12, Excel.XlHAlign.xlHAlignCenter, true);
                        ((Excel.Range)Worksheet.Cells[Height, 5]).Value = count2.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 5], 12, Excel.XlHAlign.xlHAlignCenter, true);
                        ((Excel.Range)Worksheet.Cells[Height, 6]).Value = count0.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 6], 12, Excel.XlHAlign.xlHAlignCenter, true);
                        ((Excel.Range)Worksheet.Cells[Height, 7]).Value = пропуски.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 7], 12, Excel.XlHAlign.xlHAlignCenter, true);
                        ((Excel.Range)Worksheet.Cells[Height, 8]).Value = опоздания.ToString();
                        Styles((Excel.Range)Worksheet.Cells[Height, 8], 12, Excel.XlHAlign.xlHAlignCenter, true);

                        Height++;
                    }
                    Excel.Worksheet StatsSheet = Workbook.Worksheets.Add();
                    StatsSheet.Name = "Статистика";

                    ((Excel.Range)StatsSheet.Cells[1, 1]).Value = "Студент";
                    ((Excel.Range)StatsSheet.Cells[1, 2]).Value = "5";
                    ((Excel.Range)StatsSheet.Cells[1, 3]).Value = "4";
                    ((Excel.Range)StatsSheet.Cells[1, 4]).Value = "3";
                    ((Excel.Range)StatsSheet.Cells[1, 5]).Value = "2";
                    ((Excel.Range)StatsSheet.Cells[1, 6]).Value = "Несдано";
                    ((Excel.Range)StatsSheet.Cells[1, 7]).Value = "Пропуски";
                    ((Excel.Range)StatsSheet.Cells[1, 8]).Value = "Опоздания";

                    for (int s = 0; s < Students.Count; s++)
                    {
                        ((Excel.Range)StatsSheet.Cells[s + 2, 1]).Value = Students[s].Lastname;  
                        ((Excel.Range)StatsSheet.Cells[s + 2, 2]).Value = studentStats[s, 1];
                        ((Excel.Range)StatsSheet.Cells[s + 2, 3]).Value = studentStats[s, 2];
                        ((Excel.Range)StatsSheet.Cells[s + 2, 4]).Value = studentStats[s, 3];
                        ((Excel.Range)StatsSheet.Cells[s + 2, 5]).Value = studentStats[s, 4];
                        ((Excel.Range)StatsSheet.Cells[s + 2, 6]).Value = studentStats[s, 5];
                        ((Excel.Range)StatsSheet.Cells[s + 2, 7]).Value = studentStats[s, 6];
                        ((Excel.Range)StatsSheet.Cells[s + 2, 8]).Value = studentStats[s, 7];
                    }
                    int bestStudentRow = -1;
                    int worstStudentRow = -1;
                    double bestScore = double.MinValue;
                    double worstScore = double.MaxValue;

                    for (int row = 5; row < Height; row++)
                    {
                        int idx = row - 5;
                        int count5 = studentStats[idx, 1];
                        int count4 = studentStats[idx, 2];
                        int count3 = studentStats[idx, 3];
                        int count2 = studentStats[idx, 4];
                        int count0 = studentStats[idx, 5];
                        int пропуски = studentStats[idx, 6];
                        int опоздания = studentStats[idx, 7];

                        double studentScore = (count5 * 5) + (count4 * 4) + (count3 * 3) +
                                              (count2 * 1) - (count0 * 2) - (пропуски * 3) - (опоздания * 1);

                        if (studentScore > bestScore)
                        {
                            bestScore = studentScore;
                            bestStudentRow = row;
                        }

                        if (studentScore < worstScore)
                        {
                            worstScore = studentScore;
                            worstStudentRow = row;
                        }
                    }

                    // Выделение лучшего (зеленый) и худшего (красный)
                    if (bestStudentRow != -1)
                    {
                        Excel.Range bestRange = Worksheet.Rows[bestStudentRow];
                        bestRange.Interior.Color = 5296274; // Зеленый
                        bestRange.Font.Bold = true;
                    }

                    if (worstStudentRow != -1 && worstStudentRow != bestStudentRow)
                    {
                        Excel.Range worstRange = Worksheet.Rows[worstStudentRow];
                        worstRange.Interior.Color = 255; // Красный
                        worstRange.Font.Bold = true;
                    }
                    // ОЦЕНКА ОТЛИЧНО - отдельные листы
                    int sheetCounter = 1;
                    foreach (StudentContext Student in Students)
                    {
                        try
                        {
                            Excel.Worksheet StudentSheet = Workbook.Worksheets.Add();
                            StudentSheet.Name = Student.Lastname.Length > 30 ? Student.Lastname.Substring(0, 30) : Student.Lastname;

                            // Заголовок
                            ((Excel.Range)StudentSheet.Cells[1, 1]).Value = $"Успеваемость студента {Student.Lastname} {Student.Firstname}";
                            StudentSheet.Range[StudentSheet.Cells[1, 1], StudentSheet.Cells[1, 4]].Merge();
                            Styles((Excel.Range)StudentSheet.Cells[1, 1], 16);

                            // Заголовки таблицы
                            string[] studentHeaders = { "Дисциплина", "Работа", "Тип", "Оценка" };
                            for (int col = 0; col < studentHeaders.Length; col++)
                            {
                                ((Excel.Range)StudentSheet.Cells[3, col + 1]).Value = studentHeaders[col];
                                Styles((Excel.Range)StudentSheet.Cells[3, col + 1], 12, Excel.XlHAlign.xlHAlignCenter, true);
                                ((Excel.Range)StudentSheet.Cells[3, col + 1]).ColumnWidth = 25;
                            }

                            int rowData = 4;
                            int count5_local = 0, count4_local = 0, count3_local = 0, count2_local = 0, count0_local = 0;
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

                                    // Подсчет статистики
                                    if (оценка == "5") count5_local++;
                                    else if (оценка == "4") count4_local++;
                                    else if (оценка == "3") count3_local++;
                                    else if (оценка == "2") count2_local++;
                                    else if (оценка == "не сдано" || оценка == "") count0_local++;

                                    ((Excel.Range)StudentSheet.Cells[rowData, 1]).Value = Discipline.Name;
                                    ((Excel.Range)StudentSheet.Cells[rowData, 2]).Value = Work.Name;

                                    string тип = Work.IdType == 1 ? "Практика" : Work.IdType == 2 ? "Теория" : "Другое";
                                    ((Excel.Range)StudentSheet.Cells[rowData, 3]).Value = тип;
                                    ((Excel.Range)StudentSheet.Cells[rowData, 4]).Value = оценка;

                                    // Цвет оценки
                                    Excel.Range оценкаCell = (Excel.Range)StudentSheet.Cells[rowData, 4];
                                    if (оценка == "5" || оценка == "4")
                                        оценкаCell.Interior.Color = 5296274;
                                    else if (оценка == "3")
                                        оценкаCell.Interior.Color = 65535;
                                    else if (оценка == "2" || оценка == "не сдано")
                                        оценкаCell.Interior.Color = 255;

                                    rowData++;
                                }
                            }

                            // Добавляем итоговую строку с подсчетом
                            int итоговаяСтрока = rowData + 2;

                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 1]).Value = "ИТОГО:";
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 1]).Font.Bold = true;

                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 2]).Value = $"5: {count5_local}";
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 2]).Font.Bold = true;

                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 3]).Value = $"4: {count4_local}";
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 3]).Font.Bold = true;

                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 4]).Value = $"3: {count3_local}";
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 4]).Font.Bold = true;

                            итоговаяСтрока++;
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 2]).Value = $"2: {count2_local}";
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 2]).Font.Bold = true;

                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 3]).Value = $"Несдано: {count0_local}";
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 3]).Font.Bold = true;

                            // Подсчет балла по формуле
                            double studentScoreLocal = (count5_local * 5) + (count4_local * 4) + (count3_local * 3) +
                                                        (count2_local * 1) - (count0_local * 2);

                            итоговаяСтрока++;
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 1]).Value = "Общий балл:";
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 1]).Font.Bold = true;
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 2]).Value = studentScoreLocal.ToString();
                            ((Excel.Range)StudentSheet.Cells[итоговаяСтрока, 2]).Font.Bold = true;

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

                    System.Windows.MessageBox.Show($"Отчёт успешно сохранён!\nЛучший студент - зеленый, худший - красный.");
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