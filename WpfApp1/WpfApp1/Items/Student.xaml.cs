using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using ReportGeneration_Toshmatov.Classes;
using WpfApp1.Pages;

namespace ReportGeneration_Toshmatov.Items
{
    public partial class Student : UserControl
    {
        private StudentContext student;
        private Main mainPage;

        public Student(StudentContext student, Main mainPage) 
        {
            this.student = student;
            this.mainPage = mainPage; 
            InitializeComponent();

            TBFio.Text = $"{student.Lastname} {student.Firstname}";

            CBExplored.IsChecked = student.Expelled;

            // Используем mainPage вместо Main
            List<DisciplineContext> StudentDisciplines = mainPage.AllDisciplines.FindAll(
                x => x.IdGroup == student.IdGroup);

            int NecessarilyCount = 0;
            int WorksCount = 0;
            int DoneCount = 0;
            int MissedCount = 0;

            foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
            {
                List<WorkContext> StudentWorks = mainPage.AllWorks.FindAll(x =>
                    (x.IdType == 1 || x.IdType == 2 || x.IdType == 3) &&
                    x.IdDiscipline == StudentDiscipline.Id);

                NecessarilyCount += StudentWorks.Count;

                foreach (WorkContext StudentWork in StudentWorks)
                {
                    EvaluationContext Evaluation = mainPage.AllEvaluation.Find(x =>
                        x.IdWork == StudentWork.Id &&
                        x.IdStudent == student.Id);

                    if (Evaluation != null && !string.IsNullOrWhiteSpace(Evaluation.Value) && Evaluation.Value != "2")
                        DoneCount++;
                }
            }

            List<WorkContext> AllStudentWorks = mainPage.AllWorks.FindAll(x =>
                x.IdType != 4 && x.IdType != 3);

            WorksCount += AllStudentWorks.Count;

            foreach (WorkContext StudentWork in AllStudentWorks)
            {
                EvaluationContext Evaluation = mainPage.AllEvaluation.Find(x =>
                    x.IdWork == StudentWork.Id &&
                    x.IdStudent == student.Id);

                if (Evaluation != null && !string.IsNullOrWhiteSpace(Evaluation.Lateness))
                {
                    if (int.TryParse(Evaluation.Lateness, out int lateness))
                        MissedCount += lateness;
                }
            }

            if (NecessarilyCount > 0)
                doneWorks.Value = (100f / NecessarilyCount) * DoneCount;

            if (WorksCount > 0)
                missedCount.Value = (100f / (WorksCount * 90f)) * MissedCount;

            TBGroup.Text = mainPage.Allgroups.Find(x => x.Id == student.IdGroup)?.Name ?? "";
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e) { }
    }
}