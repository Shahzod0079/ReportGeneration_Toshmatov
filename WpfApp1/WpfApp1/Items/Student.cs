using System.Windows;
using ReportGeneration_Toshmatov.Classes;
using WpfApp1.Pages;

namespace Items
{
    internal class Student : UIElement
    {
        private StudentContext student;
        private Main main;

        public Student(StudentContext student, Main main)
        {
            this.student = student;
            this.main = main;
        }
    }
}