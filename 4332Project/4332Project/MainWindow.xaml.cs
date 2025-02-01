using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using _4332Project.Students;

namespace _4332Project
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SidorovButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Имя: Игнат\nВозраст: 18 лет");
        }

        private void ShowAuthorInfo_Click(object sender, RoutedEventArgs e)
        {
            var newForm = new Novikov_4332();
            newForm.Show();
        }

        private void SulimovButton_Click(object sender, RoutedEventArgs e)
        {
            var newForm = new Sulimov_4332();
            newForm.Show();
        }

        private void RomanovButton_OnClick(object sender, RoutedEventArgs e)
        {
            new Romanov_4332().Show();
        }

        private void ZaripovButton_OnClick(object sender, RoutedEventArgs e)
        {
            new Zaripov_4332().Show();
        }

        private void SayfutdinovaButton_Click(object sender, RoutedEventArgs e)
        {
            new Sayfut4332().Show();
        }
        
        private void SalakhovButton_Click(object sender, RoutedEventArgs e)
        {
            new Salakhov_4332().Show();
        }

        private void StepanovButton_Click(object sender, RoutedEventArgs e)
        {
            new Stepanov_4332().Show();
        }
    }
}
