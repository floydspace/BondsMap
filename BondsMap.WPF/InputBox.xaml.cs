using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using MessageBox = System.Windows.Forms.MessageBox;

namespace BondsMapWPF
{
    /// <summary>
    /// Логика взаимодействия для InputBox.xaml
    /// </summary>
    public partial class InputBox
    {
        private readonly Func<string, bool> _action;
        public InputBox(string defaultName, Func<string, bool> action)
        {
            _action = action;
            
            InitializeComponent();

            NameTextBox.Text = defaultName;
            NameTextBox.SelectAll();
            NameTextBox.Focus();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (_action(NameTextBox.Text))
                Close();
            else
            {
                MessageBox.Show(@"Группа с таким именем уже существует!", @"Ошибка валидации",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                NameTextBox.SelectAll();
                NameTextBox.Focus();
            }
        }

        public void ShowDialog(Window owner)
        {
            Owner = owner;
            ShowDialog();
        }

        public void Show(Window owner)
        {
            Owner = owner;
            Show();
        }
    }
}
