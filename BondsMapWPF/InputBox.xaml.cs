using System;
using System.Windows;

namespace BondsMapWPF
{
    /// <summary>
    /// Логика взаимодействия для InputBox.xaml
    /// </summary>
    public partial class InputBox
    {
        private readonly Action<string> _action;
        public InputBox(string defaultName, Action<string> action)
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
            _action(NameTextBox.Text);
            Close();
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
