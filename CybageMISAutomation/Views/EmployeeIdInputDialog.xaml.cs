using System.Windows;

namespace CybageMISAutomation.Views
{
    public partial class EmployeeIdInputDialog : Window
    {
        public string EmployeeId { get; private set; } = string.Empty;
        
        public EmployeeIdInputDialog()
        {
            InitializeComponent();
            txtEmployeeId.Focus();
        }
        
        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            var employeeId = txtEmployeeId.Text?.Trim();
            
            if (string.IsNullOrWhiteSpace(employeeId))
            {
                MessageBox.Show("Please enter a valid Employee ID.", "Invalid Input", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                txtEmployeeId.Focus();
                return;
            }
            
            EmployeeId = employeeId;
            DialogResult = true;
            Close();
        }
        
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}