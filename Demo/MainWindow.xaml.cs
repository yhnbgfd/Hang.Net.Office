using System;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Demo
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private string _log;

        private string _wordExtend;

        public string Log
        {
            get
            {
                return _log;
            }
            set
            {
                _log = value; NotifyPropertyChanged("Log");
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        protected void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private void TestCompleted(Task obj)
        {
            WriteLine("Test Completed");
        }

        private void WriteLine(string log, params string[] parm)
        {
            Log += string.Format(log, parm) + "\n";
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            TextBox_Log.TextChanged += (s, ee) =>
            {
                TextBox_Log.ScrollToEnd();
            };
        }

        private void Button_TestWord_MicrosoftWord_Click(object sender, RoutedEventArgs e)
        {
            Task.Factory.StartNew(TestMsWord).ContinueWith(TestCompleted);
        }

        private void TestMsWord()
        {
            TestWord tw = new TestWord();
            tw.TestMicrosoftWord(_wordExtend);
        }

        private void Button_TestWord_SpireDoc_Click(object sender, RoutedEventArgs e)
        {
            TestWord tw = new TestWord();
            Task.Factory.StartNew(tw.TestSpireDoc).ContinueWith(TestCompleted);
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Environment.Exit(Environment.ExitCode);
        }

        private void ComboBox_WordExtend_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _wordExtend = (ComboBox_WordExtend.SelectedItem as ComboBoxItem).Content.ToString();
        }
    }
}
