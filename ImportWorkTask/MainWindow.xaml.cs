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
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using NewEventLogDLL;
using WorkTaskDLL;

namespace ImportWorkTask
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        WorkTaskClass TheWorkTaskClass = new WorkTaskClass();

        ImportWorkTaskDataSet TheImportWorkTaskDataSet = new ImportWorkTaskDataSet();
        FindWorkTaskByWorkTaskDataSet TheFindWorkTaskByWorkTaskDataSet = new FindWorkTaskByWorkTaskDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            bool blnFatalError;
            string strWorkTask;

            try
            {
                intNumberOfRecords = TheImportWorkTaskDataSet.worktask.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strWorkTask = TheImportWorkTaskDataSet.worktask[intCounter].WorkTask;

                    blnFatalError = TheWorkTaskClass.InsertWorkTask(strWorkTask, 1);

                    if (blnFatalError == true)
                        throw new Exception();
                }
            }
            catch(Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Work Task // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void btnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            string strWorkTask;
            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;

            try
            {
                TheImportWorkTaskDataSet.worktask.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count - 1;
                intColumnRange = range.Columns.Count;

                for (intCounter = 2; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strWorkTask = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2);

                    TheFindWorkTaskByWorkTaskDataSet = TheWorkTaskClass.FindWorkTaskByWorkTask(strWorkTask);

                    intRecordsReturned = TheFindWorkTaskByWorkTaskDataSet.FindWorkTaskByWorkTask.Rows.Count;

                    if (intRecordsReturned == 0)
                    {

                        ImportWorkTaskDataSet.worktaskRow NewTaskRow = TheImportWorkTaskDataSet.worktask.NewworktaskRow();

                        NewTaskRow.WorkTask = strWorkTask.ToUpper();
                        NewTaskRow.TaskCost = 1;

                        TheImportWorkTaskDataSet.worktask.Rows.Add(NewTaskRow);
                    }
                }

                PleaseWait.Close();
                dgrExcel.ItemsSource = TheImportWorkTaskDataSet.worktask;
                btnProcess.IsEnabled = true;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Inventory Project // Import Inventory Count // Select Spreadsheet Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
