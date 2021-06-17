using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace Unity_Excel_to_Scriptable_Exporter
{
    /////////////////////////////////////////////////////////////////
        
    public partial class MainWindow : Window
    {
        ////////////////////////////////

        public enum Script_DataType
        {
            STRING,
            BOOL,
            INT,
            FLOAT,
            COLOR,
        }

        ////////////////////////////////

        List<string> Script_FoundVariables;
        List<string> Excel_FoundVariables;
        List<string> Excel_FoundNames;

        List<CompletedDataset> Data_MatchedVariableSet_List;

        ////////////////////////////////

        string Script_CurrentFileLocation;
        string Excel_CurrentFileLocation;
        string FinalFile_OutputFileLocation;

        /////////////////////////////////////////////////////////////////

        public MainWindow()
        {
            //This is the Main Run Area of the App
            InitializeComponent();

            //Set Output Default Path
            FinalFile_OutputFileLocation = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            Textbox_SelectOutputLocation.Text = FinalFile_OutputFileLocation;

            //Refresh Buttons
            RefreshAllButtonStatuses();
        }

        /////////////////////////////////////////////////////////////////

        private void SelectScript_Button_Click(object sender, RoutedEventArgs e)
        {
            //Open File Picker
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "C-Sharp files (*.cs) | *.cs";
            openFileDialog.ValidateNames = false;

            //When File Picker Is Closed Open If Loop
            if (openFileDialog.ShowDialog() == true)
            {
                //Record Data For later Then Set Title of Path Location
                Script_CurrentFileLocation = openFileDialog.FileName;
                Textbox_SelectScriptFile.Text = System.IO.Path.GetFileName(Script_CurrentFileLocation);

                //Reset Found Variables List and add a deafult value
                Script_FoundVariables = new List<string>();
                Script_FoundVariables.Add("ScriptableName (STRING)");

                foreach (var line in File.ReadLines(openFileDialog.FileName))
                {
                    //Regex Match Pattern for a variable
                    if (Regex.IsMatch(line, WildCardToRegular("*public*;")))
                    {
                        string[] parts = line.Split(' ');
                        string lastWord_VariableName = parts[parts.Length - 1];
                        string secondLastWord_VariableType = parts[parts.Length - 2];
                        lastWord_VariableName = lastWord_VariableName.Remove(lastWord_VariableName.Length - 1);

                        //Add Matching String to List
                        Script_FoundVariables.Add(lastWord_VariableName + " (" + secondLastWord_VariableType.ToUpper() + ")");
                    }
                }


                ScriptVariables_Textbox.Document.Blocks.Clear();
                foreach (string variableText in Script_FoundVariables)
                {
                    //Split Strings
                    string[] parts = variableText.Split(' ');
                    string variableName = parts[parts.Length - 2];
                    string variableType = parts[parts.Length - 1];

                    //Check if Variable Type can be parsed
                    if (IsVariableValidType(variableType))
                    {
                        SetRichTextBox_Black(ScriptVariables_Textbox, variableName);
                        SetRichTextBox_Green(ScriptVariables_Textbox, " " + variableType + Environment.NewLine);
                    }
                    else
                    {
                        SetRichTextBox_Black(ScriptVariables_Textbox, variableName);
                        SetRichTextBox_Red(ScriptVariables_Textbox, " " + variableType + Environment.NewLine);
                    }
                }
            }

            //Refresh Buttons
            RefreshAllButtonStatuses();
        }

        private void SelectExcel_Button_Click(object sender, RoutedEventArgs e)
        {
            //Open File Picker
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx) | *.xlsx; *.xls";
            openFileDialog.ValidateNames = false;

            //When File Picker Is Closed Open If Loop
            if (openFileDialog.ShowDialog() == true)
            {
                //Record Data For later Then Set Title of Path Location
                Excel_CurrentFileLocation = openFileDialog.FileName;
                Textbox_SelectExcelFile.Text = System.IO.Path.GetFileName(Excel_CurrentFileLocation);

                //Create an Excel COM Object. Create a COM object for everything that is referenced
                Excel.Application excelApplication = new Excel.Application();
                Excel.Workbook inputExcel_Workbook = excelApplication.Workbooks.Open(Excel_CurrentFileLocation);

                //Arrays Start at 1 in Excel
                Excel._Worksheet inputExcel_WorksheetMain = inputExcel_Workbook.Sheets[1];
                Excel.Range inputExcel_SheetRange = inputExcel_WorksheetMain.UsedRange;

                //Excel Row Coloum Info
                int rowCount = inputExcel_SheetRange.Rows.Count;
                int colCount = inputExcel_SheetRange.Columns.Count;



                //Reset Lists
                Excel_FoundVariables = new List<string>();
                Excel_FoundNames = new List<string>();


                //Loop Allk Excel Rows and Colunms
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (i == 1)
                        {
                            //Add Each Top Cell of the first row To the List
                            if (CheckForNullCell(inputExcel_SheetRange.Cells[i, j]))
                            {
                                //Split Strings
                                string[] parts = inputExcel_SheetRange.Cells[i, j].Value2.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

                                string variableName = null;
                                string variableType = null;
                                if (parts.Length >= 2)
                                {
                                    variableName = parts[parts.Length - 2];
                                    variableType = parts[parts.Length - 1];


                                    //Add Value
                                    Excel_FoundVariables.Add(variableName + " " + variableType.ToUpper());
                                }
                            }
                        }
                        else
                        {
                            //Add Each First Colum for a name
                            if (j == 1)
                            {
                                if (CheckForNullCell(inputExcel_SheetRange.Cells[i, j]))
                                {
                                    if (inputExcel_SheetRange.Cells[i, j].value2 != "")
                                    {
                                        Excel_FoundNames.Add("No." + (i - 1) + " - " + inputExcel_SheetRange.Cells[i, j].value2);
                                    }
                                }
                            }
                        }
                    }
                }

                //Store List Of Filtered Values in Rich Text
                ExcelVariables_Textbox.Document.Blocks.Clear();
                SetRichTextBox_Black(ExcelVariables_Textbox, "" + Environment.NewLine);
                foreach (string variableText in Excel_FoundVariables)
                {
                    //Split Strings
                    string[] parts = variableText.Split(' ');
                    string variableName = parts[parts.Length - 2];
                    string variableType = parts[parts.Length - 1];

                    //Check if Variable Type can be parsed
                    if (IsVariableValidType(variableType))
                    {
                        SetRichTextBox_Black(ExcelVariables_Textbox, variableName);
                        SetRichTextBox_Green(ExcelVariables_Textbox, " " + variableType + Environment.NewLine);
                    }
                    else
                    {
                        SetRichTextBox_Black(ExcelVariables_Textbox, variableName);
                        SetRichTextBox_Red(ExcelVariables_Textbox, " " + variableType + Environment.NewLine);
                    }
                }

                //Store List Of Filtered Values in Rich Text
                ExcelData_Textbox.Document.Blocks.Clear();
                SetRichTextBox_Black(ExcelData_Textbox, "" + Environment.NewLine);
                foreach (string variableText in Excel_FoundNames)
                {
                    SetRichTextBox_Black(ExcelData_Textbox, variableText + Environment.NewLine);
                }


                //Relase Memory taken from loading the Excel file
                ReleaseExcelFile(excelApplication, inputExcel_Workbook, inputExcel_WorksheetMain, inputExcel_SheetRange);
            }

            //Refresh Buttons
            RefreshAllButtonStatuses();
        }

        private void SelectOutputLocation_Button_Click(object sender, RoutedEventArgs e)
        {
            //Open Folder Picker
            CommonOpenFileDialog openFolderDialog = new CommonOpenFileDialog();
            openFolderDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFolderDialog.IsFolderPicker = true;
            if (openFolderDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                //Record Data For later Then Set Title of Path Location
                FinalFile_OutputFileLocation = openFolderDialog.FileName;
                Textbox_SelectOutputLocation.Text = FinalFile_OutputFileLocation;
            }

            //Refresh Buttons
            RefreshAllButtonStatuses();
        }

        /////////////////////////////////////////////////////////////////

        private void Help_Button_Click(object sender, RoutedEventArgs e)
        {
            //Throw and Info Message Box
            MessageBox.Show(
                "This application is used to," + Environment.NewLine + Environment.NewLine +
                "- Create an Excel template of a Scriptable Object Script." + Environment.NewLine +
                "- The user can then fill out the data in the Excel template." + Environment.NewLine +
                "- Then reload the script and template back into the exporter and generate Scriptable Objects based on the Excel data." + Environment.NewLine +
                "- A GUID will also need to be taken from the META file of the script inside of Unity. This is visable though a text editor." + Environment.NewLine + Environment.NewLine 

                , "How To Use - Help", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void GenerateExcelTemplate_Button_Click(object sender, RoutedEventArgs e)
        {
            //Create an Excel COM Object. Create a COM object for everything that is referenced
            Excel.Application excelApplication = new Excel.Application();

            //Create Workbook
            object misValue = System.Reflection.Missing.Value;
            Excel.Workbook inputExcel_Workbook = excelApplication.Workbooks.Add(misValue);
            Excel._Worksheet inputExcel_WorksheetMain = (Excel.Worksheet)inputExcel_Workbook.Worksheets.get_Item(1);



            //Find Starting Styles
            Excel.Style customStyle_Green = inputExcel_Workbook.Styles["Good"];
            Excel.Style customStyle_Red = inputExcel_Workbook.Styles["Bad"];

            //Modify Starting Sizes
            inputExcel_WorksheetMain.Rows.RowHeight = 40;
            inputExcel_WorksheetMain.Rows[1].RowHeight = 60;
            inputExcel_WorksheetMain.Columns.ColumnWidth = 30;




            //Set First Cell Value
            inputExcel_WorksheetMain.Cells[1, 1] = "ScriptableName" + Environment.NewLine + "(STRING)";
            inputExcel_WorksheetMain.Cells[1, 1].Style = customStyle_Green;

            //Loop All variables and Chnage Cells
            for (int i = 2; i - 1 < Script_FoundVariables.Count; i++)
            {
                string[] parts = Script_FoundVariables[i - 1].Split(' ');
                string variableName = parts[parts.Length - 2];
                string variableType = parts[parts.Length - 1];


                //Apply Info
                inputExcel_WorksheetMain.Cells[1, i] = variableName + Environment.NewLine + variableType;


                if (IsVariableValidType(variableType))
                {
                    // Apply the "Good" style to the eighth row.
                    inputExcel_WorksheetMain.Cells[1, i].Style = customStyle_Green;
                }
                else
                {
                    // Apply the "Bad" style to the eighth row.
                    inputExcel_WorksheetMain.Cells[1, i].Style = customStyle_Red;
                }
            }

            try
            {
                //Save Excel File
                inputExcel_Workbook.SaveAs(FinalFile_OutputFileLocation + "\\" + System.IO.Path.GetFileNameWithoutExtension(Script_CurrentFileLocation) + " - Template", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                //Relase Memory taken from loading the Excel file
                ReleaseExcelFile(excelApplication, inputExcel_Workbook, inputExcel_WorksheetMain, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to save Excel File, is the file currently open?" + Environment.NewLine + ex.Message, "Unable to Preform a Task", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FilterMatches_Button_Click(object sender, RoutedEventArgs e)
        {
            //Reset Tuple and get all matches
            Data_MatchedVariableSet_List = new List<CompletedDataset>();
            for (int i = 0; i < Script_FoundVariables.Count; i++)
            {
                //Split Strings
                string[] parts = Script_FoundVariables[i].Split(' ');
                string variableName = parts[parts.Length - 2];
                string variableType = parts[parts.Length - 1];

                for (int j = 0; j < Excel_FoundVariables.Count; j++)
                {
                    //Split Strings
                    string[] parts_2 = Excel_FoundVariables[j].Split(' ');
                    string variableName_2 = parts_2[parts_2.Length - 2];
                    string variableType_2 = parts_2[parts_2.Length - 1];

                    //Console.Out.WriteLine(variableName + " / " + Excel_FoundVariables[j]);
                    if (variableName == variableName_2)
                    {
                        Data_MatchedVariableSet_List.Add(new CompletedDataset(variableName, variableType, i, j));
                        break;
                    }
                }
            }

            //Reset 
            ScriptVariables_Textbox.Document.Blocks.Clear();
            ExcelVariables_Textbox.Document.Blocks.Clear();
            for (int i = 0; i < Data_MatchedVariableSet_List.Count; i++)
            {

                //Check if Variable Type can be parsed
                if (IsVariableValidType(Data_MatchedVariableSet_List[i].variableType))
                {
                    SetRichTextBox_Black(ScriptVariables_Textbox, "Line: " + Data_MatchedVariableSet_List[i].scriptVariableLineNo + " " + Data_MatchedVariableSet_List[i].variableName);
                    SetRichTextBox_Green(ScriptVariables_Textbox, " " + Data_MatchedVariableSet_List[i].variableType + Environment.NewLine);

                    SetRichTextBox_Black(ExcelVariables_Textbox, "Column: " + Data_MatchedVariableSet_List[i].excelVariableLineNo + " " + Data_MatchedVariableSet_List[i].variableName);
                    SetRichTextBox_Green(ExcelVariables_Textbox, " " + Data_MatchedVariableSet_List[i].variableType + Environment.NewLine);
                }
                else
                {
                    SetRichTextBox_Black(ScriptVariables_Textbox, "Line: " + Data_MatchedVariableSet_List[i].scriptVariableLineNo + " " + Data_MatchedVariableSet_List[i].variableName);
                    SetRichTextBox_Red(ScriptVariables_Textbox, " " + Data_MatchedVariableSet_List[i].variableType + Environment.NewLine);

                    SetRichTextBox_Black(ExcelVariables_Textbox, "Column: " + Data_MatchedVariableSet_List[i].excelVariableLineNo + " " + Data_MatchedVariableSet_List[i].variableName);
                    SetRichTextBox_Red(ExcelVariables_Textbox, " " + Data_MatchedVariableSet_List[i].variableType + Environment.NewLine);
                }
            }

            RefreshAllButtonStatuses();
        }
       
        private void ConvertToScriptable_Button_Click(object sender, RoutedEventArgs e)
        {

            // Specify the directory you want to manipulate.
            string path = FinalFile_OutputFileLocation + "\\" + System.IO.Path.GetFileNameWithoutExtension(Script_CurrentFileLocation) + " - Scriptable Exports";

            try
            {
                //Check if file already exists. If so, delete it.     
                if (Directory.Exists(path))
                {
                    Directory.Delete(path);
                }

                //Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to create Directory." + Environment.NewLine + ex.Message, "Unable to Preform a Task", MessageBoxButton.OK, MessageBoxImage.Error);
            }


            //Create an Excel COM Object. Create a COM object for everything that is referenced
            Excel.Application excelApplication = new Excel.Application();
            Excel.Workbook inputExcel_Workbook = excelApplication.Workbooks.Open(Excel_CurrentFileLocation);

            //Arrays Start at 1 in Excel
            Excel._Worksheet inputExcel_WorksheetMain = inputExcel_Workbook.Sheets[1];
            Excel.Range inputExcel_SheetRange = inputExcel_WorksheetMain.UsedRange;

            //Excel Row Coloum Info
            int rowCount = inputExcel_SheetRange.Rows.Count;
            int colCount = inputExcel_SheetRange.Columns.Count;







            //All Rows = files
            // All Colums Match the Class
            //Enter All Data from class








            //Loop All Excel Rows and Colunms
            for (int i = 2; i <= rowCount; i++)
            {
                List<byte[]> fileContents_Variables = new List<byte[]>();
                List<string> incomingData_List = CreateScriptableDataList(inputExcel_SheetRange, i, colCount);

                for (int j = 0; j < Data_MatchedVariableSet_List.Count; j++)
                {
                    //Check if Variable Type can be parsed
                    if (IsVariableValidType(Data_MatchedVariableSet_List[j].variableType))
                    {
                        //Skip the first value it does not exist in the data orgnially
                        if (incomingData_List[j] != "ScriptableName")
                        {
                            fileContents_Variables.Add(new UTF8Encoding(true).GetBytes("  " + Data_MatchedVariableSet_List[j].variableName + ": " + incomingData_List[j] + Environment.NewLine));
                        }
                        else
                        {
                            fileContents_Variables.Add(new UTF8Encoding(true).GetBytes("  " + Data_MatchedVariableSet_List[j].variableName + ": " + Environment.NewLine));
                        }
                    }
                    else
                    {
                        fileContents_Variables.Add(new UTF8Encoding(true).GetBytes("  " + Data_MatchedVariableSet_List[j].variableName + ": {fileID: 0}" + Environment.NewLine));
                    }
                }



                //Create the file
                CreateScriptableAsset(fileContents_Variables, incomingData_List[0], path);
            }








            //Relase Memory taken from loading the Excel file
            ReleaseExcelFile(excelApplication, inputExcel_Workbook, inputExcel_WorksheetMain, inputExcel_SheetRange);
        }

        /////////////////////////////////////////////////////////////////

        private void RefreshAllButtonStatuses()
        {
            if (Script_FoundVariables != null)
            {
                if (Script_FoundVariables.Count > 0)
                {
                    GenerateExcelTemplate_Button.IsEnabled = true;

                    if (Excel_FoundVariables != null)
                    {
                        if (Excel_FoundVariables.Count > 0)
                        {
                            FilterMatches_Button.IsEnabled = true;
                        }
                        else
                        {
                            FilterMatches_Button.IsEnabled = false;
                        }
                    }
                    else
                    {
                        FilterMatches_Button.IsEnabled = false;
                    }
                }
                else
                {
                    GenerateExcelTemplate_Button.IsEnabled = false;
                    FilterMatches_Button.IsEnabled = false;
                }
            }
            else
            {
                GenerateExcelTemplate_Button.IsEnabled = false;
                FilterMatches_Button.IsEnabled = false;
            }


            if (Data_MatchedVariableSet_List != null)
            {
                if (Data_MatchedVariableSet_List.Count > 0 && InputGUID_Textbox.Text != "")
                {
                    ConvertToScriptable_Button.IsEnabled = true;
                }
                else
                {
                    ConvertToScriptable_Button.IsEnabled = false;
                }
            }
            else
            {
                ConvertToScriptable_Button.IsEnabled = false;
            }
        }

        private string WildCardToRegular(string value)
        {
            return "^" + Regex.Escape(value).Replace("\\*", ".*") + "$";
        }

        private void ReleaseExcelFile(Excel.Application excelApplication, Excel.Workbook inputExcel_Workbook, Excel._Worksheet inputExcel_WorksheetMain, Excel.Range inputExcel_SheetRange)
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            if (inputExcel_SheetRange != null)
            {
                Marshal.ReleaseComObject(inputExcel_SheetRange);
            }

            Marshal.ReleaseComObject(inputExcel_WorksheetMain);

            //close and release
            inputExcel_Workbook.Close();
            Marshal.ReleaseComObject(inputExcel_Workbook);

            //quit and release
            excelApplication.Quit();
            Marshal.ReleaseComObject(excelApplication);
        }

        private bool IsVariableValidType(string variableType)
        {
            //Filter Correct Data Types
            switch (variableType.ToUpper())
            {
                case "(STRING)":
                case "(BOOL)":
                case "(INT)":
                case "(FLOAT)":
                case "(COLOR)":
                    return true;

                default:
                    return false;
            }
        }

        private List<string> CreateScriptableDataList(Excel.Range inputExcel_SheetRange, int i, int colCount)
        {
            //Data List
            List<string> dataCollected_List = new List<string>();


            for (int j = 1; j <= colCount; j++)
            {
                //Add Each Top Cell of the first row To the List
                if (CheckForNullCell(inputExcel_SheetRange.Cells[i, j]))
                {
                    dataCollected_List.Add(inputExcel_SheetRange.Cells[i, j].Value2.ToString());
                }
                else
                {
                    dataCollected_List.Add("");
                }
            }

            return dataCollected_List;
        }
        
        private void CreateScriptableAsset(List<byte[]> fileContents_Variables, string fileName, string filePath)
        {
            try
            {
                //Skip Missing Names
                if (fileName == "")
                {
                    return;
                }

                //Check if file already exists. If so, delete it.     
                if (File.Exists(fileName + ".asset"))
                {
                    File.Delete(fileName + ".asset");
                }

                // Create the file, or overwrite if the file exists.
                using (FileStream filestream = File.Create(filePath + "\\" + fileName + ".asset"))
                {
                    //File Setup Headings
                    byte[] fileContents_Header = new UTF8Encoding(true).GetBytes(
                        "%YAML 1.1" + Environment.NewLine +
                        "%TAG !u! tag:unity3d.com,2011:" + Environment.NewLine +
                        "--- !u!114 &11400000" + Environment.NewLine +
                        "MonoBehaviour:" + Environment.NewLine +
                        "  m_ObjectHideFlags: 0" + Environment.NewLine +
                        "  m_CorrespondingSourceObject: {fileID: 0}" + Environment.NewLine +
                        "  m_PrefabInstance: {fileID: 0}" + Environment.NewLine +
                        "  m_PrefabAsset: {fileID: 0}" + Environment.NewLine +
                        "  m_GameObject: {fileID: 0}" + Environment.NewLine +
                        "  m_Enabled: 1" + Environment.NewLine +
                        "  m_EditorHideFlags: 0" + Environment.NewLine +
                        "  m_Script: {fileID: 11500000, guid: " + InputGUID_Textbox.Text + ", type: 3}" + Environment.NewLine
                        );
                    byte[] fileContents_Name = new UTF8Encoding(true).GetBytes("  m_Name: " + fileName + Environment.NewLine);
                    byte[] fileContents_EditorID = new UTF8Encoding(true).GetBytes("  m_EditorClassIdentifier: " + Environment.NewLine);


                    //Write to file
                    filestream.Write(fileContents_Header, 0, fileContents_Header.Length);
                    filestream.Write(fileContents_Name, 0, fileContents_Name.Length);
                    filestream.Write(fileContents_EditorID, 0, fileContents_EditorID.Length);
                    foreach (byte[] byteArray in fileContents_Variables)
                    {
                        //Console.Out.WriteLine();
                        filestream.Write(byteArray, 0, byteArray.Length);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to create files." + Environment.NewLine + ex.Message, "Unable to Preform a Task", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CheckForNullCell(Excel.Range cell)
        {
            if (cell != null && cell.Value2 != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void SetRichTextBox_Black(RichTextBox richTextbox, string textContent)
        {
            TextRange rangeOfText = new TextRange(richTextbox.Document.ContentEnd, richTextbox.Document.ContentEnd);
            rangeOfText.Text = textContent;
            rangeOfText.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.Black);
            rangeOfText.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
        }

        private void SetRichTextBox_Green(RichTextBox richTextbox, string textContent)
        {
            TextRange rangeOfText = new TextRange(richTextbox.Document.ContentEnd, richTextbox.Document.ContentEnd);
            rangeOfText.Text = textContent;
            rangeOfText.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.Green);
            rangeOfText.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
        }

        private void SetRichTextBox_Red(RichTextBox richTextbox, string textContent)
        {
            TextRange rangeOfText = new TextRange(richTextbox.Document.ContentEnd, richTextbox.Document.ContentEnd);
            rangeOfText.Text = textContent;
            rangeOfText.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.Red);
            rangeOfText.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
        }

        private void InputGUID_Textbox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshAllButtonStatuses();
        }

        /////////////////////////////////////////////////////////////////
    }

    /////////////////////////////////////////////////////////////////

    public class CompletedDataset
    {
        public string variableName;
        public string variableType;
        public int scriptVariableLineNo;
        public int excelVariableLineNo;

        /////////////////////////////////////////////////////////////////

        public CompletedDataset(string variableName, string variableType, int scriptVariableLineNo, int excelVariableLineNo)
        {
            this.variableName = variableName;
            this.variableType = variableType;
            this.scriptVariableLineNo = scriptVariableLineNo;
            this.excelVariableLineNo = excelVariableLineNo;
        }
    }

    /////////////////////////////////////////////////////////////////

}
