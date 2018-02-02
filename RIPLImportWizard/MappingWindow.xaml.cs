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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Reflection;

//using System.Windows.Forms;

namespace RIPLImportWizard
{

    /// <summary>
    /// Interaction logic for MappingWindow.xaml
    /// </summary>
    public partial class MappingWindow : Window
    {
        public AttributeMapping win2 = new AttributeMapping();

        public ObservableCollection<string> Types { get; set; }
        public ObservableCollection<InputModelColumns> InputModelCol { get; set; }
        public ObservableCollection<SourceColumns> SourceCol { get; set; }
        #region SQL Connection
        public string Sql()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = "localhost";
            builder.UserID = "zach.hine";              // update me
            builder.IntegratedSecurity = true;
            builder.Password = "password2";      // update me
            builder.InitialCatalog = "master";
            return builder.ConnectionString;
        }
        public SqlCommand OpenConnection()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            SqlConnection connection = new SqlConnection(Sql());
            cmd.Connection = connection;
            connection.Open();
            return cmd;
        }
        #endregion
        private string transform = string.Empty;
        private List<Component> riplcomp = new List<Component>();
        private List<SourceComp> sourcecomp = new List<SourceComp>();
        private List<TotalComp> totalcomp = new List<TotalComp>();
        public List<AttMapList> myList = new List<AttMapList>();
        
        public MappingWindow()
        {
            InitializeComponent();
        }

        T MapToClass<T>(SqlDataReader reader) where T : class
        {
            T returnedObject = Activator.CreateInstance<T>();
            PropertyInfo[] modelProperties = returnedObject.GetType().GetProperties();
            for (int i = 0; i < modelProperties.Length; i++)
            {
                MappingAttribute[] attributes = modelProperties[i].GetCustomAttributes<MappingAttribute>(true).ToArray();

                if (attributes.Length > 0 && attributes[0].ColumnName != null)
                    modelProperties[i].SetValue(returnedObject, Convert.ChangeType(reader[attributes[0].ColumnName], modelProperties[i].PropertyType), null);
            }
            return returnedObject;
        }
        //public SourceSheets GetAllWorkSheetNames(string fileName)
        //{
        //    SourceSheets sheets = new SourceSheets;
        //    using (SpreadsheetDocument document =
        //       SpreadsheetDocument.Open(fileName, false))
        //    {
        //        WorkbookPart wbPart = document.WorkbookPart;
        //        sheets = wbPart.Workbook.Sheets;
        //        document.Close();
        //    }
        //    return sheets
        //}

        public static Sheets GetAllWorksheets(string fileName)
        {
            Sheets theSheets = null;

            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                theSheets = wbPart.Workbook.Sheets;
                document.Close();
            }

            return theSheets;
        }
        private DataTable GetInputModels()
        {
            DataTable dtInput = new DataTable();
            dtInput.Columns.Add("Input Model", typeof(string));


            using (SqlConnection connection = new SqlConnection(Sql()))
            {

                connection.Open();

                string sqlString = "SELECT [Model_Name] FROM [Import_78].[dbo].[Model] ORDER BY [Model_Name]";

                using (SqlCommand command = new SqlCommand(sqlString, connection))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    dtInput.Load(reader);
                }
            }
            return dtInput;
        }
        public List<RIPLVariables> GetVariables(string model)
        {
            List<RIPLVariables> lstVariables = new List<RIPLVariables>();
            try
            {
                using (SqlCommand cmd = OpenConnection())
                {
                    cmd.CommandText = String.Format("SELECT [Var_Description],[Var_ID],[Var_Type] From [Import_78].[dbo].[Variables] WHERE [Import_78].[dbo].[Variables].[Var_ID] IN (SELECT [Import_78].[dbo].[Model_Link].[Var_ID] FROM [Import_78].[dbo].[Model_Link] WHERE [Import_78].[dbo].[Model_Link].[Model_ID] IN (SELECT [Import_78].[dbo].[Model].[Model_ID] FROM [Import_78].[dbo].[Model] WHERE [Model_Name] = '{0}'))", model);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            RIPLVariables item = new RIPLVariables();
                            item.varname = reader["Var_Description"].ToString();
                            item.varID = Convert.ToInt32(reader["Var_ID"]);
                            item.vartype = Convert.ToInt32(reader["Var_Type"]);

                            lstVariables.Add(item);
                        }

                    }
                }
                return lstVariables;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public DataTable GetColumns(string fileName, string sheet)
        {
            DataTable dtColumns = new DataTable();
            dtColumns.Columns.Add("Columns", typeof(string));

            int num = 1;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(fileName.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);



            foreach (Excel.Worksheet exSheet in excelBook.Sheets)
            {
                if (exSheet.Name == sheet)
                {
                    num = exSheet.Index;

                }

            }

            Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(num);
            Excel.Range excelRange = excelSheet.UsedRange;
            int colCnt = 0;

            List<string> list = new List<string>();
            for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
            {
                string strColumn = "";
                strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                list.Add(strColumn);
            }
            foreach (var item in list)
            {
                DataRow row = dtColumns.NewRow();
                row["Columns"] = item;
                dtColumns.Rows.Add(row);
            }

            excelBook.Close(true, null, null);
            excelApp.Quit();

            return dtColumns;
        }
        public int GetVarType(string variable)
        {
            int vartype = 0;
            using (SqlConnection connection = new SqlConnection(Sql()))
            {
                connection.Open();
                string sqlString = String.Format("SELECT [Var_Type] FROM [Import_78].[dbo].Variables WHERE [Var_Description] = '{0}'", variable);
                using (SqlCommand command = new SqlCommand(sqlString, connection))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    vartype.Equals(reader);
                }
            }
            return vartype;
        }
        private string GetVarTypeString(int vartype)
        {
            string varstring = "hello";
            if (vartype == 16)
            {
                varstring = "List";
            }
            else if (vartype == 19)
            {
                varstring = "Date";
            }
            else if (vartype == 18)
            {
                varstring = "Int";
            }
            else if (vartype == 17)
            {
                varstring = "String";
            }
            return varstring;
        }
        public DataTable GetAttributes(string variable)
        {
            DataTable dtAttributes = new DataTable();

            using (SqlConnection connection = new SqlConnection(Sql()))
            {

                connection.Open();

                string sqlString = String.Format("SELECT[Description] FROM[Import_78].[dbo].[Attributes]  WHERE[Import_78].[dbo].[Attributes].Att_ID IN (SELECT[Import_78].[dbo].[Att_Link].Att_ID FROM[Import_78].[dbo].[Att_Link] WHERE[Import_78].[dbo].[Att_Link].Var_ID IN (Select[Import_78].[dbo].[Variables].Var_ID FROM[Import_78].[dbo].[Variables] WHERE[Import_78].[dbo].[Variables].Var_Description = '{0}'))", variable);

                using (SqlCommand command = new SqlCommand(sqlString, connection))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    dtAttributes.Load(reader);
                }
            }

            return dtAttributes;
        }
        public class CheckBoxListItem
        {
            public bool Checked { get; set; }
            public string Text { get; set; }

            public CheckBoxListItem(bool ch, string text)
            {
                Checked = ch;
                Text = text;
            }
        }
        public DataTable GetSourceComps(string fileName, string sheet, string column)
        {
            DataTable dtComps = new DataTable();
            dtComps.Columns.Add("Component", typeof(string));
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(fileName.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Sheets[sheet];
            Excel.Range excelRange = excelSheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int columnCount = excelRange.Columns.Count;
            Excel.Range result = null;
            string Address = null;
            Excel.Range columns = excelSheet.Rows[1] as Excel.Range;
            //foreach (Excel.Range c in columns.Cells)
            //{
            //    if(c.Value == column)
            //    {

            //    }
            //}
            List<string> columnValue = new List<string>();

            result = columns.Find(What: column, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlWhole, SearchOrder: Excel.XlSearchOrder.xlByColumns);
            Excel.Range cRng = null;
            if (result != null)
            {
                Address = result.Address;

                do
                {
                    for (int i = 2; i <= rowCount; i++)
                    {
                        cRng = excelSheet.Cells[i, result.Column] as Excel.Range;
                        if (cRng.Value != null)
                        {
                            columnValue.Add(cRng.Value.ToString());
                        }
                    }
                } while (result == null);

            }

            List<string> list = columnValue.Distinct().ToList();
            foreach (var item in list)
            {
                dtComps.Rows.Add(item.ToString());
            }
            return dtComps;
        }
        public DataTable GetRIPLComps()
        {
            DataTable dtComponent = new DataTable();

            using (SqlConnection connection = new SqlConnection(Sql()))
            {

                connection.Open();
                string sqlString = "SELECT TOP (7) [Comp_Name] FROM [Import_78].[dbo].[Component]";

                using (SqlCommand command = new SqlCommand(sqlString, connection))
                {
                    SqlDataReader reader = command.ExecuteReader();
                    dtComponent.Load(reader);
                }
            }

            return dtComponent;
        }
        private void New_Sheet_Columns(object sender, EventArgs e)
        {

            int num = 0;
            num = listSheets.SelectedIndex + 1;
            string selectedFile = ((MainWindow)Application.Current.MainWindow).filePath.Text;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(selectedFile.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(num); ;
            Excel.Range excelRange = excelSheet.UsedRange;

            string strCellData = "";
            double douCellData;
            int rowCnt = 0;
            int colCnt = 0;

            DataTable dt = new DataTable();
            for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
            {
                string strColumn = "";
                strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                dt.Columns.Add(strColumn, typeof(string));
            }

            for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
            {
                string strData = "";
                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    try
                    {
                        strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                        strData += strCellData + "|";
                    }
                    catch (Exception ex)
                    {
                        douCellData = (excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                        strData += douCellData.ToString() + "|";
                    }
                }
                strData = strData.Remove(strData.Length - 1, 1);
                dt.Rows.Add(strData.Split('|'));
            }

            dtGrid.ItemsSource = dt.DefaultView;

            excelBook.Close(true, null, null);
            excelApp.Quit();
        }
        private void Load_Models(object sender, EventArgs e)
        {

            List<InputModels> inputModels = new List<InputModels>();

            DataTable dtInputModels = GetInputModels();

            foreach (DataRow row in dtInputModels.Rows)
            {
                inputModels.Add(new InputModels() { InputModel = row[1].ToString() });
            }

            Input.ItemsSource = inputModels;

            CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(Input.ItemsSource);
            view.Filter = UserFilter;
        }
        public List<string> GetWorksheets(string fileName)
        {
            int num = 1;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(fileName.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(num);
            Excel.Range excelRange = excelSheet.UsedRange;
            List<string> list = new List<string>();
            foreach (Excel.Worksheet exSheet in excelBook.Sheets)
            {
                list.Add(exSheet.Name);
            }
            return list;
        }
        public void MappingWindowLoaded(object sender, RoutedEventArgs e)
        {

            string selectedFile = ((MainWindow)Application.Current.MainWindow).filePath.Text;
            var results = GetAllWorksheets(selectedFile);

           

            foreach (Sheet item in results)
            {
                listSheets.Items.Add(item.Name);
            }

            string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + selectedFile + ";Extended Properties=Excel 12.0;");

            
        }
        private bool UserFilter(object item)
        {
            if (String.IsNullOrEmpty(ModelFilter.Text))
                return true;
            else
                return ((item as InputModels).InputModel.IndexOf(ModelFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0);
        }
        private void Selector_OnSelectionChanged(object sender, EventArgs e)
        {
            string model;

            if (Input.SelectedItems.Count > 0)
            {
                model = Input.SelectedItem.ToString();
                DataTable dtModelColumns = GetModelColumns(model);
                SourceColumns.Items.Clear();
                foreach (DataRow row in dtModelColumns.Rows)
                {
                    SourceColumns.Items.Add(row[0].ToString());
                }
            }



        }
        private void ModelFilter_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            CollectionViewSource.GetDefaultView(Input.ItemsSource).Refresh();
        }

        public int GetModelPositionType(string model)
        {
            try
            {
                int positionType = 0;
                using (SqlCommand cmd = OpenConnection())
                {
                    cmd.CommandText = String.Format("Select [Positioning] From [Import_78].[dbo].[Model] Where Model_Name = '{0}'", model);
                    using(SqlDataReader reader = cmd.ExecuteReader())
                    {

                        while (reader.Read())
                        {
                            positionType = Convert.ToInt32(reader["Positioning"]);
                            
                        }
                    }
                }
                return positionType;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public string GetStationingNames(int varID)
        {
            try
            {
                string station ="";
                using (SqlCommand cmd = OpenConnection())
                {
                    cmd.CommandText = String.Format("Select [Var_Description] From [Import_78].[dbo].[Variables] Where Var_ID = '{0}'", varID);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {

                        while (reader.Read())
                        {
                            station = Convert.ToString(reader["Var_Description"]);

                        }
                    }
                }
                return station;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        //SQL Builde

        private void Load_Mapping(object sender, EventArgs e)
        {
            BuildReferenceList();
            List<string> inputList = new List<string>();
            foreach (var item in Input.SelectedItems)
            {
                inputList.Add((string)item.ToString());
            }
            Selected_Models.ItemsSource = inputList;

            string selectedFile = ((MainWindow)Application.Current.MainWindow).filePath.Text;

            List<string> worksheets = new List<string>();
            worksheets = GetWorksheets(selectedFile);
            cbSourceSheet.ItemsSource = worksheets;
        }
        private void BuildReferenceList()
        {
            cbReferences.Items.Clear();
            try
            {
                List<Reference> refs = GetReferences();
                #region Add Feature Classes to combobox
                int idx = 0;
                cbReferences.Items.Add("");
                for (int i = 0; i < refs.Count; i++)
                {
                    cbReferences.Items.Add(refs[i].Ref_Name);
                }

                cbReferences.SelectedIndex = idx > 0 ? idx : 0;
                
                #endregion Feature Classes to combobox

            }
            catch (Exception e)
            {
                throw e;
            }


        }
        public List<Reference> GetReferences()
        {
            List<Reference> refs = new List<Reference>();
            try
            {
                using (SqlCommand cmd = OpenConnection())
                {
                    cmd.CommandText = "SELECT [Ref_Name],[Ref_ID] FROM [Import_78].[dbo].[Ref_Def]";
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Reference item = new Reference();
                            item.Ref_Name = reader["Ref_Name"].ToString();
                            item.Ref_ID = Convert.ToInt32(reader["Ref_ID"]);

                            refs.Add(item);
                        }
                    }
                }
                return refs;

            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public void Selected_Models_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            int model_ID;
            string selModel = Selected_Models.SelectedItem.ToString();
            using (SqlConnection connection = new SqlConnection(Sql()))
            {

                connection.Open();
                string sqlString = String.Format("SELECT Model_ID FROM [Import_78].[dbo].[Model] WHERE [Model_Name] = '{0}'", selModel);
                SqlCommand cmd = new SqlCommand(sqlString, connection);
                model_ID = Convert.ToInt32(cmd.ExecuteScalar());
                SelectedInputModels a = new SelectedInputModels(selModel, model_ID);

            }
        }
        //ComboBox Selection Changes
        public void cbReferences_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var References = sender as ComboBox;
            string reference = cbReferences.SelectedItem as string;
            MessageBox.Show(myList.Count.ToString());

        }

        private void cbSourceSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var SourceSheet = sender as ComboBox;
            string selectedSheet = SourceSheet.SelectedItem as string;
            MessageBox.Show(selectedSheet);

            List<RIPLVariables> lstVariables = new List<RIPLVariables>();
            lstVariables = GetVariables(Selected_Models.SelectedItem.ToString());


            int z = GetModelPositionType(Selected_Models.SelectedItem.ToString());

            DataRow toInsertComp = dt.NewRow();
            toInsertComp[0] = "Component";
            toInsertComp[1] = 16;
            dt.Rows.InsertAt(toInsertComp, 0);

           
            if (z == 5)
            {
                DataRow toInsertStation1 = dt.NewRow();
                toInsertStation1[0] = GetStationingNames(1);
                toInsertStation1[1] = 18;
                dt.Rows.InsertAt(toInsertStation1, 1);
                DataRow toInsertStation2 = dt.NewRow();
                toInsertStation2[0] = GetStationingNames(4);
                toInsertStation2[1] = 18;
                dt.Rows.InsertAt(toInsertStation2, 2);
            }
            else if (z == 2)
            {
                DataRow toInsertStation = dt.NewRow();
                toInsertStation[0] = GetStationingNames(2);
                toInsertStation[1] = 18;
                dt.Rows.InsertAt(toInsertStation, 1);
            }
            VarMapping.DataContext = dt.DefaultView;
            List<InputModelColumns> inputmodelcol = new List<InputModelColumns>();

            foreach (DataRow row in dt.Rows)
            {
                string tcol;
                int x = int.Parse(row[1].ToString());
                if (x == 16)
                {
                    tcol = "Use List Transformation";
                }
                else
                {
                    tcol = "Use Values From Source Field";
                }

                inputmodelcol.Add(new InputModelColumns { ModelCol = row[0].ToString(), VarType = GetVarTypeString(x), TransformCol = tcol });

            }
            VarMapping.ItemsSource = inputmodelcol;

            string selectedFile = ((MainWindow)Application.Current.MainWindow).filePath.Text;

            DataTable dtSourceColumns = new DataTable();
            dtSourceColumns = GetColumns(selectedFile, selectedSheet);

            List<SourceColumns> _sourcecol = new List<SourceColumns>();
            foreach (DataRow row in dtSourceColumns.Rows)
            {
                _sourcecol.Add(new SourceColumns { SourceCol = row[0].ToString() });
            }


            cbMapVars.ItemsSource = _sourcecol;
        }
        //Mapping Components
        private void MapComponents_Loaded(object sender, RoutedEventArgs e)
        {
            DataTable dt = new DataTable();
            dt = GetRIPLComps();


            foreach (DataRow row in dt.Rows)
            {

                riplcomp.Add(new Component { component = row[0].ToString() });
                totalcomp.Add(new TotalComp { ripl = row[0].ToString() });
            }
            RIPLComp.ItemsSource = riplcomp;

            string selectedFile = ((MainWindow)Application.Current.MainWindow).filePath.Text;
            string selectedSheet = cbSourceSheet.SelectedValue.ToString();
            string componentColumn = "Component";

            DataTable dtSourceComponent = new DataTable();
            dtSourceComponent = GetSourceComps(selectedFile, selectedSheet, componentColumn);

            foreach (DataRow row in dtSourceComponent.Rows)
            {
                sourcecomp.Add(new SourceComp { sourcecompvar = row[0].ToString() });
                totalcomp.Add(new TotalComp { xcel = row[0].ToString() });
            }

            Comps.ItemsSource = sourcecomp;
        }
        public void AutoMap_Click(object sender, RoutedEventArgs e)
        {
            //string comp1 = null;
            ////W.ItemsSource = riplcomp;

            ////Test.ItemsSource = riplcomp;
            ////Test2.ItemsSource = riplcomp;
            //List<CompIndex> index = new List<CompIndex>();
            //for (int i = 0; i < Comps.Items.Count; i++)
            //{
            //    DataGridRow Row = (DataGridRow)Comps.ItemContainerGenerator.ContainerFromIndex(i);
            //    DataGridCell RowAndColumn = (DataGridCell)Comps.Columns[0].GetCellContent(Row).Parent;
            //    comp1 = ((TextBlock)RowAndColumn.Content).Text;


            //    foreach (var item in riplcomp)
            //    {


            //        if (item.ToString() == comp1)
            //        {

            //            var cs = new FindIndex(comp1);
            //            int index2 = riplcomp.FindIndex(cs.Equal);
            //            index.Add(new CompIndex { compindex = riplcomp.FindIndex(cs.Equal) });
            //            //Test2.TextBinding = new Binding(comp1);
            //            //RIPLComp.TextBinding = new Binding(Path = index2);

            //            //W.SelectedIndex = index2;
            //            //MessageBox.Show("next");

            //            //MessageBox.Show()
            //            //index = riplcomp.FindIndex(cs.Equal);




            //            //RIPLComp.DisplayIndex = index;

            //            //Comps.Columns[1].DisplayIndex = index;
            //        }
            //    }
            //}
        }

        //INSERT INTO DATABASE - SQLBULKQUERY CODE
        private void WriteToDatabase(int Model_ID, DataTable import_data)
        {
            string destTable = String.Format("dbo.m_{0}", Model_ID);
            using (SqlConnection connection = new SqlConnection(Sql()))
            {
                connection.Open();
                SqlCommand commandRowCount = new SqlCommand("SELECT COUNT(*) FROM " + destTable, connection);
                long countStart = System.Convert.ToInt32(
                    commandRowCount.ExecuteScalar());
                Console.WriteLine("Starting row count = {0}", countStart);
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {

                    bulkCopy.DestinationTableName = destTable;
                    try
                    {
                        // Write from the source to the destination.
                        bulkCopy.WriteToServer(import_data);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                long countEnd = System.Convert.ToInt32(
                   commandRowCount.ExecuteScalar());
                Console.WriteLine("Ending row count = {0}", countEnd);
                Console.WriteLine("{0} rows were added.", countEnd - countStart);
                Console.WriteLine("Press Enter to finish.");
                Console.ReadLine();
            }

        }

        //Attribute Mapping
        public AttributeMapping win2 = new AttributeMapping();
        private void AttMapp_Click(object sender, RoutedEventArgs e)
        {
            
            win2.Show();
        }
        
        public void PassList(List<AttMapList> myList)
        {
            this.myList = myList;
            MessageBox.Show(myList.Count.ToString());
        }

        private void VarMapping_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        public class RIPLVariables
        {
            public string varname { get; set; }
            public override string ToString()
            {
                return this.varname;
            }
            public int varID { get; set; }
            public int vartype { get; set; }
        }

    }

    public class SourceColumns
    {
        public string SourceCol { get; set; }

        public override string ToString()
        {
            return this.SourceCol;
        }
    }
    public class InputModelColumns
    {
        public string ModelCol { get; set; }
        public string VarType { get; set; }
        public string SourceCol2 { get; set; }
        public string TransformCol { get; set; }
    }
    public class Transform
    {
        private string _transform;

        public string _Transform
        {
            get
            {
                return this._transform;
            }

            set
            {
                this._transform = value;
            }
        }

        public override string ToString()
        {
            return this._transform;
        }
    }
    public class InputModels
    {
        public string InputModel { get; set; }

        public override string ToString()
        {
            return this.InputModel;
        }
    }
    public class Component
    {

        public string component { get; set; }
        // public string compID { get; set; }

        public override string ToString()
        {
            return this.component;
        }

        //public string compID { get; set; 

    }
    public class SourceComp
    {
        public string sourcecompvar { get; set; }
        public override string ToString()
        {
            return this.sourcecompvar;
        }
    }
    public class FindIndex
    {
        String _comp;
        public FindIndex(string comp)
        {
            _comp = comp;
        }
        public bool Equal(Component com)
        {
            return com.component.Equals(_comp, StringComparison.InvariantCultureIgnoreCase);
        }
    }
    public class CompIndex
    {
        public int compindex { get; set; }
    }
    public class TotalComp
    {
        public string ripl { get; set; }
        public string xcel { get; set; }
    }
    public class SelectedInputModels
    {
        public string Model_Name;
        public int Model_ID;
        public SelectedInputModels(string Model_Name, int Model_ID)
        {
            this.Model_Name = Model_Name;
            this.Model_ID = Model_ID;
        }
        
    }
    [AttributeUsage(AttributeTargets.Property, Inherited = true)]
    [Serializable]
    public class MappingAttribute : Attribute
    {
        public string ColumnName = null;
    }
    public class Reference
    {
        [Mapping(ColumnName = "Reference Name")]
        public string Ref_Name { get; set; }
        [Mapping(ColumnName = "Reference ID")]
        public int Ref_ID { get; set; }

        //public Reference(string RefName, int Ref_ID)
        //{
        //    this.Ref_Name = Ref_Name;
        //    this.Ref_ID = Ref_ID;
        //}
        
    }
    public class Model : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private object _xcelvar;
        public object XcelVar
        {
            get { return _xcelvar; }
            set
            {
                _xcelvar = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("XcelVar"));
            }
        }

        private object _selectedVar;
        public object SelectedVar
        {
            get { return _selectedVar; }
            set
            {
                _selectedVar = value;
                this.OnPropertyChanged(new PropertyChangedEventArgs("SelectedVar"));
            }
        }
        public virtual void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, e);
        }
    }

}

