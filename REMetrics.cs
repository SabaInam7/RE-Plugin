

using System;
using EA;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Net.NetworkInformation;

namespace CodeQuality
{
    public class REMetrics
    {
        // define menu constants
        const string menuHeader = "-&OOP Metrics";
        const string menuAss = "&Association Density";
        const string menuIssues = "&TestIssues";
        const string menuLCOM = "&LCOM";
        const string menuCoupling = " &Coupling";
        const string menuUniqueElement = "&UniqueElements";
        OleDbConnection cnn;
        OleDbCommand cmd;

        private bool shouldWeSayHello = true;

        public String EA_Connect(EA.Repository Repository)
        {
            //No special processing required.
            return "a string";
        }

        public object EA_GetMenuItems(EA.Repository Repository, string Location, string MenuName)
        {
            switch (MenuName)
            {
                // defines the top level menu option
                case "":
                    return menuHeader;
                // defines the submenu options
                case menuHeader:
                    string[] subMenus = { menuAss, menuIssues , menuLCOM, menuCoupling, menuUniqueElement };
                    return subMenus;
            }
            return "";
        }

        bool IsProjectOpen(EA.Repository Repository)
        {
            try
            {
                EA.Collection c = Repository.Models;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void EA_GetMenuState(EA.Repository Repository, string Location, string MenuName, string ItemName, ref bool IsEnabled, ref bool IsChecked)
        {
            if (IsProjectOpen(Repository))
            {
            }
            else
            {
                // If no open project, disable all menu options
                IsEnabled = false;
            }
        }

        public void EA_MenuClick(EA.Repository Repository, string Location, string MenuName, string ItemName)
        {
            switch (ItemName)
            {
                // user has clicked the menuHello menu option

                case menuAss:
                    this.Ass();
                    break;
                case menuIssues:
                    this.testIssues();
                    break;

                case menuLCOM:
                    this.LCOM3();
                    break;

                 case menuCoupling:
                    this.coupling();
                    break;

                case menuUniqueElement:
                    this.UniqueElements();
                    break;
                 
            }
        }

        string connectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = 'G://Mphil CS/1st Semester/Reverse Engineering/REAssignment.eapx';";

        private void Ass()
        {
            cnn = new OleDbConnection(connectionString);

            string procedures = " Select count(*) from t_connector";
            string end = "select COUNT( End_Object_ID) from t_connector ";
            string start = "select COUNT( Start_Object_ID) from t_connector ";



            try
            {
                cnn.Open();
                cmd = new OleDbCommand(procedures, cnn);

                decimal count1 = Convert.ToDecimal(cmd.ExecuteScalar());
                OleDbCommand cmd1 = new OleDbCommand(start, cnn);
                decimal count2 = Convert.ToDecimal(cmd1.ExecuteScalar());
                OleDbCommand cmd2= new OleDbCommand(end, cnn);
                decimal count3 = Convert.ToDecimal(cmd2.ExecuteScalar());

                cmd.Dispose();
                cmd1.Dispose();
                cmd2.Dispose();
                cnn.Close();

                string newLine = Environment.NewLine;
                decimal result = count1 / (count2 * count3);

                MessageBox.Show("Assocition Density = " + result.ToString("n2") + newLine);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        // Display TestType Weight
        private void testIssues()
        {
            string procedures = "Select count (*) from t_issues";
            
            cnn = new OleDbConnection(connectionString);
            try
            {
                cnn.Open();
                cmd = new OleDbCommand(procedures, cnn);
                decimal count = Convert.ToDecimal(cmd.ExecuteScalar());

                cmd.Dispose();
                cnn.Close();

                string newLine = Environment.NewLine;


                MessageBox.Show("Test Issues  = " + count.ToString("n2") + newLine);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }

       

        private void LCOM3()
        {

            string procedures = "Select count (operationid) from t_operation, t_object where t_object.Object_Type = 'class'";
            string methods_accessed = "Select count (operationid) from t_operation, t_object where t_object.Scope='public' and t_object.Object_Type='class'";
            string variables = "Select count (operationid) from t_operationparams";
            string countOfCoupling = "Select count (connector_id) from t_connector,t_object where t_connector.Connector_ID = t_object.Object_ID and t_object.Object_Type = 'class'";
            string totalClasses = "Select count (object_id) from t_object where Object_Type='class'";
            cnn = new OleDbConnection(connectionString);

            try
            {
                cnn.Open();
                cmd = new OleDbCommand(procedures, cnn);
                decimal count1 = Convert.ToDecimal(cmd.ExecuteScalar());

                OleDbCommand cmd1 = new OleDbCommand(methods_accessed, cnn);
                decimal count2 = Convert.ToDecimal(cmd1.ExecuteScalar());

                OleDbCommand cmd2 = new OleDbCommand(variables, cnn);
                decimal count3 = Convert.ToDecimal(cmd2.ExecuteScalar());

                OleDbCommand cmd3 = new OleDbCommand(countOfCoupling, cnn);
               
                cmd.Dispose();
                cmd1.Dispose();
                cmd2.Dispose();
                
                cnn.Close();

                string newLine = Environment.NewLine;
                decimal lcom3 = (count1 - (count2 / count3)) / (count1 - 1);
               

                MessageBox.Show("Cohesion = " + lcom3.ToString("n2") + newLine);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }

        private void coupling()
        {
            string procedures = "Select count (operationid) from t_operation, t_object where t_object.Object_Type = 'class'";
            string methods_accessed = "Select count (operationid) from t_operation, t_object where t_object.Scope='public' and t_object.Object_Type='class'";
            string variables = "Select count (operationid) from t_operationparams";
            string countOfCoupling = "Select count (connector_id) from t_connector,t_object where t_connector.Connector_ID = t_object.Object_ID and t_object.Object_Type = 'class'";
            string totalClasses = "Select count (object_id) from t_object where Object_Type='class'";
            cnn = new OleDbConnection(connectionString);

            try
            {
                cnn.Open();
                cmd = new OleDbCommand(procedures, cnn);
                

                OleDbCommand cmd2 = new OleDbCommand(variables, cnn);
                decimal count3 = Convert.ToDecimal(cmd2.ExecuteScalar());

                OleDbCommand cmd3 = new OleDbCommand(countOfCoupling, cnn);
                decimal count4 = Convert.ToDecimal(cmd3.ExecuteScalar());

                OleDbCommand cmd4 = new OleDbCommand(totalClasses, cnn);
                decimal count5 = Convert.ToDecimal(cmd4.ExecuteScalar());

                cmd.Dispose();
                
                cmd3.Dispose();
                cmd4.Dispose();
                cnn.Close();

                string newLine = Environment.NewLine;
                decimal cbo = (count4 / count5);

                MessageBox.Show( "Coupling = " + cbo.ToString("n2"));

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }

        private void UniqueElements()
        {
            string procedures = "SELECT    COUNT( ElementID) AS NumberOfUniqueElements FROM   t_version";

            cnn = new OleDbConnection(connectionString);
            try
            {
                cnn.Open();
                cmd = new OleDbCommand(procedures, cnn);
                decimal count = Convert.ToDecimal(cmd.ExecuteScalar());

                cmd.Dispose();
                cnn.Close();

                string newLine = Environment.NewLine;


                MessageBox.Show("Unique Elements = " + count.ToString("n2") + newLine);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }
        public void EA_Disconnect()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
