using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Xml.Linq;
using System.Runtime.Remoting.Messaging;
using EA;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;

namespace RePlugin
{
    public class RePlugin
    {
        // define menu constants
        const string menuHeader = "-&RePlugin"; 

        const string menuQualityOfDesignCode = "&Quality of code design";
        const string menuNoOfDesignMethods = "&Number of Design Methods";
        const string menuCohesion = "&Cohesion";
        const string menuErrorPrevention = "&Error Prevention";
        const string menuCoupling = "&Coupling";
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
                    string[] subMenus = { menuQualityOfDesignCode, menuNoOfDesignMethods, menuCohesion,menuErrorPrevention,menuCoupling};
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
                switch (ItemName)
                {
                    case menuQualityOfDesignCode:
                        IsEnabled = true;
                        break;
                    case menuNoOfDesignMethods:
                        IsEnabled = true;
                        break;
                    case menuCohesion:
                        IsEnabled = true;
                        break;
                    case menuErrorPrevention:
                        IsEnabled = true;
                        break;
                    case menuCoupling:
                        IsEnabled = true;
                        break;
                    // there shouldn't be any other, but just in case disable it.
                    default:
                        IsEnabled = false;
                        break;
                }
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
                case menuQualityOfDesignCode:
                    this.QualityOfDesignCode();
                    break;
                case menuNoOfDesignMethods:
                    this.NoOfDesignMethods();
                    break;
                case menuCohesion:
                    this.Cohesion();
                    break;
                case menuErrorPrevention:
                    this.ErrorPrevention();
                    break;
                case menuCoupling:
                    this.Coupling();
                    break;
            }
        }
        OleDbConnection cnn;
        OleDbCommand cmd;

        string connetionString = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = 'E://mphil/RE/Queries.eapx';";

        private void QualityOfDesignCode()
        {
            string totalClassesPublic = "SELECT COUNT(t_object.Object_Type) AS TotalClassPublic FROM t_object JOIN t_operation ON t_object.Object_ID = t_operation.Object_ID JOIN t_attribute ON t_object.Object_ID = t_attribute.Object_ID WHERE t_object.Object_Type = 'Class' AND t_attribute.Scope = 'public';";
            string totalClasses = "select count(Object_Type) AS TotalClass from t_object JOIN t_operation JOIN t_attribute where Object_Type = 'Class';";

            cnn = new OleDbConnection(connetionString);

            try
            {
                cnn.Open();
                OleDbCommand cmd3 = new OleDbCommand(totalClassesPublic, cnn);
                decimal count4 = Convert.ToDecimal(cmd3.ExecuteScalar());

                OleDbCommand cmd4 = new OleDbCommand(totalClasses, cnn);
                decimal count5 = Convert.ToDecimal(cmd4.ExecuteScalar());

                cmd3.Dispose();
                cmd4.Dispose();

                cnn.Close();

                string newLine = Environment.NewLine;
                decimal qdc = (count4 / count5);

                MessageBox.Show("Quality Of Design Code = " + qdc.ToString("n2"));

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection! ");
            }

        }

        private void NoOfDesignMethods()
        {
            string countOfdesignMethods = "select COUNT(*) as TotalDesignMethods from t_operation where Name='load' AND Scope='public';";
            cnn = new OleDbConnection(connetionString);

            try
            {
                cnn.Open();
                OleDbCommand cmd3 = new OleDbCommand(countOfdesignMethods, cnn);
                decimal count4 = Convert.ToDecimal(cmd3.ExecuteScalar());

                cmd3.Dispose();

                cnn.Close();

                string newLine = Environment.NewLine;
                decimal dm = (count4);

                MessageBox.Show("No. of Design Methods = " + dm.ToString("n2"));

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }

        }

        private void Cohesion()
        {
            string procedures = "Select count (operationid) from t_operation, t_object where t_object.Object_Type = 'class';";
            string methods_accessed = "Select count (operationid) from t_operation, t_object where t_object.Scope='public' and t_object.Object_Type='class'";
            string variables = "Select count (operationid) from t_operationparams";

            cnn = new OleDbConnection(connetionString);

            try
            {
                cnn.Open();
                cmd = new OleDbCommand(procedures, cnn);
                decimal count1 = Convert.ToDecimal(cmd.ExecuteScalar());

                OleDbCommand cmd1 = new OleDbCommand(methods_accessed, cnn);
                decimal count2 = Convert.ToDecimal(cmd1.ExecuteScalar());

                OleDbCommand cmd2 = new OleDbCommand(variables, cnn);
                decimal count3 = Convert.ToDecimal(cmd2.ExecuteScalar());

                cmd.Dispose();
                cmd1.Dispose();
                cmd2.Dispose();

                cnn.Close();

                string newLine = Environment.NewLine;
                decimal lcom3 = (count1 - (count2 / count3)) / (count1 - 1);

                MessageBox.Show("Cohesion = " + lcom3.ToString("n2"));

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }

        }

        private void ErrorPrevention()
        {
            string NumberOfErrorMessagesEncountered = "SELECT COUNT(*) AS error_count from t_issues;";
            string totalActions = "SELECT count(*) from t_operation;";

            cnn = new OleDbConnection(connetionString);

            try
            {
                cnn.Open();
                OleDbCommand cmd3 = new OleDbCommand(NumberOfErrorMessagesEncountered, cnn);
                decimal noOfErrors = Convert.ToDecimal(cmd3.ExecuteScalar());

                OleDbCommand cmd4 = new OleDbCommand(totalActions, cnn);
                decimal totalactions = Convert.ToDecimal(cmd4.ExecuteScalar());

                cmd3.Dispose();
                cmd4.Dispose();

                cnn.Close();

                string newLine = Environment.NewLine;
                decimal ep = (noOfErrors/ totalactions);
                if (totalactions != 0)
                {
                    if (ep == 0)
                    {
                        MessageBox.Show("Error prevention is Enabled");
                    }
                    else
                    {
                        MessageBox.Show("Error prevention is Not Enabled");
                    }
                }
                else
                {
                    MessageBox.Show("Not Enabled (totalactions is zero)");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection! ");
            }
        }

        private void Coupling()
        {
            string CouplingAmongClasses = "select count (connector_id) from t_connector, t_object where t_connector.connector_id=t_object.Object_id and t_object.Object_Type='class'";
            string clases = "select count(Object_Type) from t_object JOIN t_operation where Object_Type = 'Class';";

            cnn = new OleDbConnection(connetionString);

            try
            {
                cnn.Open();
                OleDbCommand cmd3 = new OleDbCommand(CouplingAmongClasses, cnn);
                decimal count4 = Convert.ToDecimal(cmd3.ExecuteScalar());

                OleDbCommand cmd4 = new OleDbCommand(clases, cnn);
                decimal count5 = Convert.ToDecimal(cmd4.ExecuteScalar());

                cmd3.Dispose();
                cmd4.Dispose();

                cnn.Close();

                string newLine = Environment.NewLine;
                decimal coupling = (count4 / count5);

                MessageBox.Show("Coupling = " + coupling.ToString("n2"));

            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection! ");
            }

        }
        public void EA_Disconnect()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

    }
}
