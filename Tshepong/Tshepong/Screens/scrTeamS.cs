using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.IO;
using Analysis = clsAnalysis;
using TB = clsTable;
using DB = clsDBase;
using Base = clsMain;
using System.Threading;
using System.Data.OleDb;
using MetaReportRuntime;
using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip;
using System.Net;
using System.Net.Mail;

namespace Tshepong
{
    public partial class scrTeamS : Form
    {
        #region Declarations
        int columnnr = 0;
        int intNoOfDays = 0; 
        int noOFDay = 0;  
        DateTime sheetfhs = new DateTime(); 
        DateTime sheetlhs = new DateTime(); 
        int importdone = 0; 
        DataTable fixShifts = new DataTable(); 
        int intStartDay = 0; 
        int intEndDay = 0; 
        int intStopDay = 0; 
        int workedShiftsFixedClockedShift = 0;  
        int exitValue = 0;  
        string searchEmplNr = string.Empty;
        string searchEmplName = string.Empty;
        string searchEmplGang = string.Empty;
        string searchEmplNr2 = string.Empty;
        string searchEmplGang2 = string.Empty;
        string strWhereSection = string.Empty;
        int rowindex;
        string Path = string.Empty;
        string strMO = string.Empty;
        string strMonthShifts = string.Empty;

        clsBL.clsBL BusinessLanguage = new clsBL.clsBL();
        clsTable.clsTable TB = new clsTable.clsTable();
        clsGeneral.clsGeneral General = new clsGeneral.clsGeneral();
        clsShared Shared = new clsShared(); 
        clsTableFormulas TBFormulas = new clsTableFormulas();
        clsMain.clsMain Base = new clsMain.clsMain();
        clsAnalysis.clsAnalysis Analysis = new clsAnalysis.clsAnalysis();
        SqlConnection myConn = new SqlConnection();
        SqlConnection AConn = new SqlConnection();
        SqlConnection AAConn = new SqlConnection();
        SqlConnection BaseConn = new SqlConnection();
        System.Collections.Hashtable buttonCollection = new System.Collections.Hashtable();

        Dictionary<string, string> dictPrimaryKeyValues = new Dictionary<string, string>();
        Dictionary<string, string> dictGridValues = new Dictionary<string, string>();
        Dictionary<string, string> dict = new Dictionary<string, string>();
        Dictionary<string, string> GangTypes = new Dictionary<string, string>();

        string strEarningsCode = string.Empty;
        string strprevPeriod = string.Empty;
        string prevDatabaseName = string.Empty;
        string strWhere = string.Empty;
        string strWherePeriod = string.Empty;
        string strActivity = string.Empty;
        string strMiningIndicator = string.Empty;
        string strServerPath = string.Empty;
        string strName = string.Empty;
        string strMetaReportCode = "BSFnupmWkNxm8ZAA1ZhlOgL8fNdMdg4zhJj/j6T0vEyG9aSzk/HPwYcrjmawRGou66hBtseT7qJE+9hbEq9jces6bcGJmtz4Ih8Fic4UIw0Kt2lEffc05nFdiD2aQC0m";

        string dbPath = string.Empty;

        string[] ClockedShifts = new string[5];
        string[] OffShifts = new string[5];
        int intFiller = 0;
        int intCounter = 0;

        List<string> lstNames = new List<string>();
        List<string> lstTableColumns = new List<string>();
        List<string> lstPrimaryKeyColumns = new List<string>();

        Int64 intProcessCounter = 0;
        StringBuilder strSqlAlter = new StringBuilder();

        DataTable Survey = new DataTable();
        DataTable Labour = new DataTable();
        DataTable Workers = new DataTable();
        DataTable Miners = new DataTable();
        DataTable Designations = new DataTable();
        DataTable Drillers = new DataTable();
        DataTable Clocked = new DataTable();
        DataTable Rates = new DataTable();
        DataTable EmplPen = new DataTable();
        DataTable Configs = new DataTable();
        DataTable Offdays = new DataTable();
        DataTable GangLink = new DataTable();
        DataTable Abnormal = new DataTable();
        DataTable BonusSeq = new DataTable();
        DataTable Monitor = new DataTable();
        DataTable Calendar = new DataTable();
        DataTable Production = new DataTable();
        DataTable PayrollSends = new DataTable();
        DataTable Factors = new DataTable();
        DataTable Status = new DataTable();
        DataTable BonusShifts = new DataTable();
        DataTable newDataTable = new DataTable();

        string[] arrArgs = new string[1] { "" };

        SqlDataAdapter minersTA = new SqlDataAdapter();
        BindingSource bSource = new BindingSource();
        SqlCommandBuilder _cmdBuilder = new SqlCommandBuilder();


        //**************************************************************
        //*************  Tshepong APP.CONFIG BASIL = FS3032\SQLEXPRESS
        //     <add key="DevelopmentIntegrity" value="Trusted_Connection = True" />
        //<add key="DevelopmentServerPath" value="QwA6AFwAXABpAEMAYQBsAGMAXABcAEgAYQByAG0AbwBuAHkAXABcAFAAaABhAGsAaQBzAGEAXABcAEQAZQB2AGUAbABvAHAAbQBlAG4AdABcAFwARABhAHQAYQBiAGEAcwBlAHMAXABcAEQAYQB0AGEA" />
        //<add key="DevelopmentServerName" value="RgBTADMAMAAzADIAXABTAFEATABFAFgAUABSAEUAUwBTAA==" />
        //<add key="DevelopmentBackupPath" value="QwA6AFwAaQBDAGEAbABjAFwASABhAHIAbQBvAG4AeQBcAFAAaABhAGsAaQBzAGEAXABEAGUAdgBlAGwAbwBwAG0AZQBuAHQAXABEAGEAdABhAGIAYQBzAGUAcwA="/>
        //<add key="DevelopmentDrive" value="C:"/>

        private ExcelDataReader.ExcelDataReader spreadsheet = null;

        ToolTip tooltip = new ToolTip();
        #endregion

        public scrTeamS()
        {
            InitializeComponent();
            //string[] args = Program.Args;
            //arrArgs = args;
            //newdatabase(arrArgs);

        }

        internal void scrTeamsLoad(string Period, string Region, string BussUnit, string Userid, string MiningType, string BonusType, string Environment)
        {
            #region disable all functions
            //Disable all menu functions.
            foreach (ToolStripMenuItem IT in menuStrip1.Items)
            {
                if (IT.DropDownItems.Count > 0)
                {
                    foreach (ToolStripMenuItem ITT in IT.DropDownItems)
                    {
                        if (ITT.DropDownItems.Count > 0)
                        {
                            foreach (ToolStripMenuItem ITTT in ITT.DropDownItems)
                            {
                                ITTT.Enabled = false;
                            }
                        }
                        else
                        {
                            ITT.Enabled = false;
                        }
                    }
                }
                else
                {
                    IT.Enabled = false;
                }
            }
            #endregion

            #region declarations
            BusinessLanguage.Period = Period.Trim();
            BusinessLanguage.Region = Region.Trim();
            BusinessLanguage.BussUnit = BussUnit.Trim();
            BusinessLanguage.Userid = Userid.Trim();
            BusinessLanguage.MiningType = MiningType.Trim();
            BusinessLanguage.BonusType = BonusType.Trim();
            txtMiningType.Text = MiningType.Trim();
            txtBonusType.Text = BonusType.Trim();
            strServerPath = Environment.Trim();
            BusinessLanguage.Env = Environment.Trim();
            txtDatabaseName.Text = "STPTM4000";
            //Display dbname in text box.
            //txtDatabaseName.Text = txtDatabaseName.Text.Trim() + BusinessLanguage.Period;
            Base.DBName = txtDatabaseName.Text.Trim();
            Base.Period = BusinessLanguage.Period;

            //Setup the environment BEFORE the databases are moved to the classes.  This is because the environment path forms
            //part of the fisical name of the db

            setEnvironment();

            #endregion

            #region Connections
            //Open Connections and create classes

            AAConn = Analysis.AnalysisConnection;
            AAConn.Open();
            BaseConn = Base.BaseConnection;
            BaseConn.Open();

            #endregion

            DataTable useraccess = Base.SelectAccessByUserid(BusinessLanguage.Userid, Base.BaseConnectionString);

            #region Assign useraccess

            //BusinessLanguage.BussUnit = useraccess.Rows[0]["BUSSUNIT"].ToString().Trim();
            BusinessLanguage.Resp = useraccess.Rows[0]["RESP"].ToString().Trim();

            foreach (DataRow dr in useraccess.Rows)
            {
                string strCodeName = dr[6].ToString().Trim();
                foreach (ToolStripMenuItem IT in menuStrip1.Items)
                {
                    if (IT.DropDownItems.Count > 0)
                    {
                        foreach (ToolStripMenuItem ITT in IT.DropDownItems)
                        {
                            if (ITT.DropDownItems.Count > 0)
                            {
                                foreach (ToolStripMenuItem ITTT in ITT.DropDownItems)
                                {
                                    if (ITTT.Name.Trim() == strCodeName)
                                    {
                                        ITTT.Enabled = true;
                                    }
                                }
                            }
                            else
                                if (ITT.Name.Trim() == strCodeName)
                                {
                                    ITT.Enabled = true;
                                }
                        }
                    }
                    else
                    {
                        if (IT.Name.Trim() == strCodeName)
                        {
                            IT.Enabled = true;
                        }

                    }
                }

            }
            #endregion

            #region General
            //Display user details
            txtUserDetails.Text = BusinessLanguage.Userid + " - " + BusinessLanguage.Region + " - " + BusinessLanguage.BussUnit;
            //txtDatabaseName.Text = BusinessLanguage.BussUnit;

            txtPeriod.Text = BusinessLanguage.Period;

            // Set up the delays for the ToolTip.
            tooltip.AutoPopDelay = 5000;
            tooltip.InitialDelay = 1000;
            tooltip.ReshowDelay = 500;
            //Force the ToolTip text to be displayed whether or not the form is active.
            tooltip.ShowAlways = true;

            //Set up the ToolTip text for the Button and Checkbox.
            tooltip.SetToolTip(this.btnImportADTeam, "Clocked Shifts");
            tooltip.SetToolTip(this.tabLabour, "Bonus Shifts");
            tooltip.SetToolTip(this.btnSearch, "Search");

            listBox2.Enabled = false;
            listBox3.Enabled = false;


            #endregion

            #region Status button collection

            //Add the buttons needed for this bonus scheme and that are on the STATUS tab.
            buttonCollection["tabCalendar"] = btnLockCalendar;
            buttonCollection["tabSurvey"] = btnLockSurvey;
            buttonCollection["tabLabour"] = btnLockBonusShifts;
            buttonCollection["tabGangLinking"] = btnLockGangLink;
            buttonCollection["tabDrillers"] = btnLockDrillers;
            buttonCollection["tabMiners"] = btnLockMiners;
            buttonCollection["tabEmplPen"] = btnLockEmplPen; 
            buttonCollection["GanglinkEARN10"] = btnBaseCalcs;
            buttonCollection["GanglinkEARN70"] = btnGangLinkCalcs; 
            buttonCollection["Bonus Report Process - Phase 1"] = btnBonusPrints;
            buttonCollection["BonusShiftsEARN60"] = btnBonusShiftsCalcs;
            buttonCollection["Bonus Report Process - Phase 1"] = btnBonusPrints;
            buttonCollection["MinersEARN10"] = btnMinersCalc;
            buttonCollection["Input Process"] = btnInputProcess; 
            buttonCollection["Paysend"] = btnLockPaysend;
            #endregion

            #region BaseData Extracts
            
            //Extract Base data
            extractDesignations();
            extractConfiguration();
            //extractOccupations();
            //extractEarningsCodes();

            #endregion

            //Extract Tab Info
            loadInfo();
            //extractEarningsCodes();
            //Load tablename of tables in database
            //extractDBTableNames(listBox1);

            //Create the tab names
            foreach (TabPage tp in tabInfo.TabPages)
            {
                tp.Text = tp.Tag.ToString();
            }

        }

        private void setEnvironment()
        {

            Base.Drive = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Drive"];
            Base.Integrity = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Integrity"];
            Base.Userid = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Userid"])).Trim();
            Base.PWord = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Password"])).Trim();
            Base.ServerName = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.MachineName = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "MachineName"].Trim();
            Base.BaseConnectionString = Base.ServerName;
            Base.Directory = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerPath"])).Trim();

            Analysis.Drive = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Drive"];
            Analysis.Integrity = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Integrity"];
            Analysis.Userid = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Userid"])).Trim();
            Analysis.PWord = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Password"])).Trim();
            Analysis.ServerName = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Analysis.AnalysisConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();

            Base.ClockConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.DBConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.StopeConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.AnalysisConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.BackupPath = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "BackupPath"])).Trim();

            Base.StopeDatabaseName = "STPTM1000" + BusinessLanguage.Period;
            Base.DevDatabaseName = "DEVTM1000" + BusinessLanguage.Period;

            #region oleDBConnectionStringBuilder
            OleDbConnectionStringBuilder builder = new OleDbConnectionStringBuilder();
            builder.ConnectionString = @"Data Source=" + Base.ServerName;
            builder.Add("Provider", "SQLOLEDB.1");
            builder.Add("Initial Catalog", Base.DBName);
            //builder.Add("Persist Security Info", "False");
            builder.Add("User ID", Base.Userid);
            builder.Add("Password", Base.PWord);


            string strdb = Base.DBName;

            if (strServerPath.ToString().Contains("Development"))
            {
                strServerPath = "Development";
            }

           string strPath = "c:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\" + strdb.Replace(BusinessLanguage.Period, "").Replace("4000", "Conn") + ".udl";

            FileInfo fil = new FileInfo(strPath);

            try
            {
                File.Delete(strPath);
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                MessageBox.Show("delete of udl failed: " + ex.Message);
            }

            switch (strServerPath)
            {
                case "Test":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Trusted_Connection", "True");
                    break;


                case "Development":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Integrated Security", "SSPI");
                    builder.Add("Trusted_Connection", "True");
                    break;

                case "Production":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Trusted_Connection", "True");
                    break;

            }

            //MessageBox.Show("Path: " + strPath);
            bool _check = Shared.CreateUDLFile(strPath, builder);

            if (_check)
            { }
            else
            {
                MessageBox.Show("Error in creation of UDL file", "ERROR", MessageBoxButtons.OK);
            }
            //xxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            #endregion
            //xxxxxxxxxxxxxxxxxxxxxxxxxxxxx

            myConn.ConnectionString = Base.DBConnectionString;

            //xxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        }

        public void extractDBTableNames(ListBox lstbox)
        {
            connectToDB();

            if (myConn.State == ConnectionState.Open)
            {
                List<string> lstTableNames = Base.getListOfTableNamesInDatabase(Base.DBConnectionString);
                Base.DBTables = lstTableNames;
                lstbox.Items.Clear();
                switch (lstTableNames.Count)
                {
                    case 0:
                        lstbox.Items.Add("No tables in database");
                        break;
                    default:
                        foreach (string s in lstTableNames)
                        {
                            lstbox.Items.Add(s);
                        }
                        break;
                }

            }
        }

        private void extractConfiguration()
        {

            Configs = Base.SelectConfigs(Base.BaseConnectionString, BusinessLanguage.MiningType, BusinessLanguage.BonusType);

            grdConfigs.DataSource = Configs;

            foreach (DataRow dr in Configs.Rows)
            {
                //This extract the value identifying the first 3 digits that the gang must conform to.
                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGLINKING"
                    && dr["PARM1"].ToString().Trim() == "MININGTYPE")
                {
                    for (int i = 5; i <= 10; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strMiningIndicator = strMiningIndicator + ",'" + dr[i].ToString().Trim() + "'";
                        }
                    }

                    strMiningIndicator = "(" + strMiningIndicator.Trim().Substring(1) + ")";

                }


                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGLINKING"
                    && dr["PARM1"].ToString().Trim() == "ACTIVITY")
                {
                    strActivity = string.Empty;

                    for (int i = 5; i <= 10; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strActivity = strActivity + ",'" + dr[i].ToString().Trim() + "'";
                        }
                    }

                    strActivity = "(" + strActivity.Trim().Substring(1) + ")";
                }

                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGTYPE"
                   && dr["BONUSTYPE"].ToString().Trim() == txtBonusType.Text.Trim()
                   && dr["PARM1"].ToString().Trim() == "GANGTYPES")
                {
                    for (int i = 5; i <= Configs.Columns.Count - 1; i++)
                    {
                        if (dr[i].ToString().Trim() == "Q")
                        {
                        }
                        else
                        {
                            cboGangLinkGangType.Items.Add(dr[i].ToString().Trim());
                        }
                    }
                }

                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGTYPE"
                        && dr["PARM1"].ToString().Trim() == "INDICATOR")
                {

                    GangTypes.Add(dr["PARM2"].ToString().Trim(), dr["PARM3"].ToString().Trim());

                }

            }

            loadMO();
        }

        private void loadMO()
        {
            strMO = "";
            foreach (DataRow dr in Configs.Rows)
            {
                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGLINKING" &&
                    dr["PARM1"].ToString().Trim() == "MO"
                    && dr["PARM2"].ToString().Trim() == txtSelectedSection.Text)
                {
                    for (int i = 6; i <= Configs.Columns.Count - 1; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strMO = strMO + ",'" + dr[i].ToString().Trim() + "'";
                        }
                    }

                    strMO = "(" + strMO.Trim().Substring(1) + ")";
                }
            }
        }

        private void extractDesignations()
        {
            cboDesignation.Items.Clear();
            Designations = Base.GetDataByDestination("grdMiners", Base.BaseConnectionString, BusinessLanguage.MiningType, BusinessLanguage.BonusType);

            foreach (DataRow x in Designations.Rows)
            {
                cboDesignation.Items.Add(x["DESIGNATION"].ToString().Trim() + "  -  " + x["DESIGNATION_DESC"].ToString().Trim());
            }
        }

        private void extractEarningsCode()
        {
            //Extract the records by miningtype, bonustype and paymethod
            DataTable t = Base.GetDataByMintypeBontypePaytype(txtMiningType.Text, txtBonusType.Text, "3", Base.BaseConnectionString);
            strEarningsCode = t.Rows[0]["EARNINGSCODE"].ToString().Trim();
        }

        private void loadFactors()
        {
            //Check if Factors exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Factors");

            if (intCount > 0)
            {
                //YES

                Factors = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Factors", " where period = '" + BusinessLanguage.Period + "'");

            }
            else
            {
                //NO - Factors DOES NOT EXIST 
            }

            grdFactors.DataSource = Factors;

            hideColumnsOfGrid("grdFactors");

            cboVarName.Items.Clear();

            foreach (DataRow row in Factors.Rows)
            {
                cboVarName.Items.Add(row["VARNAME"].ToString().Trim());
            }
        }

        private void loadInfo()
        {
            strWherePeriod = "  where period = '" + BusinessLanguage.Period + "'";
            //Check if records in calendar exists with the selected period
            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, 
                       "CALENDAR", " where period = '" + BusinessLanguage.Period + "'");


            if (Calendar.Rows.Count > 0)
            {

                //Run the extraction of the primary keys on its own threat.
                Shared.extractPrimaryKeys(Base);

                //Run the extraction of the views.
                Shared.createViews(Base);

                if (myConn.State == ConnectionState.Open)
                {
                    evaluateAll();
                }
                else
                {
                    connectToDB();
                    evaluateAll();
                }

            }
            else
            {
                //NO....
                //1. Get Previous months info  ==> MAG NIE MEER HIERIN GAAN NIE!!!!!!!!!!!!!!!!!!!!!!!

                getHistory();

                //2. Check if PREVIOUS months DB exists
                //if (BusinessLanguage.checkIfFileExists(Base.Directory + "\\" + prevDatabaseName + Base.DBExtention))
                //{
                //3. If exist - Create this selected DB and copy Formulas, Rates and Factors to the new database.
                DialogResult result = MessageBox.Show("Do you want to start a new Bonus Period: " + BusinessLanguage.Period + "?",
                                       "Information", MessageBoxButtons.YesNo);

                switch (result)
                {
                    case DialogResult.Yes:
                        this.Cursor = Cursors.WaitCursor;
                        backupAndRestoreDB();
                        copyFormulas();
                        extractDBTableNames(listBox1);
                        //Base.createNewPeriodsData(Base.DBConnectionString, BusinessLanguage.Period.Trim(), strprevPeriod.Trim());
                        //Base.deleteExtras2000(Base.DBConnectionString);
                        //createAndCopyCalendar();

                        //Run the extraction of the primary keys on its own threat.
                        Shared.extractPrimaryKeys(Base);
                        evaluateAll();

                        this.Cursor = Cursors.Arrow;
                        break;

                    case DialogResult.No:
                        btnSelect_Click("METHOD", null);
                        break;
                }

              

            }
        }

        private void evaluateAll()
        {
            evaluateAbnormal();
            evaluateCalendar();
            evaluateProduction();
            evaluateSurvey();
            evaluateClockedShifts();
            evaluateLabour();
            evaluateMiners();
            evaluateDrillers();
            evaluateGangLinking();
            evaluateEmployeePenalties();
            evaluateOffDays();
            evaluateRates();
            evaluateFactors();
            extractDBTableNames(listBox1);
        }

        private void evaluateDrillers()
        {
            Drillers.Rows.Clear();

            loadDrillers();

            lstDrillers.Items.Clear();

        }

        private void loadDrillers()
        {
            //Check if ganglinking exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "DRILLERS");

            if (intCount > 0)
            {
                //YES
                Drillers = TB.createDataTableWithAdapter(Base.DBConnectionString,
                           "SELECT * FROM DRILLERS WHERE SECTION = '" + txtSelectedSection.Text.Trim() + 
                           "'AND PERIOD = '" + txtPeriod.Text.Trim() + "'");
                txtAutoDGang.Clear();
            }
            else
            {

            }

            grdDrillers.DataSource = Drillers;

            hideColumnsOfGrid("grdDrillers");

        }


        private void copyFormulas()
        {
            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable dtBaseFormulas = Analysis.SelectAllFormulasPerDatabaseName(Base.DBCopyName + strprevPeriod.Trim(), Base.AnalysisConnectionString);
            if (dtBaseFormulas.Rows.Count > 0)
            {
                foreach (DataRow row in dtBaseFormulas.Rows)
                {
                    //Check if the receiving table already contains this formula.
                    object intCount = Analysis.countcalcbyname(Base.DBName + BusinessLanguage.Period.Trim(), row["TABLENAME"].ToString(), 
                                      row["CALC_NAME"].ToString(), Base.AnalysisConnectionString);

                    if ((int)intCount > 0)
                    {
                        //rename the formula name to be inserted to NEW

                    }
                    else
                    {
                        //insert the formula.
                        Base.CopyFormulas(Base.DBName + strprevPeriod.Trim(),
                                          Base.DBName + BusinessLanguage.Period.Trim(), 
                                          Analysis.AnalysisConnectionString);
                        break;
                    }
                }
            }
            else
            {
                MessageBox.Show("No formulas exist on " + "\n" + "database: " + Base.DBCopyName + "\n" + "tablename: " + TB.TBCopyName + "\n" + "therefor" + "\n" + "nothing will be copied", "Information", MessageBoxButtons.OK);
            }
        }

        private void countSections()
        {

            panel3.Enabled = true;
            panel4.Enabled = true;
            listBox2.Enabled = true;
            listBox3.Enabled = true;

            if (listBox1.Items.Contains("SURVEY"))
            {
                string strSQL = "SELECT * from SURVEY;";

                Survey = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                if (Survey.Rows.Count > 0)
                {
                    if (Survey.Columns.Contains("SECTION"))
                    {

                    }
                    else
                    {
                        strSQL = " Alter table Survey add SECTION varchar(50);";
                        TB.InsertData(Base.DBConnectionString, strSQL);

                        strSQL = " Update survey set SECTION = substring(Contract,2,3);";
                        TB.InsertData(Base.DBConnectionString, strSQL);
                    }

                    grdSurvey.DataSource = Survey;

                }
            }
            else
            {

            }
        }

        private void confirmCopyandCreate()
        {
            listBox2.Items.Add("No sections found");

            this.Cursor = Cursors.WaitCursor;

            #region Create the new DB
            //Create the new database
            Base.createDatabase(Base.DBName, Base.ServerName);

            myConn = Base.DBConnection;
            myConn.Open();

            TB.createEmployeePenalties(Base.DBConnectionString);
            TB.createCalendarTable(Base.DBConnectionString);
            TB.createOffday(Base.DBConnectionString);
            TB.createEmployeePenalties(Base.DBConnectionString);

            //Extract Calendar again and insert into 
            DataTable calendar = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Calendar");
            grdCalendar.DataSource = calendar;

            listBox2.Items.Clear();
            listBox2.Items.Add("No sections exist yet");

            panel2.Enabled = false;
            panel3.Enabled = false;
            panel4.Enabled = false;

            this.Cursor = Cursors.Arrow;

            #endregion

        }

        private void getHistory()
        {
            #region Generate previous months db name
            //Calculate the previous months db name
            string Year = txtPeriod.Text.Trim().Substring(0, 4);
            strprevPeriod = txtPeriod.Text.Trim();

            if (txtPeriod.Text.Trim().Substring(txtPeriod.Text.Trim().Length - 2) == "01")
            {
                strprevPeriod = Convert.ToString(Convert.ToInt16(Year) - 1) + "12";
                prevDatabaseName = Base.DBName.Replace(txtPeriod.Text.Trim(), strprevPeriod);
            }
            else
            {
                string strMonth = Convert.ToString(Convert.ToInt16(txtPeriod.Text.Trim().Substring(txtPeriod.Text.Trim().Length - 2)) - 1);
                if (strMonth.Length == 1)
                {
                    strMonth = "0" + strMonth;
                }

                strprevPeriod = Year + strMonth;
                prevDatabaseName = Base.DBName.Replace(txtPeriod.Text.Trim(), strprevPeriod);
            }

            Base.DBCopyName = prevDatabaseName;

            #endregion
            //#region Generate previous months db name
        }

        private void createAndCopyCalendar()
        {

            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar");

            foreach (DataRow rr in Calendar.Rows)
            {
                rr["FSH"] = (Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(1)).ToString("yyyy-MM-dd");
                rr["LSH"] = (Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(31)).ToString("yyyy-MM-dd");
            }

            TB.saveCalculations2(Calendar, Base.DBConnectionString, "", "CALENDAR");
            this.Cursor = Cursors.Arrow;
        }

        private void createAndCopyStatus()
        {
            getHistory();

            TB.createStatusTable(Base.DBConnectionString);
            myConn.Close();

            //create the Status datatable from the previous periods'table.
            Base.DBName = Base.DBCopyName;
            connectToDB();

            Status = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Status");

            #region signoff from previous months DB and signon to this new DB

            myConn.Close();

            Base.DBName = TB.DBName;

            //Connect to the database that you want to copy from and load the tables into the listbox2.  Afterwards, change the db.dbname to the main database name.
            connectToDB();

            #endregion

            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            foreach (DataRow rr in Status.Rows)
            {
                strSQL.Append("insert into Status values('" + rr["MININGTYPE"].ToString().Trim() +
                              "','" + rr["BONUSTYPE"].ToString().Trim() + "','" + rr["SECTION"].ToString().Trim() +
                              "','" + txtPeriod.Text.Trim() + "','" + rr["CATEGORY"].ToString().Trim() + "','" + rr["PROCESS"].ToString().Trim() +
                              "','" + rr["STATUS"].ToString().Trim() + "','" + rr["LOCKED"].ToString().Trim() + "');");

            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            Application.DoEvents();
            TB.InsertData(Base.DBConnectionString, "Update Status set status = 'N', locked = '0'");
            Status = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Status");
            Application.DoEvents();
            this.Cursor = Cursors.Arrow;
        }


        private void backupAndRestoreDB()
        {
            //copy the data of the previous period to the current period.
            this.Cursor = Cursors.WaitCursor;
            Base.createNewPeriodsData(Base.DBConnectionString, BusinessLanguage.Period, strprevPeriod);
            this.Cursor = Cursors.Arrow;

        }

        private void createAndCopyDB()
        {

            this.Cursor = Cursors.WaitCursor;

            #region Create the new DB and base tables
            //Backup the PREVIOUS MONTH's db


            Base.createDatabase(Base.DBName, Base.ServerName);

            myConn = Base.DBConnection;
            myConn.Open();

            TB.createEmployeePenalties(Base.DBConnectionString);
            TB.createCalendarTable(Base.DBConnectionString);
            TB.createOffday(Base.DBConnectionString);
            TB.createEmployeePenalties(Base.DBConnectionString);
            TB.createMonitor(Base.DBConnectionString);
            TB.createStatusTable(Base.DBConnectionString);

            //Copy the formulas.
            //Copy the rates
            //Copy the factors


            panel2.Enabled = false;
            panel3.Enabled = false;
            panel4.Enabled = false;

            #endregion

            #region Create the Factors

            TB.createFactorTable(Base.DBConnectionString);

            #endregion

            #region Create the Rates

            TB.createRatesTable(Base.DBConnectionString);

            #endregion

            #region Create the rest
            DataTable tableFormulas = Analysis.SelectAllFormulasPerDatabaseName(Base.DBName, Base.AnalysisConnectionString);
            foreach (DataRow r in tableFormulas.Rows)
            {
                Analysis.insertQuery(Base.AnalysisConnectionString, Base.DBCopyName, r["TABLENAME"].ToString().Trim(), r["PHASENAME"].ToString().Trim(), r["FORMULA_NAME"].ToString().Trim(), r["FORMULA_CALL"].ToString().Trim(), r["CALC_NAME"].ToString().Trim(), r["CALC_SEQ"].ToString().Trim(), r["A"].ToString().Trim(), r["B"].ToString().Trim(), r["C"].ToString().Trim(), r["D"].ToString().Trim(), r["E"].ToString().Trim(), r["F"].ToString().Trim(), r["G"].ToString().Trim(), r["H"].ToString().Trim(), r["I"].ToString().Trim(), r["J"].ToString().Trim(), r["SAVECOLUMN"].ToString().Trim());
            }


            #endregion

            #region Close connection to new DB and signon to previous period DB
            myConn.Close();
            Base.DBName = Base.DBCopyName;
            //=================================================================================================
            //Connect to the database that you want to copy from and load the tables into the listbox2.  
            //Afterwards, change the db.dbname to the main database name.
            //=================================================================================================
            connectToDB();
            #endregion

            #region create a Factors and Rates and Status datatables from the previous months database

            DataTable factors = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Factors");
            DataTable Rates = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Rates");
            Calendar = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Calendar");
            Status = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Status");

            #endregion

            #region signoff from previous months DB and signon to this new DB

            myConn.Close();

            Base.DBName = TB.DBName;

            //Connect to the database that you want to copy from and load the tables into the listbox2.  Afterwards, change the db.dbname to the main database name.
            connectToDB();

            #endregion

            #region Copy Factors and Rates and Status from datatable to the new DB

            //========================================RATES===========================================================================
            StringBuilder strSQL = new StringBuilder();

            strSQL.Append("BEGIN transaction; ");


            foreach (DataRow r in Rates.Rows)
            {

                strSQL.Append("insert into rates values('" + r["MiningType"].ToString().Trim() + "','" + r["BonusType"].ToString().Trim() + "','" + r["Period"].ToString().Trim() + "','" + r["Rate_Type"].ToString().Trim() + "','" + r["LOW_Value"].ToString().Trim() + "','" + r["High_Value"].ToString().Trim() + "','" + r["Rate"].ToString().Trim() + "');");

            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));

            //=========================================STATUS==========================================================================
            strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");
            TB.InsertData(Base.DBConnectionString, "Delete from Status");

            foreach (DataRow rr in Status.Rows)
            {

                strSQL.Append("insert into STATUS values('" + rr["MININGTYPE"].ToString().Trim() +
                              "','" + rr["BONUSTYPE"].ToString().Trim() + "','" + rr["SECTION"].ToString().Trim() +
                              "','" + txtPeriod.Text.Trim() + "','" + rr["PROCESS"].ToString().Trim() + "','N'" +
                              ",'0');");

            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));

            //=========================================CALENDAR==========================================================================
            strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");
            TB.InsertData(Base.DBConnectionString, "Delete from Calendar");

            foreach (DataRow rr in Calendar.Rows)
            {
                DateTime FSH = Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(1);
                DateTime LSH = Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(31);

                strSQL.Append("insert into Calendar values('" + rr["MININGTYPE"].ToString().Trim() +
                              "','" + rr["BONUSTYPE"].ToString().Trim() + "','" + rr["SECTION"].ToString().Trim() +
                              "','" + txtPeriod.Text.Trim() + "','" + FSH.ToShortDateString() + "','" + LSH.ToShortDateString() +
                              "','23');");

            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));

            //===========================================FACTORS========================================================================
            strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");
            TB.InsertData(Base.DBConnectionString, "Delete from Factors");

            foreach (DataRow rr in factors.Rows)
            {
                strSQL.Append("insert into Factors values('" + rr["VARNAME"].ToString().Trim() + "','" + rr["VARVALUE"].ToString().Trim() + "');");

            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));

            myConn.Close();

            this.Cursor = Cursors.Arrow;

            #endregion

        }

        private void evaluateInputProcessStatus()
        {

            Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere +
                                                            " and category = 'Input Process'" +
                                                            " and period = '" + BusinessLanguage.Period + "'");


            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where category = 'Header' and process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

            }
            else
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where process = 'Input Process'" +
                                      " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where category = 'Header' and process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                btnLock.Text = "Lock";

            }

            evaluateStatus();

        }

        private void evaluateStatus()
        {

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "STATUS");

            if (intCount > 0)
            {
                //Status exists,  
                loadStatus();
            }
            else
            {
                createAndCopyStatus();
            }

        }

        private void statusChangeButtonColors()
        {
            foreach (DataRow rr in Status.Rows)
            {
                if (rr["CATEGORY"].ToString().Trim().Substring(0, 4) == "Exit")
                {
                    if (rr["STATUS"].ToString().Trim() == "Y")
                    {
                        btnRefresh.Visible = false;
                        btnx.Visible = false;

                        pictBox.Visible = false;
                        pictBox2.Visible = false;
                        calcTime.Enabled = false;
                    }
                }
                else
                {
                    if (rr["STATUS"].ToString().Trim() == "Y")
                    {
                        string strButtonName = rr["PROCESS"].ToString().Trim();
                        Control c = (Control)buttonCollection[strButtonName];
                        c.BackColor = Color.LightGreen;

                    }
                    else
                    {
                        if (rr["STATUS"].ToString().Trim() == "P")
                        {
                            string strButtonName = rr["PROCESS"].ToString().Trim();
                            Control c = (Control)buttonCollection[strButtonName];
                            c.BackColor = Color.Orange;
                        }
                        else
                        {
                            if (rr["STATUS"].ToString().Trim() == "N" &&
                                pictBox.Visible == true &&
                                rr["CATEGORY"].ToString().Trim().Substring(0, 4) == "CALC")
                            {
                                string strButtonName = rr["PROCESS"].ToString().Trim();
                                Control c = (Control)buttonCollection[strButtonName];
                                c.BackColor = Color.Orange;
                            }
                            else
                            {
                                string strButtonName = rr["PROCESS"].ToString().Trim();
                                Control c = (Control)buttonCollection[strButtonName];
                                c.BackColor = Color.PowderBlue;
                            }
                        }
                    }
                }

                Application.DoEvents();
            }
        }

        private void evaluateProduction()
        {
            // Display die Production info
            Production.Rows.Clear();
            loadProduction();

        }
        private void loadProduction()
        {
            //Check if Production exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Production");

            if (intCount > 0)
            {
                //YES
                Production = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Production", strWhere);

            }
            else
            {
                //NO - Production DOES NOT EXIST 
            }

            grdProduction.DataSource = Production;
            grdProduction.Refresh();

        }
        private void evaluateSurvey()
        {
            if (myConn.State == ConnectionState.Open)
            {
            }
            else
            {
                connectToDB();
            }

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "SURVEY");

            if (intCount > 0)
            {
                //YES

                string strSQL = "SELECT * from SURVEY " + strWhere;

                Survey = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                if (Survey.Rows.Count == 0)
                {
                    strSQL = "SELECT * from SURVEY " ;

                    Survey = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                }


                if (Survey.Rows.Count > 0)
                {
                    btnImportSurvey.Text = "Refresh Survey";

                    if (Survey.Columns.Contains("SECTION"))
                    {

                    }
                    else
                    {
                        strSQL = " Alter table Survey add SECTION varchar(50);";
                        TB.InsertData(Base.DBConnectionString, strSQL);

                        strSQL = " Update survey set SECTION = substring(Contract,2,3);";
                        TB.InsertData(Base.DBConnectionString, strSQL);
                    }

                    grdSurvey.DataSource = Survey;

                    DataTable Section = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                    if (Section.Rows.Count > 0)
                    {

                        //Extract distinct Sections
                        DataTable CntrWP = TB.createDataTableWithAdapter(Base.DBConnectionString, 
                                           "Select distinct Workplace from survey " + strWhere);

                        lstNames = TB.loadDistinctValuesFromColumn(Survey, "SECTION");
                        if (lstNames.Count > 1)
                        {

                            foreach (string s in lstNames)
                            {
                                if (listBox2.Items.Contains(s))
                                { }
                                else
                                {
                                    listBox2.Items.Add(s.Trim());
                                }
                            }
                        }

                        //Extract distinct Workplaces
                        lstNames = TB.loadDistinctValuesFromColumn(Survey, "WORKPLACE");

                        cboGangLinkWorkplace.Items.Clear();

                        if (lstNames.Count > 1)
                        {

                            foreach (string s in lstNames)
                            {
                                cboGangLinkWorkplace.Items.Add(s);
                            }

                            cboGangLinkWorkplace.Sorted = true;
                        }

                         
                    }
                    else
                    {
                        listBox2.Items.Add("No survey data exists");
                    }

                }
            }
            else
            {
                MessageBox.Show("Survey data does not exist.  Please import the data before trying to process.", "Information", MessageBoxButtons.OK);
            }

        }

        private void evaluateClockedShifts()
        {

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "ClockedShifts");

            if (intCount > 0)
            {
                //YES
                // Display die clocked info
                //amp
                if (strWhere.Trim().Length > 0)
                {
                    Clocked = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Clockedshifts", strWhere);
                }
                else
                {
                    Clocked = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Clockedshifts");

                }


                if (Clocked.Rows.Count > 0)
                {
                    foreach (DataColumn dc in Clocked.Columns)
                    {
                        if (dc.Caption.Substring(0, 3) == "DAY")
                        {
                            //cboOffDayValue.Items.Add(dc.Caption.Substring(4));
                            double d = Convert.ToDouble(dc.Caption.Substring(3).Trim());
                            string strTemp = Clocked.Rows[0]["FSH"].ToString().Trim();
                            DateTime temp = Convert.ToDateTime(strTemp);
                            temp = temp.AddDays(d);
                            Clocked.Columns[dc.Caption].ColumnName = Convert.ToString(temp.Day) + '-' + Convert.ToString(temp.Month);
                            //dc.Caption = Convert.ToString(temp.Day) + '-' + Convert.ToString(temp.Month);             
                        }
                    }
                }
                grdClocked.DataSource = Clocked;
                //amp
            }
            else
            {
                //MessageBox.Show("Clocked Shifts data does not exist.  Please import the data before trying to process.", "Information", MessageBoxButtons.OK);
            }

        }

        private void evaluateLabour()
        {

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "BONUSSHIFTS");

            if (intCount > 0)
            {

                string strSQL = "select * from BONUSSHIFTS " + strWhere;

                Labour = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                if (Labour.Rows.Count > 0)
                {
                    //amp
                    string strLSH = Labour.Rows[0]["LSH"].ToString().Trim();
                    DateTime LSH = Convert.ToDateTime(strLSH);
                    string Mnth = string.Empty;
                    string Day = string.Empty;

                    foreach (DataColumn dc in Labour.Columns)
                    {
                        if (dc.Caption.Substring(0, 3) == "DAY")
                        {
                            double d = Convert.ToDouble(dc.Caption.Substring(3).Trim());
                            string strTemp = Labour.Rows[0]["FSH"].ToString().Trim();
                            DateTime temp = Convert.ToDateTime(strTemp);
                            temp = temp.AddDays(d);
                            if (temp > LSH)  //remember the days start at 0
                            {
                                if (Convert.ToString(temp.Day).Length < 2)
                                {
                                    Day = "0" + Convert.ToString(temp.Day);
                                }
                                else
                                {
                                    Day = Convert.ToString(temp.Day);
                                }
                                if (Convert.ToString(temp.Month).Length < 2)
                                {
                                    Mnth = "0" + Convert.ToString(temp.Month);
                                }
                                else
                                {
                                    Mnth = Convert.ToString(temp.Month);
                                }
                                Labour.Columns[dc.Caption].ColumnName = "x" + Day + '-' + Mnth;
                            }
                            else
                            {
                                if (Convert.ToString(temp.Day).Length < 2)
                                {
                                    Day = "0" + Convert.ToString(temp.Day);
                                }
                                else
                                {
                                    Day = Convert.ToString(temp.Day);
                                }
                                if (Convert.ToString(temp.Month).Length < 2)
                                {
                                    Mnth = "0" + Convert.ToString(temp.Month);
                                }
                                else
                                {
                                    Mnth = Convert.ToString(temp.Month);
                                }
                                Labour.Columns[dc.Caption].ColumnName = "d" + Day + '-' + Mnth;
                            }
                        }
                    }
                }
                //amp}
                grdLabour.DataSource = Labour;

                string strNames = "select * from CLOCKEDSHIFTS";

                Workers = TB.createDataTableWithAdapter(Base.DBConnectionString, strNames);

                cboMinersEmpName.Items.Clear();

                foreach (DataRow data in Workers.Rows)
                {
                    cboMinersEmpName.Items.Add(data["EMPLOYEE_NAME"].ToString().Trim());
                }

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "EMPLOYEE_NAME");
                //cboMinersEmpName.Items.Clear();
                cboEmplPenEmployeeName.Items.Clear();

                foreach (string s in lstNames)
                {

                    //cboMinersEmpName.Items.Add(s.Trim());
                    cboEmplPenEmployeeName.Items.Add(s.Trim());

                }


                lstNames = TB.loadDistinctValuesFromColumn(Labour, "EMPLOYEE_No");
                //cboMinersEmpName.Items.Clear();
                cboEmplPenEmployeeNo.Items.Clear();

                foreach (string s in lstNames)
                {

                    //cboMinersEmpName.Items.Add(s.Trim());
                    cboEmplPenEmployeeNo.Items.Add(s.Trim());

                }

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "GANG");  //amp
                cboMinersGangNo.Items.Clear();
                cboBonusShiftsGang.Items.Clear();


                foreach (string s in lstNames)
                {
 
                    cboGangLinkGang.Items.Add(s.Trim());
                    cboMinersGangNo.Items.Add(s.Trim());
                    cboBonusShiftsGang.Items.Add(s.Trim());

                }    //amp

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "WAGECODE");  //amp
                cboBonusShiftsWageCode.Items.Clear();
                foreach (string s in lstNames)
                {

                    cboBonusShiftsWageCode.Items.Add(s.Trim());

                }    //amp

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "LINERESPCODE");  //amp
                cboBonusShiftsResponseCode.Items.Clear();
                foreach (string s in lstNames)
                {

                    cboBonusShiftsResponseCode.Items.Add(s.Trim());

                }    //amp
            }

            else
            {
                MessageBox.Show("Bonus Shifts data does not exist.  Please import the data before trying to process.", "Information", MessageBoxButtons.OK);
            }

            hideColumnsOfGrid("grdLabour");
        }

        private void hideColumnsOfGrid(string gridname)
        {

            switch (gridname)
            {
                case "grdMiners":
                    #region grdMiners
                    if (grdMiners.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdMiners.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdMiners.Columns.Contains("MININGTYPE"))
                    {
                        this.grdMiners.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdMiners.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdMiners.Columns["BONUSTYPE"].Visible = false;
                    }
                    return;
                    #endregion

                
                case "grdGangLink":
                    #region grdGangLink
                    if (grdGangLink.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdGangLink.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdGangLink.Columns.Contains("MININGTYPE"))
                    {
                        this.grdGangLink.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdGangLink.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdGangLink.Columns["BONUSTYPE"].Visible = false;
                    }
                    return;
                    #endregion

                case "grdSurvey":
                    #region grdSurvey
                    if (grdSurvey.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdSurvey.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdSurvey.Columns.Contains("MININGTYPE"))
                    {
                        this.grdSurvey.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdSurvey.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdSurvey.Columns["BONUSTYPE"].Visible = false;
                    }
                    return;
                    #endregion

                case "grdProduction":
                    #region grdProduction
                    if (grdProduction.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdProduction.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdProduction.Columns.Contains("MININGTYPE"))
                    {
                        this.grdProduction.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdProduction.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdProduction.Columns["BONUSTYPE"].Visible = false;
                    }
                    return;
                    #endregion

                case "grdLabour":
                    #region grdLabour
                    if (grdLabour.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdLabour.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdLabour.Columns.Contains("MININGTYPE"))
                    {
                        this.grdLabour.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdLabour.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdLabour.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;
                    #endregion

                case "grdRates":
                    #region grdRates
                    if (grdRates.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdRates.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdRates.Columns.Contains("MININGTYPE"))
                    {
                        this.grdRates.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdRates.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdRates.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;
                    #endregion

                case "grdCalendar":
                    #region grdCalendar
                    if (grdCalendar.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdCalendar.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdCalendar.Columns.Contains("MININGTYPE"))
                    {
                        this.grdCalendar.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdCalendar.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdCalendar.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;
                    #endregion

                case "grdActiveSheet":
                    #region grdActiveSheet
                    if (grdActiveSheet.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdActiveSheet.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdActiveSheet.Columns.Contains("MININGTYPE"))
                    {
                        this.grdActiveSheet.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdActiveSheet.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdActiveSheet.Columns["BONUSTYPE"].Visible = false;
                    }

                    break;
                    #endregion

                case "grdAbnormal":
                    #region grdAbnormal
                    if (grdAbnormal.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdAbnormal.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdAbnormal.Columns.Contains("MININGTYPE"))
                    {
                        this.grdAbnormal.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdAbnormal.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdAbnormal.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;
                    #endregion

                case "grdDrillers":
                    #region grdDrillers
                    if (grdDrillers.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdDrillers.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdDrillers.Columns.Contains("MININGTYPE"))
                    {
                        this.grdDrillers.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdDrillers.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdDrillers.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;
                    #endregion
            }
        }

        private void evaluateCalendar()
        {
            panel3.Enabled = true;
            panel4.Enabled = true;
            listBox2.Enabled = true;
            listBox3.Enabled = true;

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "CALENDAR");

            if (intCount > 0)
            {
                //Calendar exists,
                loadCalendar();
                loadDatePickers(0);
                loadSectionsFromCalendar();
            }
            else
            {
                createAndCopyCalendar();
            }
        }

        private void loadCalendar()
        {
            // Display die calendar info

            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar", " where period = '" + BusinessLanguage.Period + "'");

            grdCalendar.DataSource = Calendar;


        }

        private void loadStatus()
        {
            // Display die STATUS info

            Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere);
            if (Status.Rows.Count > 0)
            {
                statusChangeButtonColors();
            }
            else
            {
                Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status");
                string tempSection = Status.Rows[0]["SECTION"].ToString().Trim();
                DataTable temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "STATUS", "Where section = '" + tempSection + "'");
                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("BEGIN transaction; ");

                foreach (DataRow rr in temp.Rows)
                {
                    strSQL.Append("insert into Status values('" + rr["BUSSUNIT"].ToString().Trim() + "','" + rr["MININGTYPE"].ToString().Trim() +
                                    "','" + rr["BONUSTYPE"].ToString().Trim() + "','" + txtSelectedSection.Text +
                                  "','" + txtPeriod.Text.Trim() + "','" + rr["CATEGORY"].ToString().Trim() + "','" + rr["PROCESS"].ToString().Trim() +
                                  "','N','0');");

                }

                strSQL.Append("Commit Transaction;");
                TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
                Application.DoEvents();
                Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere);
            }
        }

        private void loadDatePickers(int Position)
        {
            //xxxxxxxxxxxxxxxx
            if (Calendar.Rows.Count > 0)
            {
                dateTimePicker1.Value = Convert.ToDateTime(Calendar.Rows[Position]["FSH"].ToString().Trim());
                dateTimePicker2.Value = Convert.ToDateTime(Calendar.Rows[Position]["LSH"].ToString().Trim());

                intNoOfDays = Base.calcNoOfDays(dateTimePicker2.Value, dateTimePicker1.Value);

                lstOffDayValue.Items.Clear();
                //Load the possible dates that the user can select in this measuring period for the offday calendar
                for (DateTime i = dateTimePicker1.Value; i <= dateTimePicker2.Value; i = i.AddDays(1))
                {
                    lstOffDayValue.Items.Add(i.ToString("yyyy-MM-dd"));
                }
            }
        }

        private void loadSectionsFromCalendar()
        {
            lstNames = TB.loadDistinctValuesFromColumn(Calendar, "SECTION");

            if (lstNames.Count > 0)
            {
                //HJ
                //txtSelectedSection.Text = "***";
                txtSelectedSection.Text = Calendar.Rows[0]["Section"].ToString().Trim();
                label15.Text = Calendar.Rows[0]["Section"].ToString().Trim();
                label30.Text = BusinessLanguage.Period;
                strWhere = "where section = '" + Calendar.Rows[0]["Section"].ToString().Trim() + "' and period = '" + BusinessLanguage.Period + "'";
                strWhereSection = "where section = '" + Calendar.Rows[0]["Section"].ToString().Trim() + "'";
                listBox2.Items.Clear();

                if (lstNames.Count > 1)
                {
                    foreach (string s in lstNames)
                    {
                        if (s != "XXX")
                        {
                            listBox2.Items.Add(s.Trim());
                        }
                    }
                }
                else
                {
                    if (lstNames.Count == 1)
                    {
                        foreach (string s in lstNames)
                        {
                            listBox2.Items.Add(s.Trim());
                        }
                    }
                }
            }
        }

        private void evaluateMiners()
        {
            // Display die Miners info
            Miners.Rows.Clear();

            loadMiners();

        }

        private void evaluateFactors()
        {
            // Display die Rates info
            Factors.Rows.Clear();

            loadFactors();

        }

        

        private void evaluateGangLinking()
        {
            // Display die Ganglink info
            GangLink.Rows.Clear();
            cboOffDaysGang.Items.Clear();

            loadGangLinking();

            lstNames = TB.loadDistinctValuesFromColumn(Labour, "GANG");
            //Load the distinct gang numbers into the ganglink listbox (lstGangs)
            lstGangs.Items.Clear();
            cboGangLinkGang.Items.Clear();
            if (lstNames.Count > 1)
            {
                foreach (string s in lstNames)
                {
                    cboGangLinkGang.Items.Add(s.Trim());
                    cboOffDaysGang.Items.Add(s.Trim());
                    lstGangs.Items.Add(s.Trim());
                }
            }

            cboGangLinkGangType.Text = "STOPING";

        }

        private void evaluatePayroll()
        {
            // Display die Ganglink info
            PayrollSends.Rows.Clear();

            loadPayroll();

        }

        private void loadPayroll()
        {
            //Check if Payroll exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");

            if (intCount > 0)
            {
                //YES
                PayrollSends = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Payroll", strWhere);
            }

            grdPayroll.DataSource = PayrollSends;
        }

        

        private void loadGangLinking()
        {
            //Check if ganglinking exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "GANGLINK");

            if (intCount > 0)
            {
                //YES
                GangLink = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "GangLink", strWhere);
                cboGangLinkGang.Items.Clear();
                List<string> lstGangs = TB.loadDistinctValuesFromColumn(Labour, "GANG");
                for (int i = 0; i <= lstGangs.Count - 1; i++)
                {
                    cboGangLinkGang.Items.Add(lstGangs[i].ToString().Trim());
                }
                 
            }
            else
            {
                //NO the ganglink table does not exist. 
                //Create the ganglink table
                //Check if BonusShifts Exists

                intCount = TB.checkTableExist(Base.DBConnectionString, "CLOCKEDSHIFTS");

                if (intCount > 0)
                {

                    //loadMonitor();
                    //TB.createGangLink(Base.DBConnectionString);
                    //TB.TBName = "GANGLINK";

                    //DialogResult result = MessageBox.Show("GangLink table does not exist or is corrupted. Do you want to recreate the table?", "Information", MessageBoxButtons.YesNo);

                    //switch (result)
                    //{
                    //    case DialogResult.Yes:

                    //        TB.createGangLink(Base.DBConnectionString);
                    //        return;

                    //    case DialogResult.No:
                    //        return;
                    //}

                    ////extractGangLinkData();
                    //saveXXXGangLink();

                }
                else
                {
                }

            }

            grdGangLink.DataSource = GangLink;

            hideColumnsOfGrid("grdGangLink");
        }

        private void evaluateRates()
        {
            // Display die Abnormal info
            Rates.Rows.Clear();

            loadRates();

        }

        private void evaluateEmployeePenalties()
        {
            // Display die EmployeePenalties info
            EmplPen.Rows.Clear();

            loadEmployeePenalties();

        }

        private void evaluateOffDays()
        {
            // Display die Offday Info
            Offdays.Rows.Clear();

            loadOffdays();
        }

        private void loadOffdays()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "OFFDAYS");

            if (intCount > 0)
            {
                //YES
                Offdays = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Offdays", strWhere);
                if (Offdays.Columns.Count > 0)
                {
                }
                else
                {
                    TB.insertOffdays(Base.DBConnectionString, "XXX", "XXX", "XXX");
                }

            }
            else
            {
                TB.createOffday(Base.DBConnectionString);
                TB.TBName = "OFFDAYS";
                Offdays = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Offdays");

            }

            grdOffDays.DataSource = Offdays;

            hideColumnsOfGrid("grdOffdays");
        }

        private void loadEmployeePenalties()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "EMPLOYEEPENALTIES");

            if (intCount > 0)
            {
                //YES

                EmplPen = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EMPLOYEEPENALTIES");

            }
            else
            {
                //NO
                //Check if Bonusshifts Exists

                intCount = TB.checkTableExist(Base.DBConnectionString, "BONUSSHIFTS");

                if (intCount > 0)
                {
                    TB.createEmployeePenalties(Base.DBConnectionString);
                    TB.TBName = "EMPLOYEEPENALTIES";
                    EmplPen = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EMPLOYEEPENALTIES ", strWhere);

                }
                else
                {
                }

            }

            grdEmplPen.DataSource = EmplPen;

            hideColumnsOfGrid("grdEmplPen");

        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            Application.Exit();

            //**********Old code

            //DialogResult result = MessageBox.Show("Have you saved your data? If not sure, please SAVE.", "REMINDER", MessageBoxButtons.YesNo);

            //switch (result)
            //{
            //    case DialogResult.Yes:
            //        this.Close();
            //        //scrMain main = new scrMain();
            //        //main.MainLoad(BusinessLanguage, DB, Survey, Labour, Miners, Designations, Occupations, Clocked, EmplList, EmplPen, Configs);
            //        //main.ShowDialog();
            //        myConn.Close();
            //        AAConn.Close();
            //        AConn.Close();
            //        this.Close();
            //        break;

            //    case DialogResult.No:
            //        break;
            //}

        }

        private void connectToDB()
        {

            if (myConn.State == ConnectionState.Closed)
            {
                try
                {
                    myConn.Open();
                }
                catch (SystemException eee)
                {
                    MessageBox.Show(eee.ToString());
                }
            }
        }

        private void btnImportSurvey_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DataTable temp = new DataTable();

            if (Status.Rows.Count > 0)
            {
                IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                              where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                              where locks.Field<string>("PROCESS").TrimEnd() == "tabSurvey"
                                              select locks;


                temp = query1.CopyToDataTable<DataRow>();
            }
            else
            {
                evaluateStatus();
                IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                              where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                              where locks.Field<string>("PROCESS").TrimEnd() == "tabSurvey"
                                              select locks;


                temp = query1.CopyToDataTable<DataRow>();




            }
            if (temp.Rows[0]["STATUS"].ToString().Trim() == "N")
            {
                refreshSurvey();

            }
            else
            {
                MessageBox.Show("Production is locked. Unlock before refresh", "Information", MessageBoxButtons.OK);
            }

            this.Cursor = Cursors.Arrow;

        }

        private void refreshSurvey()
        {
            //Fire the sql to import the production data.

            bool XLSX_exists = false;
            bool XLS_exists = false;
            string FilePath_XLSX = "C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\productionstp_" + BusinessLanguage.Period.Trim() + ".xlsx";
            string FilePath_XLS = "C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\productionstp_" + BusinessLanguage.Period.Trim() + ".xls";
            string FilePath = "";

            #region extract the sheet name and FSH and LSH of the extract
            XLSX_exists = File.Exists(FilePath_XLSX);
            XLS_exists = File.Exists(FilePath_XLS);

            if (XLS_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\productionstp_" + BusinessLanguage.Period.Trim() + ".xls";
            }
            else
            {
                if (XLSX_exists.Equals(true))
                {
                    FilePath = "C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\productionstp_" + BusinessLanguage.Period.Trim() + ".xlsx";
                }
            }
            bool test = File.Exists(FilePath);
            string[] sheetNames = GetExcelSheetNames(FilePath);
            string sheetName = sheetNames[0];
            #endregion

            #region import Production

            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();

            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 8.0;'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            da = new OleDbDataAdapter("select * from [" + sheetName + "]", con);


            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                
                #region Change the column names
                //Change the column names to the correct column names.
                Dictionary<string, string> dictNames = new Dictionary<string, string>();
                DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                     "Select * from varnames");
                dictNames.Clear();

                dictNames = TB.loadDict(varNames, dictNames);

                //If it is a column with a date as a name.
                foreach (DataColumn column in dt.Columns)
                {
                    if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                    {
                        column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                    }

                }

              
                ////Add the extra columns
                dt.Columns.Add("MININGTYPE");
                dt.Columns.Add("BONUSTYPE");
                dt.Columns.Add("SECTION");
                dt.Columns.Add("FSH");
                dt.Columns.Add("LSH"); 
                dt.AcceptChanges();

                //Replace all the columns containing nulls to '-'
                foreach (DataRow row in dt.Rows)
                {
                    if (string.IsNullOrEmpty(row["PENALTYMETERSIND"].ToString()))
                    {
                        row["PENALTYMETERSIND"] = "0";
                    }
                    row["MININGTYPE"] = "STOPE";
                    row["BONUSTYPE"] = "TEAM";
                    row["FSH"] = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    row["LSH"] = dateTimePicker2.Value.ToString("yyyy-MM-dd");

                    if (row["CONTRACT"].ToString().Length > 0)
                    {
                        row["SECTION"] = row["CONTRACT"].ToString().Substring(0, 3);
                    }
                    else
                    {
                        row["SECTION"] = "XXX";
                    }

                    row["STOPETYPE"] = row["STOPETYPE"].ToString().ToUpper();
                    row["FSH"] = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    row["LSH"] = dateTimePicker2.Value.ToString("yyyy-MM-dd");

                    for (int i = 0; i <= dt.Columns.Count - 1; i++)
                    {
                        if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                        {
                            row[i] = "0";
                        }
                    }
                }
                #endregion

                #region save to production database
                int counter = 1;
                foreach (DataRow kr in dt.Rows)
                {
                    kr["WORKPLACETYPE"] = counter;
                    counter++;

                }

                #region save to production database
                //save to the  PRODUCTION database
                TB.saveCalculations2(dt, Base.DBConnectionString, strWhere, "PRODUCTION");

                Application.DoEvents();

                grdProduction.DataSource = dt;

                Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PRODUCTION");

                #endregion

                if (intCount > 0)
                {
                    this.Cursor = Cursors.WaitCursor;
                    btnx.Visible = true;
                    btnx.Enabled = true;
                    btnx.Text = "Run";
                    TB.deleteProcess(Base.AnalysisConnectionString, Base.DBName + BusinessLanguage.Period);
                    //clear the monitor table
                    TB.deleteAllExcept(Base.DBConnectionString, "Monitor");
                    Calcs("Production", "ProductionEarn10", "Y");
                    Calcs("Production", "ProductionEarn20", "N");
                    Calcs("Production", "ProductionEarn30", "N");
                    Calcs("Production", "ProductionEarn40", "N");
                    Calcs("Production", "ProductionEarn50", "N");
                    Calcs("Exit", "Exit", "N");
                    //Remember to add the calcnames manually to the survey file.
                    btnx.Visible = true;
                    btnx.Enabled = true;
                    btnx.Text = "Run";

                    btnx_Click_1("Method", null);
                }
                #endregion

            }
            else
            {
                MessageBox.Show("No rows for section: " + txtSelectedSection.Text.Trim() + " were found on the spreadsheet.", "Information", MessageBoxButtons.OK);
            }
            #endregion

        }

        private void loadRates()
        {
            //Check if ABNORMAL exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Rates");

            if (intCount > 0)
            {
                //YES

                Rates = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Rates"," where period = '" + BusinessLanguage.Period + "'");

            }
            else
            {
                //NO - Rates DOES NOT EXIST 
            }

            grdRates.DataSource = Rates;

            hideColumnsOfGrid("grdRates");

        }

        private void loadMiners()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "MINERS");

            if (intCount > 0)
            {
                //YES

                Miners = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Miners ", strWhere);


            }

            grdMiners.DataSource = Miners;

            hideColumnsOfGrid("grdMiners");

        }

        private void extractMinersData()
        {
            string strSQL = "select SECTION, PERIOD, WORKPLACE, 'XXX' as EMPLOYEE_NO, '1' as DESIGNATION, '0' as SHIFTS_WORKED," +  
                            "'0' as AWOP_SHIFTS,'0' as PAYSHIFTS,'0' as SAFETYIND " +
                            "from Survey  where period = '" + BusinessLanguage.Period + "'";

            DataTable tempDataTable = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

            foreach (DataRow _row in tempDataTable.Rows)
            {
                if (string.IsNullOrEmpty(_row[0].ToString()))
                {
                }
                else
                {
                    Miners.Rows.Add(_row.ItemArray);
                }
            }

            saveXXXMiners();

        }

        private void importTheSheet(string importFilename)
        {
            string path = BusinessLanguage.InputDirectory + Base.DBName;

            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                string filename = BusinessLanguage.InputDirectory + Base.DBName + importFilename;
                bool fileCheck = BusinessLanguage.checkIfFileExists(filename);

                if (fileCheck)
                {
                    FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                    spreadsheet = new ExcelDataReader.ExcelDataReader(fs);
                    fs.Close();
                    //If the file was SURVEY, all sections production data will be on this datatable.
                    //Only the selected section's data must be saved.

                    saveTheSpreadSheetToTheDatabase();
                }
                else
                {
                    MessageBox.Show("File " + filename + " - does not exist", "Check", MessageBoxButtons.OK);
                }

                //Check if file exists
                //If not  = Message
                //If exists ==>  Import
            }
            catch
            {
                MessageBox.Show("File " + importFilename + " - is inuse by another package?", "Check", MessageBoxButtons.OK);
            }
        }

        private void saveTheSpreadSheetToTheDatabase()
        {
            foreach (DataTable dt in spreadsheet.WorkbookData.Tables)
            {
                if (dt.TableName == "SURVEY" || dt.TableName == "Survey")
                {
                    for (int i = 1; i <= dt.Rows.Count - 1; i++)
                    {
                        if (dt.Rows[i][3].ToString().Trim() == txtSelectedSection.Text.Trim())
                        {
                        }
                        else
                        {
                            dt.Rows[i].Delete();

                        }
                    }

                }

                dt.AcceptChanges();
                //checker = true;

                TB.TBName = dt.TableName.ToString().ToUpper();
                TB.recreateDataTable();

                //Extract column names
                string strColumnHeadings = TB.getFirstRowValues(dt, Base.AnalysisConnectionString);

                switch (strColumnHeadings)
                {
                    case null:
                        break;

                    case "":
                        break;

                    default:


                        if (myConn.State == ConnectionState.Closed)
                        {
                            try
                            {
                                myConn = Base.DBConnection;
                                myConn.Open();

                                //create a table
                                bool tableCreate = TB.createDatabaseTable(Base.DBConnectionString, strColumnHeadings);

                                tableCreate = TB.copySpreadsheetToDatabaseTable(Base.DBConnectionString, dt);

                                if (tableCreate)
                                {
                                    MessageBox.Show("Data successfully imported", "Information", MessageBoxButtons.OK);
                                }
                                else
                                {
                                    MessageBox.Show("Try again after correction of spreadsheet - input data.", "Information", MessageBoxButtons.OK);
                                }

                                //checker = false;
                            }
                            catch (System.Exception ex)
                            {
                                System.Windows.Forms.MessageBox.Show(ex.GetHashCode() + " " + ex.ToString(), "MyProgram", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            //create a table
                            bool tableCreate = TB.createDatabaseTable(Base.DBConnectionString, strColumnHeadings);

                            if (tableCreate)
                            {
                                tableCreate = TB.copySpreadsheetToDatabaseTable(Base.DBConnectionString, dt);
                                MessageBox.Show("Data successfully imported", "Information", MessageBoxButtons.OK);

                            }
                            else
                            {
                                MessageBox.Show("Data was not imported.", "Information", MessageBoxButtons.OK);
                            }
                        }

                        break;
                }
            }
        }

        //private void saveXXXAbnormal()
        //{
        //    StringBuilder strSQL = new StringBuilder();
        //    strSQL.Append("BEGIN transaction; ");

        //    #region tabAbnormal
        //    foreach (DataRow rr in Abnormal.Rows)
        //    {

        //        strSQL.Append("insert into Abnormal values('" + rr["SECTION"].ToString().Trim() + "','" + rr["PERIOD"].ToString().Trim() +
        //                      "','" + rr["CONTRACT"].ToString().Trim() + "','" + rr["WORKPLACE"].ToString().Trim() + "','" +
        //                      rr["ABNORMALLEVEL"].ToString().Trim() + "','" + rr["ABNORMALTYPE"].ToString().Trim() + "','" +
        //                      rr["ABNORMALVALUE"].ToString().Trim() + "');");
        //    }

        //    strSQL.Append("Commit Transaction;");
        //    TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
        //    #endregion

        //}



        private void saveXXXMiners()
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");
            string coy = "";
            string designation = "";

            #region tabMiners
            foreach (DataRow rr in Miners.Rows)
            {
                if (rr["EMPLOYEE_NO"].ToString().Trim().Contains("-"))
                {

                    coy = rr["EMPLOYEE_NO"].ToString().Substring(0, rr["EMPLOYEE_NO"].ToString().IndexOf("-")).Trim();
                }
                else
                {
                    coy = rr["EMPLOYEE_NO"].ToString().Trim();
                }

                if (rr["DESIGNATION"].ToString().Contains("-"))
                {
                    designation = rr["DESIGNATION"].ToString().Substring(0, rr["DESIGNATION"].ToString().IndexOf("-")).Trim();
                }
                else
                {
                    designation = rr["DESIGNATION"].ToString().Trim();
                }

                string test = rr["EMPLOYEE_NO"].ToString().Trim();

                strSQL.Append("insert into Miners values('" + rr["SECTION"].ToString().Trim() + "','" + rr["PERIOD"].ToString().Trim() +
                              "','xxx','" + coy + "','" + designation +
                              "','" + rr["PAYSHIFTS"].ToString().Trim() + "','" + rr["AWOP_SHIFTS"].ToString().Trim() +
                              "','" + rr["SAFETYIND"].ToString().Trim() + "');");
            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            #endregion

        }

        private void saveXXXGangLink()
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            #region tabGangLink
            foreach (DataRow rr in GangLink.Rows)
            {

                strSQL.Append("insert into GANGLINK values('" + rr["SECTION"].ToString().Trim() + "','" + rr["PERIOD"].ToString().Trim() + "','" +
                                rr["GANG"].ToString().Trim() + "','" + rr["WORKPLACE"].ToString().Trim() + "','" +
                                rr["SAFETYIND"].ToString().Trim() + "','" + rr["GANGTYPE"].ToString().Trim() + "','" + rr["CREWNO"].ToString().Trim() + "','"
                                + rr["GANGSHIFTS"].ToString().Trim() + "','0');");
            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            #endregion

        }

        public String[] GetExcelSheetNames(string excelFile)
        {
            //MessageBox.Show(excelFile);
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {
                // Connection String. Change the excel file to the file you
                // will search.
                String connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + excelFile + ";Extended Properties=Excel 8.0;";
                // Create connection object by using the preceding connection string.
                objConn = new OleDbConnection(connString);
                // Open connection with the database.
                objConn.Open();
                // Get the data table containg the schema guid.
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {

                    //MessageBox.Show(row["TABLE_NAME"].ToString());
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    // Query each excel sheet.
                }

                return excelSheets;
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message);
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        private int checkLockCalendarProcesses()
        {

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "N"
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          where locks.Field<string>("PROCESS").TrimEnd() == "tabCalendar"
                                          where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period
                                          select locks;

            try
            {
                int intcount = query1.Count<DataRow>();

                return intcount;
            }
            catch
            {
                MessageBox.Show("Error in checkLockCalendarProcess.");
                return 0;
            }

            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();

        }

        private void btnImportADTeam_Click(object sender, EventArgs e)
        {
            DataTable temp = new DataTable();

            int intCalendarProcesses = checkLockCalendarProcesses();

            if (intCalendarProcesses > 0)
            {
                MessageBox.Show("Please finalize Calendar before importing your shifts.");
            }
            else
            {
                if (Labour.Rows.Count > 0)
                {
                    IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                                  where locks.Field<string>("PROCESS").TrimEnd() == "tabLabour"
                                                  where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                                  where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period.Trim()
                                                  select locks;

                    try
                    {
                        temp = query1.CopyToDataTable<DataRow>();
                        loadDatePickers(0);
                        if (intNoOfDays <= 45)
                        {
                            refreshLabour();
                        }
                        else
                        {
                            MessageBox.Show("Shifts cannot be imported.  Please fix the shifts on calendar.", "Information",
                                MessageBoxButtons.OK);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("No records on Status for the Section,Period and tabLabour");
                    }
                }
                else
                {
                    evaluateStatus();
                    IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                                  where locks.Field<string>("PROCESS").TrimEnd() == "tabLabour"
                                                  where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                                  where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period.Trim()
                                                  select locks;

                    try
                    {
                        temp = query1.CopyToDataTable<DataRow>();
                        if (temp.Rows.Count > 0)
                        {
                            if (temp.Rows[0]["STATUS"].ToString().Trim() == "N")
                            {

                                loadDatePickers(0);
                                if (intNoOfDays <= 45)
                                {
                                    refreshLabour();
                                }
                                else
                                {
                                    MessageBox.Show("Shifts cannot be imported.  Please fix the shifts on calendar.", "Information",
                                        MessageBoxButtons.OK);
                                }

                            }
                            else
                            {
                                MessageBox.Show("BonusShifts is locked. Please unlock before refresh.  You WILL loose all previous updates.",
                                    "Information", MessageBoxButtons.OK);
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Could not find LOCK records on STATUS for selected period.  Please contract iCalc.", "Information", MessageBoxButtons.OK);
                    }
                }

            }
        }

        private void refreshLabour2()
        {
            #region extract the  FSH from the database
            this.Cursor = Cursors.WaitCursor;
            extractMeasuringDates();
            //This is the refresh from the ADTeam database.
            SqlConnection _ADTeamConn = new SqlConnection();

            _ADTeamConn = Base.ADTeamConnection;
            _ADTeamConn.Open();

            DataTable ADTeam = TB.createDataTableWithAdapter(Base.ADTeamConnectionString, "select TOP 1 *  from FREEGOLD_EMPLOYEEDETAIL");
            DateTime _lastRunDate = Convert.ToDateTime(ADTeam.Rows[0]["lastrundate"]);

            int intNoOfDays = Base.calcNoOfDays(dateTimePicker2.Value, dateTimePicker1.Value);
            int intStart = Base.calcNoOfDays(_lastRunDate, dateTimePicker1.Value) + 1;

            if (intStart > 100)
            {
                intStart = 100;
            }

            int intEnd = intStart - intNoOfDays;

            if (intEnd <= 0)
            {
                intEnd = 1;
            }


            DataTable dt = TB.ExtractADTeamShifts(Base.ADTeamConnectionString, intStart, intEnd, dateTimePicker1.Value,
                                                  dateTimePicker2.Value, intStart, intEnd,
                                                  BusinessLanguage.Period, txtSelectedSection.Text.Trim(), BusinessLanguage.MiningType,
                                                  BusinessLanguage.BonusType, BusinessLanguage.BussUnit, " where bussunit = 'JB' " +
                                                  " and substring([Gang Name],1,4) IN " + strMO);


            foreach (DataRow row in dt.Rows)
            {

                row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }


            MessageBox.Show("Save Clocked shifts");

            string tst = string.Empty;
            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                tst = tst.Trim() + "-" + dt.Columns[i].ColumnName.Trim();
            }

            MessageBox.Show(intStart.ToString().Trim() + "-" + intEnd.ToString().Trim() + "-" + intNoOfDays.ToString().Trim() + tst);

            string strDelete = " where section = '" + txtSelectedSection.Text.Trim() +
                               "' and period = '" + BusinessLanguage.Period.Trim() + "'";

            TB.saveCalculations2(dt, Base.DBConnectionString, strDelete, "CLOCKEDSHIFTS");

            MessageBox.Show("Clocked shifts were saved!");

            //========================================================================
            //string tst = string.Empty;
            //for (int i = 0; i <= dt.Columns.Count - 1; i++)
            //{
            //    tst = tst.Trim() + "-" + dt.Columns[i].ColumnName.Trim();
            //}

            //MessageBox.Show(intStart.ToString().Trim() + "-" + intEnd.ToString().Trim() + "-" + intNoOfDays.ToString().Trim() + "-" + tst);

            #endregion

            #region Apply offdays
            if (dt.Rows.Count > 0)
            {
                Clocked = dt.Copy();
                //Update clockedshifts with offday calendar data
                UpdateClockedShifts();
                dt = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Clockedshifts");

                Application.DoEvents();

                #region Calculate the shifts per employee en output to bonusshifts

                string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
                                "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where section = '" +
                                txtSelectedSection.Text.Trim() + "'";

                string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";

                fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix);
                BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

                DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
                DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

                //sheetfhs = SheetFSH;
                //sheetlhs = SheetLSH;
                //int intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
                //int intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
                int intStopDay = 0;

                if (intStartDay < 0)
                {
                    //The calendarFSH falls outside the startdate of the sheet.
                    intStartDay = 0;
                }
                else
                {

                }

                if (intEndDay < 0 && intEndDay < -44)
                {
                    intStopDay = 0;
                }
                else
                {
                    if (intEndDay < 0)
                    {
                        //the LSH of the measuring period falls within the spreadsheet
                        intStopDay = intNoOfDays + intEndDay;

                    }
                    else
                    {
                        //The LSH of the measuring period falls outside the spreadsheet
                        intStopDay = 44;
                    }

                    //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                    //were not imported.

                    #region count the shifts
                    //Count the shifts

                    DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                    switch (result)
                    {
                        case DialogResult.OK:
                            extractAndCalcShifts(0, intNoOfDays);
                            MessageBox.Show("Shifts were imported successfully", "Information", MessageBoxButtons.OK);
                            break;

                        case DialogResult.Cancel:
                            MessageBox.Show("No changes was made!", "Information", MessageBoxButtons.OK);
                            break;

                    }

                    #endregion

                #endregion

                    this.Cursor = Cursors.Arrow;


                }
            }
            else
            {
                MessageBox.Show("No shifts were imported. Please check the parameters for the section.", "Information", MessageBoxButtons.OK);
                this.Cursor = Cursors.Arrow;

            }
            #endregion
        }

        private void refreshLabour()
        {
            #region extract the sheet name and FSH and LSH of the extract
            ATPMain.VkExcel excel = new ATPMain.VkExcel(false);


            bool XLSX_exists = File.Exists("C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx");
            bool XLS_exists = File.Exists("C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls");

            if (XLSX_exists.Equals(true))
            {
                //MessageBox.Show("nou in xlsx filepath");
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx", "");

                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            if (XLS_exists.Equals(true))
            {
                //MessageBox.Show("nou in xls filepath");
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls", "");

                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                //MessageBox.Show("file is saved");
                excel.CloseFile();
                //MessageBox.Show("file is closed");
            }

            excel.stopExcel();

            string FilePath = "";

            string FilePath_XLS = "C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";
            string FilePath_XLSX = "C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";

            XLSX_exists = File.Exists(FilePath_XLSX);
            XLS_exists = File.Exists(FilePath_XLS);

            if (XLS_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";
                //MessageBox.Show("gebruik die xls filepath");
            }
            if (XLSX_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Tshepong\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";
                //MessageBox.Show("gebruik die xlsx filepath");
            }
            #endregion

            if (FilePath.Trim().Length > 0)
            {

                #region Read Sheets
                //MessageBox.Show("gaan nou die sheetnames kry" + FilePath);
                string[] sheetNames = GetExcelSheetNames(FilePath);
                //MessageBox.Show("sheetnames gelees vanaf" + FilePath);
                string sheetName = sheetNames[0];
                
                string testString = sheetName.Substring(0, 3).ToString().Trim();


                //if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
                //{
                //    sheetName = sheetNames[1];
                //}

                //if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
                //{
                //    sheetName = sheetNames[2];
                //}

                //if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
                //{
                //    sheetName = sheetNames[3];
                //}
                #endregion

                #region import Clockshifts
                this.Cursor = Cursors.WaitCursor;
                DataTable dt = new DataTable();
                OleDbConnection con = new OleDbConnection();
                OleDbDataAdapter da;
                //MessageBox.Show("FilePath" + FilePath);
                con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                        + FilePath + ";Extended Properties='Excel 8.0;'";

                /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
                * "HDR=No;" indicates the opposite.
                * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
                * Note that this option might affect excel sheet write access negative.
                */

                da = new OleDbDataAdapter("select * from [" + sheetName + "] where MID([GANG NAME],1,4) IN " + strMO, con);

                da.Fill(dt);


                #endregion

                #region remove invalid records

                //extract the column names with length less than 3.  These columns must be deleted.
                string[] columnNames = new String[dt.Columns.Count];

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (dt.Columns[i].ColumnName.Length <= 2)
                    {
                        columnNames[i] = dt.Columns[i].ColumnName;
                    }
                }

                for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
                {
                    if (string.IsNullOrEmpty(columnNames[i]))
                    {

                    }
                    else
                    {
                        dt.Columns.Remove(columnNames[i].ToString().Trim());
                        dt.AcceptChanges();
                    }
                }

                dt.Columns.Remove("INDUSTRY NUMBER");
                dt.AcceptChanges();
                #endregion

                #region process spreadsheet

                string strSheetFSH = string.Empty;
                string strSheetLSH = string.Empty;

                //Extract the dates from the spreadsheet - the name of the spreadsheet contains the start and enddate of the extract
                string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
                string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

                //Correct the dates and calculate the number of days extracted.
                if (strSheetFSHx.Substring(6, 1) == "-")
                {
                    strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
                }
                else
                {
                    strSheetFSH = strSheetFSHx;
                }

                if (strSheetLSHx.Substring(6, 1) == "-")
                {
                    strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
                }
                else
                {
                    strSheetLSH = strSheetLSHx;
                }

                DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
                DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

                //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
                int intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);

                if (intNoOfDays <= 44)
                {
                    for (int j = intNoOfDays + 1; j <= 44; j++)
                    {
                        dt.Columns.Add("DAY" + j);
                    }
                }
                else
                {

                }

                #region Change the column names
                //Change the column names to the correct column names.
                Dictionary<string, string> dictNames = new Dictionary<string, string>();
                DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                     "Select * from varnames");
                dictNames.Clear();

                dictNames = TB.loadDict(varNames, dictNames);
                int counter = 0;

                //If it is a column with a date as a name.
                foreach (DataColumn column in dt.Columns)
                {
                    if (column.ColumnName.Substring(0, 1) == "2")
                    {
                        if (counter == 0)
                        {
                            strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            if (column.Ordinal == dt.Columns.Count - 1)
                            {

                                column.ColumnName = "DAY" + counter;
                                counter = counter + 1;

                            }
                            else
                            {
                                column.ColumnName = "DAY" + counter;
                                counter = counter + 1;
                            }
                        }


                    }
                    else
                    {
                        if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                        {
                            column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                        }

                    }
                }

                //Add the extra columns
                dt.Columns.Add("BUSSUNIT");
                dt.Columns.Add("FSH");
                dt.Columns.Add("LSH");
                dt.Columns.Add("SECTION");
                dt.Columns.Add("EMPLOYEETYPE");
                dt.Columns.Add("PERIOD");      //xxxxxxxx
                dt.AcceptChanges();

                foreach (DataRow row in dt.Rows)
                {
                    row["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                    row["FSH"] = strSheetFSH;
                    row["LSH"] = strSheetLSH;
                    row["MININGTYPE"] = "STOPE";
                    row["PERIOD"] = BusinessLanguage.Period;   //xxx
                    if (row["GANG"].ToString().Length > 0)
                    {
                        row["SECTION"] = txtSelectedSection.Text.Trim();
                    }
                    else
                    {
                        row["SECTION"] = "XXX";
                    }
                    if (row["WAGECODE"].ToString().Trim() == "")
                    {
                        row["WAGECODE"] = "00000";
                    }
                    else
                    {
                    }
                    row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                    for (int i = 0; i <= dt.Columns.Count - 1; i++)
                    {
                        if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                        {
                            row[i] = "-";
                        }
                    }
                }

                //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
                DataColumn dcBussunit = new DataColumn();
                dcBussunit.ColumnName = "BUSSUNIT";
                dt.Columns.Remove("BUSSUNIT");
                dt.AcceptChanges();
                InsertAfter(dt.Columns, dt.Columns["BONUSTYPE"], dcBussunit);

                foreach (DataRow dr in dt.Rows)
                {
                    dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                }


                #endregion

                #endregion

                #region write Clockedshifts
                //Write to the database

                TB.saveCalculations2(dt, Base.DBConnectionString, "", "CLOCKEDSHIFTS");
                if (dt.Rows.Count > 0)
                {
                    Clocked = dt.Copy();
                    //Update clockedshifts with offday calendar data
                    UpdateClockedShifts();
                    dt = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Clockedshifts");

                    Application.DoEvents();

                    grdClocked.DataSource = dt;
                #endregion

                #region Calculate the shifts per employee en output to bonusshifts

                    string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
                                    "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where section = '" +
                                    txtSelectedSection.Text.Trim() + "'";

                    string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";

                    fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix); 
                    BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
                    //exportToExcel("c:\\", BonusShifts);
                    string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

                    DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
                    DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

                    sheetfhs = SheetFSH;
                    sheetlhs = SheetLSH;
                    int intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
                    int intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
                    int intStopDay = 0;

                    if (intStartDay < 0)
                    {
                        //The calendarFSH falls outside the startdate of the sheet.
                        intStartDay = 0;
                    }
                    else
                    {

                    }

                    if (intEndDay < 0 && intEndDay < -44)
                    {
                        intStopDay = 0;
                    }
                    else
                    {
                        if (intEndDay < 0)
                        {
                            //the LSH of the measuring period falls within the spreadsheet
                            intStopDay = intNoOfDays + intEndDay;

                        }
                        else
                        {
                            //The LSH of the measuring period falls outside the spreadsheet
                            intStopDay = 44;
                        }

                        //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                        //were not imported.

                        #region count the shifts
                        //Count the shifts

                        DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                        switch (result)
                        {
                            case DialogResult.OK:
                                extractAndCalcShifts(intStartDay, intStopDay);
                                MessageBox.Show("Shifts were imported successfully", "Information", MessageBoxButtons.OK);
                                break;

                            case DialogResult.Cancel:
                                MessageBox.Show("No changes was made!", "Information", MessageBoxButtons.OK);
                                break;

                        }

                        #endregion

                    #endregion

                        this.Cursor = Cursors.Arrow;
                        File.Delete(FilePath);

                    }
                }
                else
                {
                    MessageBox.Show("No shifts were imported. Please check the parameters for the section.", "Information", MessageBoxButtons.OK);
                    this.Cursor = Cursors.Arrow;
                    File.Delete(FilePath);
                }
                }
            else
            {
                 MessageBox.Show("ADTEAM file does not exist.", "Information", MessageBoxButtons.OK);

            }
        }

        private void extractAndCalcShifts(int DayStart, int DayEnd)
        {
            int intSubstringLength = 0;
            int intShiftsWorked = 0;
            int intAwopShifts = 0;
            int shiftsCheck = 0;
            BonusShifts.Columns.Add("TMLEADERIND");

            foreach (DataRow row in BonusShifts.Rows)
            {
                foreach (DataColumn column in BonusShifts.Columns)
                {
                    if ((column.ColumnName.Substring(0, 3) == "DAY"))
                    {
                        if (column.ColumnName.ToString().Length == 4)
                        {
                            intSubstringLength = 1;
                        }
                        else
                        {
                            intSubstringLength = 2;
                        }

                        if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                           Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
                        {
                            if (row[column].ToString().Trim() == "U" || row[column].ToString().Trim() == "u" ||
                                row[column].ToString().Trim() == "q" || row[column].ToString().Trim() == "r" ||
                                row[column].ToString().Trim() == "Q" || row[column].ToString().Trim() == "W" ||
                                row[column].ToString().Trim() == "w")
                            {
                                intShiftsWorked = intShiftsWorked + 1;
                                shiftsCheck = 1;
                            }
                            else
                            {
                                //if (row[column].ToString().Trim() == "A" || row[column].ToString().Trim() == "b")
                                if (row[column].ToString().Trim() == "A")
                                {
                                    intAwopShifts = intAwopShifts + 1;
                                }
                                else { }

                            }
                        }
                        else
                        {
                            row[column] = "*";
                        }
                    }
                    else
                    {
                        if (column.ColumnName == "BONUSTYPE")
                        {
                            row["BONUSTYPE"] = "TEAM";
                        }
                    }
                }//foreach datacolumn

                //If shifts_worked > monthsshifts then employee_shifts = monthshifts
                if (Convert.ToInt16(intShiftsWorked) > Convert.ToInt16(strMonthShifts))
                {
                    row["SHIFTS_WORKED"] = strMonthShifts;
                }
                else
                {
                    row["SHIFTS_WORKED"] = Convert.ToString(intShiftsWorked);
                }

                row["AWOP_SHIFTS"] = intAwopShifts;
                row["TMLEADERIND"] = "0";
                intShiftsWorked = 0;
                intAwopShifts = 0;
            }
            //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
            DataColumn dcPeriod = new DataColumn();
            dcPeriod.ColumnName = "PERIOD";
            BonusShifts.Columns.Remove("PERIOD");
            BonusShifts.AcceptChanges();
            InsertAfter(BonusShifts.Columns, BonusShifts.Columns["BONUSTYPE"], dcPeriod);

            foreach (DataRow dr in BonusShifts.Rows)
            {
                dr["PERIOD"] = BusinessLanguage.Period;
            }


             string strDelete = " where section = '" + txtSelectedSection.Text.Trim() +
                               "' and period = '" + BusinessLanguage.Period.Trim() + "'";

             TB.saveCalculations2(BonusShifts, Base.DBConnectionString, strDelete, "BONUSSHIFTS");

            if (importdone == 0)
            {

                fillFixTable(fixShifts, sheetfhs, sheetlhs, noOFDay, DayStart, DayEnd);//Calls the method to load the fix clockedshiftstable
                importdone = 1;

            }

            Application.DoEvents();
        }

        public void InsertAfter(DataColumnCollection columns, DataColumn currentColumn, DataColumn newColumn)
        {
            if (columns.Contains(currentColumn.ColumnName))
            {
                columns.Add(newColumn);
                //add the new column after the current one 
                columns[newColumn.ColumnName].SetOrdinal(currentColumn.Ordinal + 1);
            }
            else
            {
                throw new ArgumentException(/** snip **/);
            }
        }

        //private void extractAndCalcShifts(int DayStart, int DayEnd)
        //{
        //    int intSubstringLength = 0;
        //    int intShiftsWorked = 0;
        //    int intAwopShifts = 0;
        //    int shiftsCheck = 0;
        //    BonusShifts.Columns.Add("TMLEADERIND");

        //    foreach (DataRow row in BonusShifts.Rows)
        //    {
        //        foreach (DataColumn column in BonusShifts.Columns)
        //        {
        //            if ((column.ColumnName.Substring(0, 3) == "DAY"))
        //            {
        //                if (column.ColumnName.ToString().Length == 4)
        //                {
        //                    intSubstringLength = 1;
        //                }
        //                else
        //                {
        //                    intSubstringLength = 2;
        //                }

        //                if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
        //                   Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
        //                {
        //                    if (row[column].ToString().Trim() == "U" || row[column].ToString().Trim() == "u" || row[column].ToString().Trim() == "q" || row[column].ToString().Trim() == "Q" || row[column].ToString().Trim() == "W" || row[column].ToString().Trim() == "w")
        //                    {
        //                        intShiftsWorked = intShiftsWorked + 1;
        //                        shiftsCheck = 1;
        //                    }
        //                    else
        //                    {
        //                        if (row[column].ToString().Trim() == "A")
        //                        {
        //                            intAwopShifts = intAwopShifts + 1;
        //                        }
        //                        else { }

        //                    }
        //                }
        //                else
        //                {
        //                    row[column] = "*";
        //                }
        //            }
        //            else
        //            {
        //                if (column.ColumnName == "BONUSTYPE")
        //                {
        //                    row["BONUSTYPE"] = "TEAM";
        //                }
        //            }
        //        }//foreach datacolumn

        //        row["SHIFTS_WORKED"] = intShiftsWorked;
        //        row["AWOP_SHIFTS"] = intAwopShifts;
        //        row["TMLEADERIND"] = "0";
        //        intShiftsWorked = 0;
        //        intAwopShifts = 0;
        //    }

        //    //exportToExcel("C:\\",BonusShifts);
        //    //if (strWhere.Contains(" and gang = '"))
        //    //{
        //    //    strWhere = strWhere.Trim().Substring(0, strWhere.Trim().IndexOf("and"));
        //    //}

        //    TB.saveCalculations2(BonusShifts, Base.DBConnectionString, strWhere, "BONUSSHIFTS");

        //    if (importdone == 0)
        //    {

        //        fillFixTable(fixShifts, sheetfhs, sheetlhs, noOFDay, DayStart, DayEnd);//Calls the method to load the fix clockedshiftstable
        //        importdone = 1;

        //    }

        //    Application.DoEvents();
        //}

        public void extractMiners()
        {
            //intCounter = 0;
            //deleteAllColumns("Miners");
            //Application.DoEvents();
            //string strSQL = Base.extractMiners(Base.DBConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType, txtSelectedSection.Text.Trim());

            ////Check and add the calc columns
            //strSQL = checkSQL(intCounter, strSQL);

            //strSQL = strSQL.Replace(")", "") + ";commit transaction;";

            DataTable tempMiners = Base.extractMiners(Base.DBConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType, txtSelectedSection.Text.Trim());

            //DataTable test = Base.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

            TB.saveCalculations2(tempMiners, Base.DBConnectionString, strWhere, "MINERS");
            evaluateMiners();

        }

        private void extractGangLink()
        {
            //Add the rigging, equipping and tramming gangs to the ganglinking.
            DataTable TmpGanglink = Base.extractGanglink(Base.DBConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType, txtSelectedSection.Text.Trim());

            if (TmpGanglink.Rows.Count > 0)
            {

                TB.saveCalculations2(TmpGanglink, Base.DBConnectionString, strWhere, "GANGLINK");
                Application.DoEvents();

            }
            else
            {
                MessageBox.Show("No records for ganglinking were extracted for section: " + txtSelectedSection.Text.Trim(), "Information", MessageBoxButtons.OK);
            }
        }

        private void btnLock_Click(object sender, EventArgs e)
        {
            //lstBErrorLog.Items.Clear();
            //string faultPath = "c:\\Reports\\BonusshiftsvsClockedShiftsReport" + DateTime.Now.ToString("yymmddhhmmss") + ".txt";//Sets the path file
            //StreamWriter sw = new StreamWriter(faultPath);
            int goOn = 1;
            List<string> lstDrillers = new List<string>();
            string strProcess = tabInfo.SelectedTab.Name;

          

            if (goOn == 1)
            {

                if (btnLock.Text == "Lock")
                {
                    switch (strProcess)
                    {
                        #region identify drillers with more drilling shifts that adteam shifts

                        case "tabDrillers":

                            //string strSQL = " delete from drillers where workplace not in (select distinct workplace from SURVEY where PERIOD = '" + BusinessLanguage.Period +
                            //             "' and SECTION = '" + txtSelectedSection.Text.Trim() + "') " +
                            //             " and PERIOD = '" + BusinessLanguage.Period +
                            //             "' and SECTION = '" + txtSelectedSection.Text.Trim() + "'";

                            //TB.InsertData(Base.DBConnectionString, strSQL);

                            //a driller cannot have more driller shifts that month shifts
                            string strSQL = "select t1.* from (select employee_no,sum(convert(float,drillershifts)) as totaldrillershifts " +
                                                 "from drillers where drillerind = '1' and section = '" + txtSelectedSection.Text.Trim() +
                                                 "' and period = '" + BusinessLanguage.Period + "' group by employee_no) as t1" +
                                                 " where totaldrillershifts > '" + strMonthShifts + "'";

                            DataTable invalidDrillers = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                            if (invalidDrillers.Rows.Count > 0)
                            {
                                btnShowAll.Visible = true;
                                btnShowEmpl.Visible = true;
                                grdDrillers.DataSource = invalidDrillers;
                                MessageBox.Show("The following employees have more shifts that the allowed measuring shifts." + Environment.NewLine +
                                                "The Drillers input screen will not finalize until shifts have been fixed.", "Warning", MessageBoxButtons.OK);
                            }
                            else
                            {
                                btnShowAll.Visible = false;
                                btnShowEmpl.Visible = false;
                                
                                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = '" + strProcess +
                                         "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                                btnLock.Text = "Unlock";
                                evaluateInputProcessStatus();
                                openTab(tabProcess);

                        
                                Application.DoEvents();


                            }

                            break;
                        #endregion


                        case "tabGangLinking":

                            //strSQL = " delete from ganglink where workplace not in (select distinct workplace from SURVEY where PERIOD = '" + BusinessLanguage.Period +
                            //         "' and SECTION = '" + txtSelectedSection.Text.Trim() + "') " +
                            //         " and PERIOD = '" + BusinessLanguage.Period +
                            //         "' and SECTION = '" + txtSelectedSection.Text.Trim() + "'";

                            //TB.InsertData(Base.DBConnectionString, strSQL);
                            TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = '" + strProcess +
                                         "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                            btnLock.Text = "Unlock";
                            evaluateInputProcessStatus();
                            openTab(tabProcess);


                            Application.DoEvents();
                            break;

                        default:
                            TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = '" + strProcess +
                                         "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                            btnLock.Text = "Unlock";
                            evaluateInputProcessStatus();
                            openTab(tabProcess);


                            Application.DoEvents();
                            break;

                    }

                   

                }

                else
                {

                    TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where process = '" + strProcess +
                                          "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                    btnLock.Text = "Lock";
                    evaluateInputProcessStatus();
                    openTab(tabProcess);

                    Application.DoEvents();

                }

                
            }
            else
            {

               
            }

        }

        private void grdMiners_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            MessageBox.Show("column = " + e.ColumnIndex + "  row = " + e.RowIndex);
            string test = scrTeamS.ActiveForm.ActiveControl.Name;
            MessageBox.Show("name = " + test);
        }

        private void btnInsertRow_Click(object sender, EventArgs e)
        {
            string strSQL = string.Empty;
            string strName = string.Empty;
            string strDesignation = string.Empty;
            string strDesignationDesc = string.Empty;

            switch (tabInfo.SelectedTab.Name)
            {
                case "tabMiners":
                    #region tabMiners
                    //HJ
                    if (cboMinersGangNo.Text.Trim().Length != 0 && cboNames.Text.Trim().Length != 0 &&
                        cboDesignation.Text.Trim().Length != 0 && txtPayShifts.Text.Trim().Length != 0 &&
                        txtAwops.Text.Trim().Length != 0 &&
                        txtMinersSafetyInd.Text.Trim().Length != 0)
                    {

                        //intCounter = 0;
                        //deleteAllColumns("Miners");
                        //Application.DoEvents();

                        if (cboNames.Text.Contains("-"))
                        {
                            strName = cboNames.Text.Substring(0, cboNames.Text.IndexOf("-")).Trim();
                        }
                        else
                        {
                            strName = cboNames.Text.Trim();
                        }

                        if (cboDesignation.Text.Contains("-"))
                        {
                            strDesignation = cboDesignation.Text.Substring(0, cboDesignation.Text.IndexOf("-")).Trim();
                            strDesignationDesc = cboDesignation.Text.Substring((cboDesignation.Text.IndexOf("-")) + 3);
                        }
                        else
                        {
                            strDesignation = cboDesignation.Text.Trim();
                        }

                        //strSQL = "Insert into Miners values ('" + BusinessLanguage.BussUnit +
                        //         "', '" + BusinessLanguage.MiningType + "', '" + BusinessLanguage.BonusType +
                        //         "', '" + txtSelectedSection.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                        //         "', '" + cboMinersGangNo.Text.Trim() + "', '" + cboNames.Text.Trim() +
                        //         "', '" + strDesignation + "', '" + txtSurname.Text.Trim() +
                        //         "', '" + strDesignationDesc  + "', '" + txtShifts.Text.Trim() +
                        //         "', '" + txtAwops.Text.Trim() + "', '" + txtPayShifts.Text.Trim() + 
                        //         "', '" + txtMinersSafetyInd.Text.Trim() + "'" ;

                        //strSQL = checkSQL(intCounter, strSQL);

                        //TB.InsertData(Base.DBConnectionString, strSQL);

                        DataTable temp = new DataTable();
                        temp = Miners.Copy();

                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();

                        }

                        temp.AcceptChanges();
                        DataRow dr = temp.NewRow();

                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                        dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                        dr["PERIOD"] = txtPeriod.Text.Trim();
                        dr["GANG"] = cboMinersGangNo.Text.Trim();
                        dr["EMPLOYEE_NO"] = cboNames.Text.Trim();
                        dr["DESIGNATION"] = strDesignation;
                        dr["EMPLOYEE_NAME"] = cboMinersEmpName.Text.Trim();
                        dr["DESIGNATION_DESC"] = strDesignationDesc;
                        dr["SHIFTS_WORKED"] = txtPayShifts.Text.Trim();
                        dr["AWOP_SHIFTS"] = txtAwops.Text.Trim();
                        dr["PAYSHIFTS"] = txtPayShifts.Text.Trim();
                        dr["SAFETYIND"] = txtMinersSafetyInd.Text.Trim();

                        temp.Rows.Add(dr);
                        //Create a total invalid delete.
                        string strDelete = " where Bussunit = '999'";
                        int rowindex = grdMiners.CurrentCell.RowIndex;
                        TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "MINERS");
                        evaluateMiners();

                        grdMiners.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabGangLinking":
                    #region tabGangLinking

                    if (cboGangLinkSafetyInd.Text.Trim().Length > 0 &&
                        cboGangLinkGangType.Text.Trim().Length > 0 &&
                        cboGangLinkWorkplace.Text.Trim().Length > 0 &&
                        cboGangLinkGang.Text.Trim().Length > 0)
                    {
                        //Get the layout of the ganglink file.
                        DataTable temp = new DataTable();
                        temp = GangLink.Copy();

                        int intRow = 0;
                        if (grdGangLink.CurrentRow == null)
                        {
                            intRow = 0;
                        }
                        else
                        {
                            intRow = grdGangLink.CurrentCell.RowIndex;
                        }

                        //Clear the input temp table
                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();
                        }

                        temp.AcceptChanges();

                        if (lstGangs.SelectedItems.Count == 0)
                        {
                            MessageBox.Show("Please select gangs from the listbox", "Information", MessageBoxButtons.OK);
                        }
                        else
                        {

                            for (int i = 0; i < lstGangs.SelectedItems.Count; i++)
                            {
                                DataRow dr = temp.NewRow();

                                dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                                dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                                dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                                dr["SECTION"] = txtSelectedSection.Text.Trim();
                                dr["PERIOD"] = txtPeriod.Text.Trim();
                                dr["WORKPLACE"] = cboGangLinkWorkplace.Text.Trim();
                                dr["GANG"] = lstGangs.SelectedItems[i].ToString();
                                dr["SAFETYIND"] = cboGangLinkSafetyInd.Text.Trim();
                                dr["GANGTYPE"] = cboGangLinkGangType.Text.Trim();

                                temp.Rows.Add(dr);

                            }
                            //Create a invalid delete that will execute in the savecalculation2 method.
                            string strDelete = " where Bussunit = '999'";

                            TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "GANGLINK");
                            evaluateGangLinking();

                            grdGangLink.FirstDisplayedScrollingRowIndex = intRow;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Supply all input data. Please check that all input boxes contain data.", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion

                case "tabAbnormal":
                    #region tabAbnormal

                    if (txtAbnormalContract.Text.Trim().Length != 0 && cboAbnormalWorkplace.Text.Trim().Length != 0 &&
                        cboAbnormalLevel.Text.Trim().Length != 0 && cboAbnormalType.Text.Trim().Length != 0 &&
                        txtAbnormalValue.Text.Trim().Length != 0)
                    {
                        DataRow dr;
                        dr = Abnormal.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                        dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                        dr["PERIOD"] = txtPeriod.Text.Trim();

                        dr["WORKPLACE"] = cboAbnormalWorkplace.Text.Trim();
                        dr["ABNORMALLEVEL"] = cboAbnormalLevel.Text.Trim();
                        dr["ABNORMALTYPE"] = cboAbnormalType.Text.Trim();
                        dr["ABNORMALVALUE"] = txtAbnormalValue.Text.Trim();

                        Abnormal.Rows.Add(dr);
                        int intRow = 0;
                        if (grdAbnormal.CurrentRow == null)
                        {
                            intRow = 0;
                        }
                        else
                        {
                            int rowindex = grdAbnormal.CurrentCell.RowIndex;
                        }

                        strSQL = "Insert into Abnormal values ('" + BusinessLanguage.BussUnit.Trim() +
                                 "', '" + BusinessLanguage.MiningType.Trim() + "', '" + BusinessLanguage.BonusType.Trim() +
                                 "', '" + txtSelectedSection.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                                 "', '" + txtAbnormalContract.Text.Trim() + "', '" + cboAbnormalWorkplace.Text.Trim() +
                                 "', '" + cboAbnormalLevel.Text.Trim() + "', '" + cboAbnormalType.Text.Trim() +
                                 "', '" + txtAbnormalValue.Text.Trim() + "')";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        evaluateAbnormal();
                        grdAbnormal.FirstDisplayedScrollingRowIndex = intRow;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties
                    if (cboEmplPenEmployeeNo.Text.Trim().Length > 0 &&
                        txtPenaltyValue.Text.Trim().Length > 0 && cboPenaltyInd.Text.Trim().Length > 0)
                    {
                        DataRow dr;
                        dr = EmplPen.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                        dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                        dr["PERIOD"] = txtPeriod.Text.Trim();
                        dr["EMPLOYEE_NO"] = cboEmplPenEmployeeNo.Text.Trim();
                        dr["PENALTYVALUE"] = txtPenaltyValue.Text.Trim();
                        dr["PENALTYIND"] = cboPenaltyInd.Text.Trim();

                        EmplPen.Rows.Add(dr);

                        strSQL = "Insert into EmployeePenalties values ('" + BusinessLanguage.BussUnit.Trim() +
                                 "', '" + BusinessLanguage.MiningType.Trim() + "', '" + BusinessLanguage.BonusType.Trim() +
                                 "', '" + txtSelectedSection.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                                 "', '" + cboEmplPenEmployeeNo.Text.Trim() + "', '" + txtPenaltyValue.Text.Trim() +
                                 "', '" + cboPenaltyInd.Text.Trim() + "')";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabOffday":
                    #region tabOffdays
                    this.Cursor = Cursors.WaitCursor;
                    if (cboOffDaysSection.Text.Trim().Length > 0)
                    {

                        //Get the layout of the offday file.
                        DataTable temp = new DataTable();
                        temp = Offdays.Copy();

                        int intRow = 0;
                        if (grdOffDays.CurrentRow == null)
                        {
                            intRow = 0;
                        }
                        else
                        {
                            intRow = grdOffDays.CurrentCell.RowIndex;
                        }

                        //Clear the input temp table
                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();
                        }

                        temp.AcceptChanges();

                        if (lstOffDayValue.SelectedItems.Count == 0)
                        {
                            MessageBox.Show("Please select dates from the listbox", "Information", MessageBoxButtons.OK);
                        }
                        else
                        {

                            for (int i = 0; i < lstOffDayValue.SelectedItems.Count; i++)
                            {
                                DataRow dr = temp.NewRow();
                                dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();             //xxxxxxxxxxxxxxxx
                                dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();//xxxxxxxxxxxxxxxx
                                dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();//xxxxxxxxxxxxxxxx
                                dr["SECTION"] = cboOffDaysSection.Text.Trim();
                                dr["PERIOD"] = BusinessLanguage.Period;//xxxxxxxxxxxxxxxx
                                dr["GANG"] = cboOffDaysGang.Text.Trim();
                                dr["OFFDAYVALUE"] = lstOffDayValue.SelectedItems[i].ToString();

                                temp.Rows.Add(dr);

                            }
                            //Create a invalid delete that will execute in the savecalculation2 method.
                            string strDelete = " where section = '999'";

                            TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "OFFDAYS");
                            evaluateOffDays();

                            grdOffDays.FirstDisplayedScrollingRowIndex = intRow;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Supply all input data. Please check that all input boxes contain data.", "Error", MessageBoxButtons.OK);
                    }

                    this.Cursor = Cursors.Arrow;
                    break;

                    #endregion

                case "tabRates":
                    #region tabRates
                    if (txtLowValue.Text.Trim().Length != 0 &&
                        txtHighValue.Text.Trim().Length != 0 && txtRate.Text.Trim().Length != 0)
                    {
                        DataRow dr;
                        dr = Rates.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
                        dr["MININGTYPE"] = BusinessLanguage.MiningType;
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType;
                        dr["PERIOD"] = txtPeriod.Text.Trim();
                        dr["RATE_TYPE"] = txtRateType.Text.Trim();
                        dr["LOW_VALUE"] = txtLowValue.Text.Trim();
                        dr["HIGH_VALUE"] = txtHighValue.Text.Trim();
                        dr["RATE"] = txtRate.Text.Trim();

                        int rowindex = grdMiners.CurrentCell.RowIndex;
                        strSQL = "Insert into Rates values ('" + BusinessLanguage.BussUnit.Trim() +
                                 "', '" + BusinessLanguage.MiningType.Trim() + "', '" + BusinessLanguage.BonusType +
                                 "', '" + txtRateType.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                                 "', '" + txtLowValue.Text.Trim() + "', '" + txtHighValue.Text.Trim() +
                                 "', '" + txtRate.Text.Trim() + "')";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdGangLink.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabDrillers":
                    #region tabDrillers


                    if (txtAutoDGang.Text.Trim().Length > 0 &&
                        txtAutoDWorkplace.Text.Trim().Length > 0 &&
                        txtAutoDrilShifts.Text.Trim().Length > 0 &&
                        cboAutoDrillerDrilInd.Text.Trim().Length > 0)
                    {
                        if (Convert.ToInt32(cboAutoDrillerDrilInd.Text.ToString().Trim()) > 0)
                        {
                            DataTable temp = new DataTable();
                            temp = Drillers.Copy();

                            int intRow = 0;
                            if (grdDrillers.CurrentRow == null)
                            {
                                intRow = 0;
                            }
                            else
                            {
                                intRow = grdDrillers.CurrentCell.RowIndex;
                            }

                            //Clear the input temp table
                            for (int i = 0; i <= temp.Rows.Count - 1; i++)
                            {
                                temp.Rows[i].Delete();
                            }

                            temp.AcceptChanges();

                            if (lstDrillers.SelectedItems.Count == 0)
                            {
                                MessageBox.Show("Please select employees from the listbox", "Information", MessageBoxButtons.OK);
                            }
                            else
                            {
                                string strWorkplace = string.Empty;

                                if (txtAutoDWorkplace.Text.Contains("-"))
                                {
                                    strWorkplace = txtAutoDWorkplace.Text.Trim().Substring(0, txtAutoDWorkplace.Text.IndexOf("-")).Trim();

                                }
                                else
                                {
                                    strWorkplace = txtAutoDWorkplace.Text.Trim();
                                }

                                for (int i = 0; i < lstDrillers.SelectedItems.Count; i++)
                                {
                                    DataRow dr = temp.NewRow();

                                    dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                                    dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                                    dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                                    dr["SECTION"] = txtSelectedSection.Text.Trim();
                                    dr["PERIOD"] = txtPeriod.Text.Trim();
                                    dr["WORKPLACE"] = strWorkplace.Trim();
                                    dr["GANG"] = txtAutoDGang.Text.Trim();
                                    dr["EMPLOYEE_NO"] = lstDrillers.SelectedItems[i].ToString().Substring(0, lstDrillers.SelectedItems[i].ToString().IndexOf("-"));
                                    dr["DRILLERIND"] = cboAutoDrillerDrilInd.SelectedItem.ToString().Trim();
                                    dr["DRILLERSHIFTS"] = txtAutoDrilShifts.Text.Trim();
                                    temp.Rows.Add(dr);

                                }
                                //Create a invalid delete that will execute in the savecalculation2 method.
                                string strDelete = " where Bussunit = '999'";

                                TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "DRILLERS");
                                evaluateDrillers();

                                grdDrillers.FirstDisplayedScrollingRowIndex = intRow;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Employee must have a driller indicator of 1 or 2", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion


            }
        }

        private string checkSQL(int intCounter, string strSQL)
        {
            if (intCounter > 0)
            {
                for (int i = 0; i <= intCounter - 1; i++)
                {
                    strSQL = strSQL.Trim() + ",'0'";
                }
                strSQL = strSQL.Trim() + ")";
            }
            else
            {
                strSQL = strSQL.Trim() + "')";
            }

            return strSQL;
        }

        private void UpdateClockedShifts()
        {
            #region Extract dates
            //Load the section's first and last shift date
            DateTime dteFSH = dateTimePicker1.Value;
            DateTime dteLSH = dateTimePicker2.Value;

            string tempdte = Clocked.Rows[1]["FSH"].ToString().Trim();
            DateTime dteDateFrom = Convert.ToDateTime(tempdte.Trim());

            tempdte = Clocked.Rows[1]["LSH"].ToString().Trim();
            DateTime dteDateEnd = Convert.ToDateTime(tempdte.Trim());

            int intstart = dteDateFrom.Subtract(dteFSH).Days + 1;
            int intend = dteLSH.Subtract(dteDateFrom).Days + 2;

            #endregion

            foreach (DataRow dr in Offdays.Rows)
            {
                string offdate = dr["OFFDAYVALUE"].ToString();
                if (offdate.Trim() == "2009-01-01")
                {
                }
                else
                {

                    DateTime dteOffdate = Convert.ToDateTime(dr["OFFDAYVALUE"].ToString());

                    int intOffday = dteOffdate.Subtract(dteDateFrom).Days;

                    Base.UpdateOffdays(Base.DBConnectionString, intOffday);

                    Application.DoEvents();
                }
            }

        }

        private void grdMiners_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            DataTable temp = new DataTable();

            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboMinersGangNo.Text = grdMiners["GANG", e.RowIndex].Value.ToString().Trim();
                cboMinersEmpName.Text = grdMiners["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
                cboNames.Text = grdMiners["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                cboDesignation.Text = grdMiners["DESIGNATION", e.RowIndex].Value.ToString().Trim() + "  -  " + grdMiners["DESIGNATION_DESC", e.RowIndex].Value.ToString().Trim();
                txtPayShifts.Text = grdMiners["PAYSHIFTS", e.RowIndex].Value.ToString().Trim();
                txtAwops.Text = grdMiners["AWOP_SHIFTS", e.RowIndex].Value.ToString().Trim();
                cboMinersEmpName.Text = grdMiners["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
                txtMinersSafetyInd.Text = grdMiners["SAFETYIND", e.RowIndex].Value.ToString().Trim();
                //txtTotalCall.Text = grdMiners["TOTALCALL", e.RowIndex].Value.ToString().Trim();

                if (grdMiners["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim() == "XXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                    btnInsertRow.Enabled = true;
                }
                else
                {
                    btnInsertRow.Enabled = true;
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                }

                if (Clocked.Rows.Count > 0)
                {
                    IEnumerable<DataRow> query2 = from locks in Clocked.AsEnumerable()
                                                  where locks.Field<string>("EMPLOYEE_NO").TrimEnd() == cboNames.Text.Trim()
                                                  select locks;

                    try
                    {
                        temp = query2.CopyToDataTable<DataRow>();
                    }
                    catch { }
                }

                if (temp.Rows.Count > 0)
                {
                    //txtADTeamShifts.Text = temp.Rows[0]["Payshifts"].ToString().Trim();
                }
                else
                {
                    txtADTeamShifts.Text = "0";
                }

            }
        }

        private void tabInfo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtSelectedSection.Text == "***")
            {
                MessageBox.Show("Please select a section.", "Information", MessageBoxButtons.OK);
            }
            else
            {
                btnInsertRow.Enabled = true;
                btnUpdate.Enabled = true;

                btnDeleteRow.Enabled = false;
                listBox1.Enabled = false;                               //HJ
                btnLoad.Enabled = false;
                dateTimePicker1.Enabled = false;                        //HJ
                dateTimePicker2.Enabled = false;                        //HJ
                btnPrint.Enabled = false;
                btnLock.Enabled = false;
                panelLock.BackColor = Color.Lavender;

                int intCount = checkLock(tabInfo.SelectedTab.Name);
                if (intCount > 0)
                {
                    btnLock.Text = "Unlock";
                }
                else
                {
                    btnLock.Text = "Lock";
                }

                switch (tabInfo.SelectedTab.Name)
                {
                    #region tabCalendar
                    case "tabCalendar":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnLoad.Enabled = true;
                        dateTimePicker1.Enabled = true;                 //HJ
                        dateTimePicker2.Enabled = true;                 //HJ
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;


                        lstPrimaryKeyColumns.Clear();
                        extractPrimaryKey(Calendar, "CALENDAR");


                        break;
                    #endregion

                    #region tabSurvey
                    case "tabSurvey":
                        
                        btnInsertRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateSurvey();
                        break;
                    #endregion

                    #region tabClockShifts
                    case "tabClockShifts":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnPrint.Enabled = true;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        break;
                    #endregion

                    #region tabLabour
                    case "tabLabour":

                        btnInsertRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateLabour();

                        extractPrimaryKey(Labour, "BONUSSHIFTS");
                        break;
                    #endregion

                    #region tabAbnormal
                    case "tabAbnormal":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateAbnormal();

                        extractPrimaryKey(Abnormal, "ABNORMAL");

                        break;
                    #endregion

                    #region tabMiners
                    case "tabMiners":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateMiners();


                        extractPrimaryKey(Miners, "MINERS");
                        break;
                    #endregion

                    #region tabGangLinking
                    case "tabGangLinking":
                        cboGangLinkGang.Items.Clear();
                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateGangLinking();

                        extractPrimaryKey(GangLink, "GANGLINK");
                        break;

                    #endregion

                    #region tabConfig
                    case "tabConfig":

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        extractPrimaryKey(Configs, "CONFIGURATION");
                        break;

                    #endregion

                    #region tabEmplPen
                    case "tabEmplPen":

                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        extractPrimaryKey(EmplPen, "EMPLOYEEPENALTY");
                        break;

                    #endregion 

                    #region tabOffday
                    case "tabOffday":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;

                        break;

                    #endregion

                    #region tabSelected
                    case "tabSelected":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        listBox1.Enabled = true;                            //HJ
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        break;

                    #endregion

                    #region tabStatus

                    case "tabProcess":

                        evaluateStatus();
                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnDeleteRow.Enabled = false;
                        btnLoad.Enabled = false;
                        btnPrint.Enabled = false;
                        btnLock.Enabled = false;

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        break;

                    #endregion

                    #region tabRates
                    case "tabRates":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;
                        extractPrimaryKey(Rates, "RATES");
                        break;

                    #endregion

                    #region tabMonitor
                    case "tabMonitor":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;
                        break;

                    #endregion

                    #region tabDrillers
                    case "tabDrillers":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;
                        btnShowAll.Visible = false;
                        btnShowEmpl.Visible = false;
                        extractPrimaryKey(Drillers, "DRILLERS");
                         
                        break;

                    #endregion

                    #region tabFactors

                    case "tabFactors":

                        evaluateFactors();
                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = true;
                        btnDeleteRow.Enabled = false;
                        btnLoad.Enabled = false;
                        btnPrint.Enabled = false;
                        btnLock.Enabled = false;

                        panelInsert.BackColor = Color.LightGray;
                        panelUpdate.BackColor = Color.LightGray;
                        panelDelete.BackColor = Color.LightGray;
                        panelPreCalcReport.BackColor = Color.LightGray;
                        extractPrimaryKey(Factors, "FACTORS");
                        break;

                    #endregion

                }
            }

        }

        private void extractPrimaryKey(DataTable p, string tablename)
        {
            //List Names contains the primary key columns of the selected table
            lstPrimaryKeyColumns.Clear();
            switch (tablename)
            {
                case "CALENDAR":
                    lstPrimaryKeyColumns = Base.listCalendarPrimaryKey;
                    break;

                case "ABNORMAL":
                    lstPrimaryKeyColumns = Base.listAbnormalPrimaryKey;
                    break;

                case "BONUSSHIFTS":
                    lstPrimaryKeyColumns = Base.listBonusShiftsPrimaryKey;
                    break;

                case "GANGLINK":
                    lstPrimaryKeyColumns = Base.listGangLinkPrimaryKey;
                    break;

                case "SUPPORTLINK":
                    lstPrimaryKeyColumns = Base.listSupportLinkPrimaryKey;
                    break;

                case "DRILLERS":
                    lstPrimaryKeyColumns = Base.listDrillersPrimaryKey;
                    break;

                case "RATES":
                    lstPrimaryKeyColumns = Base.listRatesPrimaryKey;
                    break;

                case "FACTORS":
                    lstPrimaryKeyColumns = Base.listFactorsPrimaryKey;
                    break;

                case "MINERS":
                    lstPrimaryKeyColumns = Base.listMinersPrimaryKey;
                    break;

                case "CONFIGURATION":
                    lstPrimaryKeyColumns = Base.listConfigurationPrimaryKey;
                    break;


            }


            ////lstTableColumns contains all the column names of the table excluding "BUSSUNIT","MININGTYPE","BONUSTYPE","PERIOD")
            ////Do this extract on the table in memory, because much quicker.
            lstTableColumns.Clear();
            DataTable temp = p.Copy();
            deleteAllCalcColumnsFromTempTable(tablename, temp);

            if (temp.Columns.Count > 0)
            {
                foreach (DataColumn col in temp.Columns)
                {
                    if (col.ColumnName == "BUSSUNIT" || col.ColumnName == "MININGTYPE" || col.ColumnName == "BONUSTYPE" || col.ColumnName == "PERIOD")
                    {
                    }
                    else
                    {
                        lstTableColumns.Add(col.ColumnName.ToString().Trim());
                    }
                }
            }
        }

        private int checkLock(string processToBeChecked)
        {
            //Lynx....LINQ
            DataTable contactTable = TB.getDataTable(TB.TBName);

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "Y"
                                          where locks.Field<string>("PROCESS").TrimEnd() == processToBeChecked
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          select locks;


            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();
            int intcount = query1.Count<DataRow>();

            return intcount;

            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();

        }

        private int checkLockInputProcesses()
        {

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "N"
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          select locks;

            int intcount = query1.Count<DataRow>();

            return intcount;

            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();

        }

        private void grdSurvey_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {
            }
            else
            {

                txtSurveyWorkplace.Text = grdSurvey["WORKPLACE", e.RowIndex].Value.ToString().Trim();
                txtSurveySwpDist.Text = grdSurvey["SWEEPINGDISTANCE", e.RowIndex].Value.ToString().Trim();

            }

            Cursor.Current = Cursors.Arrow;
        }

        private void grdEmplPen_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboEmplPenEmployeeNo.Text = grdEmplPen["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                txtPenaltyValue.Text = grdEmplPen["PENALTYVALUE", e.RowIndex].Value.ToString().Trim();
                cboPenaltyInd.Text = grdEmplPen["PENALTYIND", e.RowIndex].Value.ToString().Trim();
                if (grdEmplPen["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim() == "XXXXXXXXXXXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                }
                else
                {
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                }
            }
            Cursor.Current = Cursors.Arrow;

        }

        private void grdLabour_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
                cboBonusShiftsGang.Text = grdLabour["GANG", e.RowIndex].Value.ToString().Trim();
                txtEmployeeNo.Text = grdLabour["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                txtEmployeeName.Text = grdLabour["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
                cboBonusShiftsWageCode.Text = grdLabour["WAGECODE", e.RowIndex].Value.ToString().Trim();
                cboBonusShiftsResponseCode.Text = grdLabour["LINERESPCODE", e.RowIndex].Value.ToString().Trim();
                txtShifts.Text = grdLabour["SHIFTS_WORKED", e.RowIndex].Value.ToString().Trim();
                txtAwop.Text = grdLabour["AWOP_SHIFTS", e.RowIndex].Value.ToString().Trim();
                txtStrikeShifts.Text = grdLabour["STRIKE_SHIFTS", e.RowIndex].Value.ToString().Trim();
                cboDrillerInd.Text = grdLabour["DRILLERIND", e.RowIndex].Value.ToString().Trim();
                txtDrillerShifts.Text = grdLabour["DrillerShifts", e.RowIndex].Value.ToString().Trim();
                //cboTLeader.Text = grdLabour["TEAMLEADERIND", e.RowIndex].Value.ToString().Trim();


            }

            Cursor.Current = Cursors.Arrow;

        }

        private void grdConfigs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
                cboParameterName.Text = grdConfigs["PARAMETERNAME", e.RowIndex].Value.ToString().Trim();
                cboParm1.Text = grdConfigs["PARM1", e.RowIndex].Value.ToString().Trim();
                cboParm2.Text = grdConfigs["PARM2", e.RowIndex].Value.ToString().Trim();
                cboParm3.Text = grdConfigs["PARM3", e.RowIndex].Value.ToString().Trim();
                cboParm4.Text = grdConfigs["PARM4", e.RowIndex].Value.ToString().Trim();
                cboParm5.Text = grdConfigs["PARM5", e.RowIndex].Value.ToString().Trim();
                cboParm6.Text = grdConfigs["PARM6", e.RowIndex].Value.ToString().Trim();
                cboParm7.Text = grdConfigs["PARM7", e.RowIndex].Value.ToString().Trim();
            }
            Cursor.Current = Cursors.Arrow;
        }

        private void grdOffdays_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        #region AutoSize
        private void autoSizeGrid(DataGridView DG)
        {
            if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader.ToString())
            {
                DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            else
            {
                if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCells.ToString())
                {
                    DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                }
                else
                {
                    if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.ColumnHeader.ToString())
                    {
                        DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                    }
                    else
                    {
                        if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.DisplayedCells.ToString())
                        {
                            DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
                        }
                        else
                        {
                            if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCells.ToString())
                            {
                                DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
                            }
                            else
                            {
                                if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.Fill.ToString())
                                {
                                    DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                                }
                                else
                                {
                                    if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.DisplayedCells.ToString())
                                    {
                                        DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void grdActiveSheet_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdActiveSheet);
            }
        }


        private void grdCalendar_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdCalendar);
            }
        }

        private void grdSurvey_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdSurvey);
            }
        }

        private void grdClocked_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdClocked);
            }
        }

        private void grdLabour_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdLabour);
            }
        }


        private void grdDrillers_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdDrillers);
            }
        }

        private void grdGangLink_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdGangLink);
            }
        }

        private void grdMiners_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdMiners);
            }
        }

        private void grdOccupations_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right)
            //{
            //    autoSizeGrid(grdOccupations);
            //}
        }

        //private void grdEmplList_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        //{
        //    if (e.Button == MouseButtons.Right)
        //    {
        //        autoSizeGrid(grdAbnormal);
        //    }
        //}

        private void grdConfigs_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdConfigs);
            }
        }

        private void DoDataExtract()
        {
            connectToDB();
            TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName);

        }
        #endregion

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string FormulaTableName = string.Empty;

            TB.TBName = (string)listBox1.SelectedItem;

            if (TB.TBName.Trim().ToUpper().Contains("EARN") && TB.TBName.Trim().ToUpper().Contains("20"))
            {
                FormulaTableName = TB.TBName.Trim().Substring(0, TB.TBName.Trim().ToUpper().IndexOf("20"));   //xxxxxxxxxxxxxxxxxx
            }
            else
            {
                FormulaTableName = TB.TBName;
            }

            TB.DBName = Base.DBName;

            connectToDB();
            cboColumnValues.Items.Clear();
            cboColumnNames.Items.Clear();
            cboColumnNames.Text = string.Empty;
            cboColumnValues.Text = string.Empty;

            List<string> lstColumnNames = General.getListOfColumnNames(Base.DBConnectionString, TB.TBName);

            foreach (string s in lstColumnNames)
            {
                cboColumnNames.Items.Add(s.Trim());
                cboColumnShow.Items.Add(s.Trim());
            }

            TB.ListOfSelectedTableColumns = lstColumnNames;

            DoDataExtract(strWhere);
            newDataTable = TB.getDataTable(TB.TBName);
            if (newDataTable == null)
            {
                DoDataExtract(strWherePeriod);
                newDataTable = TB.getDataTable(TB.TBName);

            }
            else
            {

            }

            grdActiveSheet.DataSource = TB.getDataTable(TB.TBName);

            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName + BusinessLanguage.Period.Trim(), 
                 FormulaTableName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }

            hideColumnsOfGrid("grdActiveSheet");
        }

        private void DoDataExtract(string Where)
        {
            connectToDB();
            if (Where.Trim().Length == 0)
            {
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName);
            }
            else
            {
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName, Where);

            }
        }

        private void exportToExcel(string path, DataTable dt)
        {
            if (dt.Columns.Count > 0)
            {
                string OPath = path + "\\" + TB.TBName + ".xls";
                try
                {
                    StreamWriter SW = new StreamWriter(OPath);
                    System.Web.UI.HtmlTextWriter HTMLWriter = new System.Web.UI.HtmlTextWriter(SW);
                    System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();

                    grid.DataSource = dt;
                    grid.DataBind();

                    using (SW)
                    {
                        using (HTMLWriter)
                        {
                            grid.RenderControl(HTMLWriter);
                        }
                    }

                    SW.Close();
                    HTMLWriter.Close();
                    MessageBox.Show("Your spreadsheet was created at: " + OPath, "Information", MessageBoxButtons.OK);
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Could not create " + OPath.Trim() + ".  Create the directory first." + exx.Message, "Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Your spreadsheet could not be created.  No columns found in datatable.", "Error Message", MessageBoxButtons.OK);
            }

        }

        private void TBExport_Click(object sender, EventArgs e)
        {
            saveTheSpreadSheet();
        }

        private void saveTheSpreadSheet()
        {
            string path = @"c:\" + TB.DBName + "\\" + TB.TBName;
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                DoDataExtract();
                DataTable outputTable = TB.getDataTable(TB.TBName);
                exportToExcel(path, outputTable);
                MessageBox.Show("Successfully Downloaded.", "Information", MessageBoxButtons.OK);

            }
            catch (Exception ee)
            {
                Console.WriteLine("The process failed: {0}", ee.ToString());
            }

            finally { }
        }

        private void grdActiveSheet_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //Get calc name
            this.Cursor = Cursors.WaitCursor;
            int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
            int rownr = grdActiveSheet.CurrentCell.RowIndex;
            TBFormulas.CalcName = grdActiveSheet.Columns[columnnr].HeaderText;

            //Check if it is a calculated column
            object intCount = Analysis.countcalcbyname(TB.DBName, TB.TBName, TBFormulas.CalcName.Trim(), Base.AnalysisConnectionString);
            if ((int)intCount > 0)
            {
                //It is a calculated column.
                DataTable dtFormula = Analysis.GetCalcDetails(TB.DBName, TB.TBName, TBFormulas.CalcName, Base.AnalysisConnectionString);
                //Extract the formula details:
                decimal decValue = 0;
                try
                {
                    decValue = Convert.ToDecimal(grdActiveSheet.CurrentCell.Value);
                }
                catch
                {
                    decValue = 0;
                }

                //Extract Factors
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, "FACTORS");
                DataTable dtFactors = TB.getDataTable("FACTORS");
                dict.Clear();
                loadDict(dtFactors);

                if (dtFormula.Rows.Count > 0)
                {
                    TBFormulas.A = dtFormula.Rows[0]["A"].ToString().Trim();
                    TBFormulas.B = dtFormula.Rows[0]["B"].ToString().Trim();
                    TBFormulas.C = dtFormula.Rows[0]["C"].ToString().Trim();
                    TBFormulas.D = dtFormula.Rows[0]["D"].ToString().Trim();
                    TBFormulas.E = dtFormula.Rows[0]["E"].ToString().Trim();
                    TBFormulas.F = dtFormula.Rows[0]["F"].ToString().Trim();
                    TBFormulas.G = dtFormula.Rows[0]["G"].ToString().Trim();
                    TBFormulas.H = dtFormula.Rows[0]["H"].ToString().Trim();
                    TBFormulas.I = dtFormula.Rows[0]["I"].ToString().Trim();
                    TBFormulas.J = dtFormula.Rows[0]["J"].ToString().Trim();
                    TBFormulas.TableFormulaCall = dtFormula.Rows[0]["FORMULA_CALL"].ToString().Trim();
                    decimal decA = 0;
                    decimal decB = 0;
                    decimal decC = 0;
                    decimal decD = 0;
                    decimal decE = 0;
                    decimal decF = 0;
                    decimal decG = 0;
                    decimal decH = 0;
                    decimal decI = 0;
                    decimal decJ = 0;

                    if (TBFormulas.TableFormulaCall.Contains("SQL"))
                    {
                        MessageBox.Show("SQL extract", "Not available to be tested", MessageBoxButtons.OK);
                    }
                    else
                    {
                        if (TBFormulas.CalcName.Contains("xx") || TBFormulas.TableFormulaCall.Contains("Concat"))
                        {
                        }
                        else
                        {
                            if (grdActiveSheet.Columns.Contains(TBFormulas.A))
                            {
                                decA = Convert.ToDecimal(grdActiveSheet[TBFormulas.A, rownr].Value);
                            }
                            else
                                if (dict.ContainsKey(TBFormulas.A))
                                {
                                    decA = Convert.ToDecimal(dict[TBFormulas.A]);
                                }
                                else
                                {
                                    decA = 9999;
                                }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.B))
                            {
                                decB = Convert.ToDecimal(grdActiveSheet[TBFormulas.B, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.B))
                                {
                                    decB = Convert.ToDecimal(dict[TBFormulas.B]);
                                }
                                else
                                {
                                    decB = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.C))
                            {
                                decC = Convert.ToDecimal(grdActiveSheet[TBFormulas.C, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.C))
                                {
                                    decC = Convert.ToDecimal(dict[TBFormulas.C]);
                                }
                                else
                                {
                                    decC = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.D))
                            {
                                decD = Convert.ToDecimal(grdActiveSheet[TBFormulas.D, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.D))
                                {
                                    decD = Convert.ToDecimal(dict[TBFormulas.D]);
                                }
                                else
                                {
                                    decD = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.E))
                            {
                                decE = Convert.ToDecimal(grdActiveSheet[TBFormulas.E, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.E))
                                {
                                    decE = Convert.ToDecimal(dict[TBFormulas.E]);
                                }
                                else
                                {
                                    decE = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.F))
                            {
                                decF = Convert.ToDecimal(grdActiveSheet[TBFormulas.F, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.F))
                                {
                                    decF = Convert.ToDecimal(dict[TBFormulas.F]);
                                }
                                else
                                {
                                    decF = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.G))
                            {
                                decG = Convert.ToDecimal(grdActiveSheet[TBFormulas.G, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.G))
                                {
                                    decG = Convert.ToDecimal(dict[TBFormulas.G]);
                                }
                                else
                                {
                                    decG = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.H))
                            {
                                decH = Convert.ToDecimal(grdActiveSheet[TBFormulas.H, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.H))
                                {
                                    decH = Convert.ToDecimal(dict[TBFormulas.H]);
                                }
                                else
                                {
                                    decH = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.I))
                            {
                                decI = Convert.ToDecimal(grdActiveSheet[TBFormulas.I, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.I))
                                {
                                    decI = Convert.ToDecimal(dict[TBFormulas.I]);
                                }
                                else
                                {
                                    decI = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.J))
                            {
                                decJ = Convert.ToDecimal(grdActiveSheet[TBFormulas.J, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.J))
                                {
                                    decJ = Convert.ToDecimal(dict[TBFormulas.J]);
                                }
                                else
                                {
                                    decJ = 9999;
                                }
                            }

                            MessageBox.Show("Database Name:     " + TB.DBName + '\n' + "Table Name:           " + TB.TBName + '\n' + "Calculation Name:   " +
                            TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + "A =             " +
                            TBFormulas.A + "   =   " + decA + '\n' + "B =             " + TBFormulas.B + "   =   " + decB + '\n' + "C =             " +
                            TBFormulas.C + "   =   " + decC + '\n' + "D =             " +
                            TBFormulas.D + "   =   " + decD + '\n' + "E =             " +
                            TBFormulas.E + "   =   " + decE + '\n' + "F =             " +
                            TBFormulas.F + "   =   " + decF + '\n' + "G =             " +
                            TBFormulas.G + "   =   " + decG + '\n' + "H =             " +
                            TBFormulas.H + "   =   " + decH + '\n' + "I  =            " +
                            TBFormulas.I + "   =    " + decI + '\n' + "J  =            " +
                            TBFormulas.J + "   =    " + decJ, "FORMULA DETAILS - of selected value: ---------------------------------------------------->        ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                else
                {
                    this.Cursor = Cursors.Arrow;
                    MessageBox.Show("Calculation does not exist anymore. Delete the column.", "ERROR", MessageBoxButtons.OK);
                }
            }
            this.Cursor = Cursors.Arrow;
        }

        private void loadDict(DataTable _datatable)
        {
            foreach (DataRow _row in _datatable.Rows)
            {
                string str = _row[0].ToString().Trim();
                if (dict.ContainsKey(str))
                {
                    dict.Remove(str);
                    dict.Add(str, _row[1].ToString().Trim());
                }
                else
                {
                    dict.Add(str, _row[1].ToString().Trim());
                }
            }
            dict.Remove("X");
            dict.Add("X", "0");

        }

        private void buildDisplaySQL(string strwhere, decimal decValue)
        {
            string strSQL = "";

            strSQL = "Database Name:     " + TB.DBName + '\n' + "Table Name:           " + TB.TBName + '\n' + "Calculation Name:   " +
                         TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + TBFormulas.A + TBFormulas.B + TBFormulas.C + TBFormulas.D + TBFormulas.E + TBFormulas.F + TBFormulas.G + TBFormulas.H + " " + strwhere;
            strSQL = strSQL.Replace("#", "").Replace(":and:", "and").Replace(" from ", "\n from ").Replace(" and ", "\n and ").Replace(" where ", "\n where ");

            if (strSQL.Trim().Contains(" as t1XXXXX"))
            {
                string A = strSQL.Trim().Substring(0, strSQL.Trim().IndexOf("@ ) as t1XXXXX") + 1);
                string B = strSQL.Trim().Substring(strSQL.Trim().IndexOf("@ ) as t1XXXXX") + 1);
                strSQL = A + " and period = '" + BusinessLanguage.Period.Trim() + "'" + B;
            }

            General.textTestSQL = strSQL;
            scrQuerySQL testsql = new scrQuerySQL();
            testsql.TestSQL(Base.DBConnection, General, Base.DBConnectionString);
            testsql.ShowDialog();

        }

        private void userProfile_Click(object sender, EventArgs e)
        {
            scrProfile userProfile = new scrProfile();
            userProfile.FormLoad(BusinessLanguage, BaseConn);
            userProfile.Show();
        }

        private void grantAccessToolStripMenuItem_Click(object sender, EventArgs e)
        {
            scrSecurity useraccess = new scrSecurity();
            useraccess.userAccessLoad(myConn, Base, TB, BusinessLanguage.Userid, strServerPath.ToString().ToUpper());
            useraccess.Show();
        }

        private void btn0_Click(object sender, EventArgs e)
        {

            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "0";

        }

        private void btn1_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "1";
        }

        private void btn2_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "2";
        }

        private void btn3_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "3";
        }

        private void btn4_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "4";
        }

        private void btn5_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "5";
        }

        private void btn6_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "6";
        }

        private void btn7_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "7";
        }

        private void btn8_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "8";
        }

        private void btn9_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "9";
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = "";
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            DataTable searchEmpl = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from ClockedShifts where employee_no like '%" + txtSearchEmpl.Text.Trim() + "%'");

            if (searchEmpl.Rows.Count > 0)
            {
                //amp
                string strLSH = Clocked.Rows[0]["LSH"].ToString().Trim();
                DateTime LSH = Convert.ToDateTime(strLSH);
                string Mnth = string.Empty;
                string Day = string.Empty;
                foreach (DataColumn dc in searchEmpl.Columns)
                {
                    if (dc.Caption.Substring(0, 3) == "DAY")
                    {
                        double d = Convert.ToDouble(dc.Caption.Substring(3).Trim());
                        string strTemp = Clocked.Rows[0]["FSH"].ToString().Trim();
                        DateTime temp = Convert.ToDateTime(strTemp);
                        temp = temp.AddDays(d);
                        if (temp > LSH)  //remember the days start at 0
                        {
                            if (Convert.ToString(temp.Day).Length < 2)
                            {
                                Day = "0" + Convert.ToString(temp.Day);
                            }
                            else
                            {
                                Day = Convert.ToString(temp.Day);
                            }
                            if (Convert.ToString(temp.Month).Length < 2)
                            {
                                Mnth = "0" + Convert.ToString(temp.Month);
                            }
                            else
                            {
                                Mnth = Convert.ToString(temp.Month);
                            }
                            searchEmpl.Columns[dc.Caption].ColumnName = "x" + Day + '-' + Mnth;
                        }
                        else
                        {
                            if (Convert.ToString(temp.Day).Length < 2)
                            {
                                Day = "0" + Convert.ToString(temp.Day);
                            }
                            else
                            {
                                Day = Convert.ToString(temp.Day);
                            }
                            if (Convert.ToString(temp.Month).Length < 2)
                            {
                                Mnth = "0" + Convert.ToString(temp.Month);
                            }
                            else
                            {
                                Mnth = Convert.ToString(temp.Month);
                            }
                            searchEmpl.Columns[dc.Caption].ColumnName = "d" + Day + '-' + Mnth;
                        }
                    }
                }
            }
            grdClocked.DataSource = searchEmpl;
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            grdClocked.DataSource = Clocked;
        }

        private void grdActiveSheet_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //Get calc name
            this.Cursor = Cursors.WaitCursor;
            int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
            int rownr = grdActiveSheet.CurrentCell.RowIndex;
            TBFormulas.CalcName = grdActiveSheet.Columns[columnnr].HeaderText;

            //Check if it is a calculated column
            string FormulaTableName = string.Empty;

            if (TB.TBName.Trim().ToUpper().Contains("EARN"))
            {
                FormulaTableName = TB.TBName.Trim().Substring(0, TB.TBName.Trim().ToUpper().IndexOf("20"));
            }
            else
            {
                FormulaTableName = TB.TBName;
            }

            object intCount = Analysis.countcalcbyname(TB.DBName + BusinessLanguage.Period.Trim(), FormulaTableName,
                                                       TBFormulas.CalcName.Trim(), Base.AnalysisConnectionString);

            if ((int)intCount > 0)
            {
                //It is a calculated column.
                DataTable dtFormula = Analysis.GetCalcDetailsDCript(TB.DBName + BusinessLanguage.Period.Trim(), FormulaTableName,
                                                                    TBFormulas.CalcName, Base.AnalysisConnectionString);
                //Extract the formula details:
                decimal decValue = 0;
                try
                {
                    decValue = Convert.ToDecimal(grdActiveSheet.CurrentCell.Value);
                }
                catch
                {
                    decValue = 0;
                }

                //Extract Factors
                //TB.extractDBTableIntoDataTable(Base.DBConnectionString, "FACTORS"," Where period = '" + BusinessLanguage.Period + "'");
                DataTable dtFactors = TB.createDataTableWithAdapter(Base.DBConnectionString,
                                    "Select Varname,Varvalue from FACTORS where period = '" + BusinessLanguage.Period + "'");
                dict.Clear();
                loadDict(dtFactors);

                if (dtFormula.Rows.Count > 0)
                {
                    TBFormulas.A = dtFormula.Rows[0]["A"].ToString().Trim();
                    TBFormulas.B = dtFormula.Rows[0]["B"].ToString().Trim();
                    TBFormulas.C = dtFormula.Rows[0]["C"].ToString().Trim();
                    TBFormulas.D = dtFormula.Rows[0]["D"].ToString().Trim();
                    TBFormulas.E = dtFormula.Rows[0]["E"].ToString().Trim();
                    TBFormulas.F = dtFormula.Rows[0]["F"].ToString().Trim();
                    TBFormulas.G = dtFormula.Rows[0]["G"].ToString().Trim();
                    TBFormulas.H = dtFormula.Rows[0]["H"].ToString().Trim();
                    TBFormulas.I = dtFormula.Rows[0]["I"].ToString().Trim();
                    TBFormulas.J = dtFormula.Rows[0]["J"].ToString().Trim();
                    TBFormulas.TableFormulaCall = dtFormula.Rows[0]["FORMULA_CALL"].ToString().Trim();
                    decimal decA = 0;
                    decimal decB = 0;
                    decimal decC = 0;
                    decimal decD = 0;
                    decimal decE = 0;
                    decimal decF = 0;
                    decimal decG = 0;
                    decimal decH = 0;
                    decimal decI = 0;
                    decimal decJ = 0;

                    if (TBFormulas.TableFormulaCall.Contains("SQL"))
                    {
                        string strWhere = " ";
                        for (int i = 0; i < grdActiveSheet.Columns.Count - 1; i++)
                        {

                            strWhere = strWhere.Trim() + " and t1." + grdActiveSheet.Columns[i].HeaderText.Trim() +
                                       " = '" + (string)(grdActiveSheet[i, e.RowIndex].Value).ToString().Trim() + "'";

                        }

                        buildDisplaySQL(strWhere, decValue);
                    }
                    else
                    {
                        if (TBFormulas.CalcName.Contains("xx") || TBFormulas.TableFormulaCall.Contains("Concat"))
                        {
                        }
                        else
                        {

                            if (grdActiveSheet.Columns.Contains(TBFormulas.A))
                            {
                                decA = Convert.ToDecimal(grdActiveSheet[TBFormulas.A, rownr].Value);
                            }
                            else
                                if (dict.ContainsKey(TBFormulas.A))
                                {
                                    decA = Convert.ToDecimal(dict[TBFormulas.A]);
                                }
                                else
                                {
                                    decA = 9999;
                                }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.B))
                            {
                                decB = Convert.ToDecimal(grdActiveSheet[TBFormulas.B, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.B))
                                {
                                    decB = Convert.ToDecimal(dict[TBFormulas.B]);
                                }
                                else
                                {
                                    decB = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.C))
                            {
                                decC = Convert.ToDecimal(grdActiveSheet[TBFormulas.C, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.C))
                                {
                                    decC = Convert.ToDecimal(dict[TBFormulas.C]);
                                }
                                else
                                {
                                    decC = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.D))
                            {
                                decD = Convert.ToDecimal(grdActiveSheet[TBFormulas.D, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.D))
                                {
                                    decD = Convert.ToDecimal(dict[TBFormulas.D]);
                                }
                                else
                                {
                                    decD = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.E))
                            {
                                decE = Convert.ToDecimal(grdActiveSheet[TBFormulas.E, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.E))
                                {
                                    decE = Convert.ToDecimal(dict[TBFormulas.E]);
                                }
                                else
                                {
                                    decE = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.F))
                            {
                                decF = Convert.ToDecimal(grdActiveSheet[TBFormulas.F, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.F))
                                {
                                    decF = Convert.ToDecimal(dict[TBFormulas.F]);
                                }
                                else
                                {
                                    decF = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.G))
                            {
                                decG = Convert.ToDecimal(grdActiveSheet[TBFormulas.G, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.G))
                                {
                                    decG = Convert.ToDecimal(dict[TBFormulas.G]);
                                }
                                else
                                {
                                    decG = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.H))
                            {
                                decH = Convert.ToDecimal(grdActiveSheet[TBFormulas.H, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.H))
                                {
                                    decH = Convert.ToDecimal(dict[TBFormulas.H]);
                                }
                                else
                                {
                                    decH = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.I))
                            {
                                decI = Convert.ToDecimal(grdActiveSheet[TBFormulas.I, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.I))
                                {
                                    decI = Convert.ToDecimal(dict[TBFormulas.I]);
                                }
                                else
                                {
                                    decI = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.J))
                            {
                                decJ = Convert.ToDecimal(grdActiveSheet[TBFormulas.J, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.J))
                                {
                                    decJ = Convert.ToDecimal(dict[TBFormulas.J]);
                                }
                                else
                                {
                                    decJ = 9999;
                                }
                            }

                            MessageBox.Show("Database Name:     " + TB.DBName + BusinessLanguage.Period.Trim() + '\n' + "Table Name:           " + FormulaTableName + '\n' + "Calculation Name:   " +
                            TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + "A =             " +
                            TBFormulas.A + "   =   " + decA + '\n' + "B =             " + TBFormulas.B + "   =   " + decB + '\n' + "C =             " +
                            TBFormulas.C + "   =   " + decC + '\n' + "D =             " +
                            TBFormulas.D + "   =   " + decD + '\n' + "E =             " +
                            TBFormulas.E + "   =   " + decE + '\n' + "F =             " +
                            TBFormulas.F + "   =   " + decF + '\n' + "G =             " +
                            TBFormulas.G + "   =   " + decG + '\n' + "H =             " +
                            TBFormulas.H + "   =   " + decH + '\n' + "I  =            " +
                            TBFormulas.I + "   =    " + decI + '\n' + "J  =            " +
                            TBFormulas.J + "   =    " + decJ, "FORMULA DETAILS - of selected value: ---------------------------------------------------->        ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                else
                {
                    this.Cursor = Cursors.Arrow;
                    MessageBox.Show("Calculation does not exist anymore. Delete the column.", "ERROR", MessageBoxButtons.OK);
                }
            }
            this.Cursor = Cursors.Arrow;
        }

        private void grdGangLink_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            int intRow = e.RowIndex;
            if (e.RowIndex < 0)
            {

            }
            else
            {


                cboGangLinkGang.Text = grdGangLink["GANG", e.RowIndex].Value.ToString().Trim();
                cboGangLinkWorkplace.Text = grdGangLink["WORKPLACE", e.RowIndex].Value.ToString().Trim();
                cboGangLinkSafetyInd.Text = grdGangLink["SAFETYIND", e.RowIndex].Value.ToString().Trim();
                cboGangLinkGangType.Text = grdGangLink["GANGTYPE", e.RowIndex].Value.ToString().Trim(); 
            }
            #region Trigger output
            //load the CURRENT values into dictionaries before the update 
            // that was loaded in tabInfo_SelectedIndexChanged
            dictPrimaryKeyValues.Clear();
            dictGridValues.Clear();

            foreach (string s in lstPrimaryKeyColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictPrimaryKeyValues.Add(s, grdGangLink[s, e.RowIndex].Value.ToString().Trim());
                }
            }

            foreach (string s in lstTableColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictGridValues.Add(s, grdGangLink[s, e.RowIndex].Value.ToString().Trim());
                }
            }
            #endregion
        }

        private void writeAudit(string tablename, string function, string fieldname, string oldValue, string newValue)
        {
            string PK = string.Empty;
            foreach (string key in dictPrimaryKeyValues.Keys)
            {
                PK = PK + "<" + key.Trim() + "=" + dictPrimaryKeyValues[key] + ">";
            }

            DataTable audit = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "AUDIT");
            audit.Clear();

            DataRow dr = audit.NewRow();
            dr["Type"] = function.Substring(0, 1);
            dr["TableName"] = tablename;
            dr["PK"] = PK;
            dr["FieldName"] = fieldname;
            dr["OldValue"] = oldValue;
            dr["NewValue"] = newValue;
            dr["UpdateDate"] = DateTime.Today.ToLongDateString();
            dr["UserName"] = BusinessLanguage.Userid;

            audit.Rows.Add(dr);
            audit.AcceptChanges();

            TB.saveCalculations2(audit, Base.DBConnectionString, " where type = 'x'", "AUDIT");
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            txtSelectedSection.Text = listBox2.SelectedItem.ToString().Trim();

            Base.Section = txtSelectedSection.Text.Trim();    //xxxxxxxxxxxxxxxxxx

            int intRowPosition = 0;
            for (int i = 0; i <= Calendar.Rows.Count - 1; i++)
            {
                if (Calendar.Rows[i]["SECTION"].ToString().Trim() == txtSelectedSection.Text.Trim() &&
                    Calendar.Rows[i]["PERIOD"].ToString().Trim() == BusinessLanguage.Period)
                {
                    intRowPosition = i;
                }
            }


            cboOffDaysSection.Text = txtSelectedSection.Text.Trim();
            cboOffDaysGang.Text = @"DUMMY";
            label15.Text = listBox2.SelectedItem.ToString().Trim();
            label30.Text = BusinessLanguage.Period;
            strWhere = "where section = '" + listBox2.SelectedItem.ToString().Trim() + "' and period = '" + BusinessLanguage.Period + "'";
              
            loadMO();

            evaluateStatus();
            evaluateSurvey();
            evaluateClockedShifts();
            evaluateLabour();
            evaluateMiners();
            evaluateGangLinking();
            evaluatePayroll();
            evaluateEmployeePenalties();
            evaluateOffDays();
            evaluateDrillers();
            this.Cursor = Cursors.Arrow;

            extractMeasuringDates();

        }

        private void extractMeasuringDates()
        {

            IEnumerable<DataRow> query1 = from locks in Calendar.AsEnumerable()
                                          where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                          where locks.Field<string>("PERIOD").Trim() == BusinessLanguage.Period.Trim()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();
            dateTimePicker1.Value = Convert.ToDateTime(temp.Rows[0]["FSH"].ToString().Trim());
            dateTimePicker2.Value = Convert.ToDateTime(temp.Rows[0]["LSH"].ToString().Trim());
            strMonthShifts = temp.Rows[0]["MONTHSHIFTS"].ToString().Trim();
             
            lstOffDayValue.Items.Clear();
            //Load the possible dates that the user can select in this measuring period for the offday calendar
            for (DateTime i = dateTimePicker1.Value; i <= dateTimePicker2.Value; i = i.AddDays(1))
            {
                lstOffDayValue.Items.Add(i.ToString("yyyy-MM-dd"));
            }

        }

        private void btnEmployeeCalc_Click(object sender, EventArgs e)
        {

            string strSQL = "BEGIN transaction; Delete from monitor ; commit transaction;";
            TB.InsertData(Base.DBConnectionString, strSQL);

        }

        private void dataSort_Click(object sender, EventArgs e)
        {

        }

        private void DataPrintCrewPrint_Click(object sender, EventArgs e)
        {

        }

        private void btnUpdate_Click_1(object sender, EventArgs e)
        {
            int intRow = 0;
            int intColumn = 0;

            string strSQL = "";

            switch (tabInfo.SelectedTab.Name)
            {
                case "tabAbnormal":
                    #region tabAbnormal

                    //HJ
                    if (cboAbnormalWorkplace.Text.Trim().Length != 0 &&
                        cboAbnormalLevel.Text.Trim().Length != 0 && cboAbnormalType.Text.Trim().Length != 0 &&
                        txtAbnormalValue.Text.Trim().Length != 0)
                    {

                        intRow = grdAbnormal.CurrentCell.RowIndex;
                        intColumn = grdAbnormal.CurrentCell.ColumnIndex;

                        strSQL = "BEGIN transaction; Update Abnormal set Section = '" + txtSelectedSection.Text.Trim() + "', Period = '" + txtPeriod.Text.Trim() +
                                             "', Workplace = '" + cboAbnormalWorkplace.Text +
                                             "', AbnormalLevel = '" + cboAbnormalLevel.Text.Trim() + "', AbnormalType = '" + cboAbnormalType.Text.Trim() +
                                             "', AbnormalValue = '" + txtAbnormalValue.Text.Trim() + "'" +
                                             " Where Section = '" + grdAbnormal["SECTION", intRow].Value.ToString().Trim() +
                                             "' and Period = '" + grdAbnormal["PERIOD", intRow].Value.ToString().Trim() +
                                            "' and Workplace = '" + grdAbnormal["WORKPLACE", intRow].Value.ToString().Trim() +
                                             "' and AbnormalLevel = '" + grdAbnormal["ABNORMALLEVEL", intRow].Value.ToString().Trim() +
                                             "' and AbnormalType = '" + grdAbnormal["ABNORMALTYPE", intRow].Value.ToString().Trim() +
                                             "' and AbnormalValue = '" + grdAbnormal["ABNORMALVALUE", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        if (grdAbnormal[2, intRow].Value.ToString().Trim() != "XXX")
                        {
                            grdAbnormal["Section", intRow].Value = txtSelectedSection.Text.Trim();
                            grdAbnormal["Section", intRow].Style.BackColor = Color.Lavender;
                            grdAbnormal["Period", intRow].Value = txtPeriod.Text.Trim();
                            grdAbnormal["Period", intRow].Style.BackColor = Color.Lavender;
                            grdAbnormal["Workplace", intRow].Value = cboAbnormalWorkplace.Text.Trim();
                            grdAbnormal["Workplace", intRow].Style.BackColor = Color.Lavender;
                            grdAbnormal["AbnormalLevel", intRow].Value = cboAbnormalLevel.Text.Trim();
                            grdAbnormal["AbnormalLevel", intRow].Style.BackColor = Color.Lavender;
                            grdAbnormal["AbnormalType", intRow].Value = cboAbnormalType.Text.Trim();
                            grdAbnormal["AbnormalLevel", intRow].Style.BackColor = Color.Lavender;
                            grdAbnormal["AbnormalValue", intRow].Value = txtAbnormalValue.Text.Trim();
                            grdAbnormal["AbnormalValue", intRow].Style.BackColor = Color.Lavender;

                            TB.InsertData(Base.DBConnectionString, strSQL);
                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabSurvey":
                    #region tabSurvey

                    if (txtSurveyWorkplace.Text.Trim().Length != 0 &&
                        txtSurveySwpDist.Text.Trim().Length != 0 )
                    {

                        intRow = grdSurvey.CurrentCell.RowIndex;
                        intColumn = grdSurvey.CurrentCell.ColumnIndex;

                        strSQL = "BEGIN transaction; Update Survey set SweepingDistance = '" + txtSurveySwpDist.Text.Trim() +
                                 "' Where Section = '" + grdSurvey["SECTION", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdSurvey["PERIOD", intRow].Value.ToString().Trim() +
                                 "' and Workplace = '" + grdSurvey["WORKPLACE", intRow].Value.ToString().Trim() +
                                 "' and SweepingDistance = '" + grdSurvey["SWEEPINGDISTANCE", intRow].Value.ToString().Trim() + 
                                 "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdSurvey["Sweepingdistance", intRow].Value = txtSurveySwpDist.Text.Trim();

                        for (int i = 0; i <= grdSurvey.Columns.Count - 1; i++)
                        {
                            grdSurvey[i, intRow].Style.BackColor = Color.Coral;
                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabMiners":
                    #region tabMiners

                    //HJ
                    if (cboMinersGangNo.Text.Trim().Length != 0 && cboNames.Text.Trim().Length != 0 &&
                        cboDesignation.Text.Trim().Length != 0 && txtPayShifts.Text.Trim().Length != 0 &&
                        txtAwops.Text.Trim().Length != 0 &&
                        txtMinersSafetyInd.Text.Trim().Length != 0)
                    {
                        intRow = grdMiners.CurrentCell.RowIndex;
                        intColumn = grdMiners.CurrentCell.ColumnIndex;

                        string strName = string.Empty;
                        string strDesignation = string.Empty;
                        string strDesignationDesc = string.Empty;

                        if (cboNames.Text.Contains("-"))
                        {
                            strName = cboNames.Text.Substring(0, cboNames.Text.IndexOf("-")).Trim();
                        }
                        else
                        {
                            strName = cboNames.Text.Trim();
                        }

                        if (cboDesignation.Text.Contains("-"))
                        {
                            strDesignation = cboDesignation.Text.Substring(0, cboDesignation.Text.IndexOf("-")).Trim();
                            strDesignationDesc = cboDesignation.Text.Substring((cboDesignation.Text.IndexOf("-")) + 3);
                        }
                        else
                        {
                            strDesignation = cboDesignation.Text.Trim();
                            strDesignationDesc = cboDesignation.Text.Trim();
                        }

                        int rowindex = grdMiners.CurrentCell.RowIndex;
                        strSQL = "BEGIN transaction; Update Miners set Period = '" + txtPeriod.Text.Trim() +
                                 "', miningtype = '" + BusinessLanguage.MiningType.Trim() +
                                 "', bonustype = '" + BusinessLanguage.BonusType.Trim() +
                                  "', bussunit = '" + BusinessLanguage.BussUnit.Trim() +
                                 "', Employee_No = '" + cboNames.Text.Trim() + 
                                 "', Designation = '" + strDesignation +
                                 "', Awop_Shifts = '" + txtAwops.Text.Trim() +
                                 "', Payshifts = '" + txtPayShifts.Text.Trim() +
                                 "', Gang = '" + cboMinersGangNo.Text.Trim() +
                                 "', Employee_name = '" + cboMinersEmpName.Text.Trim() +
                                 "', Shifts_Worked = '" + txtADTeamShifts.Text.Trim() +
                                 "', Designation_desc = '" + strDesignationDesc +
                                 "', SafetyInd = '" + txtMinersSafetyInd.Text.Trim() + "'" +
                                 " Where Section = '" + grdMiners["SECTION", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdMiners["PERIOD", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + grdMiners["EMPLOYEE_NO", intRow].Value.ToString().Trim() +
                                 "' and Gang = '" + grdMiners["GANG", intRow].Value.ToString() +
                                 "' and SafetyInd = '" + grdMiners["SAFETYIND", intRow].Value.ToString() +
                                 "' and Awop_Shifts = '" + grdMiners["AWOP_SHIFTS", intRow].Value.ToString().Trim() +
                                 "' and Payshifts = '" + grdMiners["PAYSHIFTS", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        for (int i = 0; i <= grdMiners.Columns.Count - 1; i++)
                        {
                            grdMiners[i, intRow].Style.BackColor = Color.Lavender;
                        }

                        grdMiners["PERIOD", intRow].Value = txtPeriod.Text.Trim();
                        grdMiners["GANG", intRow].Value = cboMinersGangNo.Text.Trim();
                        grdMiners["EMPLOYEE_NO", intRow].Value = cboNames.Text.Trim();
                        grdMiners["EMPLOYEE_NAME", intRow].Value = cboMinersEmpName.Text.Trim();
                        grdMiners["DESIGNATION", intRow].Value = cboDesignation.Text.Trim();
                        grdMiners["DESIGNATION_DESC", intRow].Value = strDesignationDesc;
                        grdMiners["SHIFTS_WORKED", intRow].Value = txtADTeamShifts.Text.Trim();
                        grdMiners["PAYSHIFTS", intRow].Value = txtPayShifts.Text.Trim();
                        grdMiners["SAFETYIND", intRow].Value = txtMinersSafetyInd.Text.Trim();
                        grdMiners["AWOP_SHIFTS", intRow].Value = txtAwops.Text.Trim();

                        grdMiners.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion

                case "tabLabour":
                    #region tabLabour

                    //HJ
                    if (txtEmployeeNo.Text.Trim().Length > 0
                        && txtEmployeeName.Text.Trim().Length > 0
                        && cboBonusShiftsGang.Text.Trim().Length > 0
                        && cboBonusShiftsWageCode.Text.Trim().Length > 0
                        && cboBonusShiftsResponseCode.Text.Trim().Length > 0
                        && txtShifts.Text.Trim().Length > 0
                        && txtAwop.Text.Trim().Length > 0
                        && txtStrikeShifts.Text.Trim().Length > 0
                        && cboDrillerInd.Text.Trim().Length > 0
                        && txtDrillerShifts.Text.Trim().Length > 0)
                    {

                        intRow = grdLabour.CurrentCell.RowIndex;

                        string strWagecode = Convert.ToString(grdLabour["WAGECODE", intRow].Value);
                        string strEmployeeName = Convert.ToString(grdLabour["EMPLOYEE_NAME", intRow].Value);
                        string strGang = Convert.ToString(grdLabour["GANG", intRow].Value);
                        string strResponseCo = Convert.ToString(grdLabour["LINERESPCODE", intRow].Value);
                        string strShiftsWorked = Convert.ToString(grdLabour["SHIFTS_WORKED", intRow].Value);
                        string strAwops = Convert.ToString(grdLabour["AWOP_SHIFTS", intRow].Value);
                        string strStrikes = Convert.ToString(grdLabour["STRIKE_SHIFTS", intRow].Value);
                        string strDrillerInd = Convert.ToString(grdLabour["DRILLERIND", intRow].Value);
                        string strDrillerShifts = Convert.ToString(grdLabour["DRILLERSHIFTS", intRow].Value);
                        string strTeamLeadind = Convert.ToString(grdLabour["TEAMLEADERIND", intRow].Value);

                        strSQL = "Update bonusshifts set wagecode = '" + cboBonusShiftsWageCode.Text.Trim() +
                                 "' , Gang = '" + cboBonusShiftsGang.Text.Trim() +
                                 "' , Linerespcode = '" + cboBonusShiftsResponseCode.Text.Trim() +
                                 "' , Shifts_Worked = '" + txtShifts.Text.Trim() +
                                 "' , Awop_Shifts = '" + txtAwop.Text.Trim() +
                                 "' , Strike_Shifts = '" + txtStrikeShifts.Text.Trim() +
                                 "' , DrillerInd = '" + cboDrillerInd.Text.Trim() +
                                 "' , DrillerShifts = '" + txtDrillerShifts.Text.Trim() +
                                 "' , TEAMLEADERIND = '" + cboTLeader.Text.Trim() +
                                 "' where employee_no = '" + grdLabour["Employee_No", intRow].Value +
                                 "' and Linerespcode = '" + grdLabour["Linerespcode", intRow].Value +
                                 "' and Employee_name = '" + grdLabour["Employee_Name", intRow].Value +
                                 "' and WageCode = '" + grdLabour["WageCode", intRow].Value +
                                 "' and Period = '" + grdLabour["Period", intRow].Value +   
                                 "' and Gang = '" + grdLabour["Gang", intRow].Value + "'";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdLabour["WAGECODE", intRow].Value = cboBonusShiftsWageCode.Text.Trim();
                        grdLabour["GANG", intRow].Value = cboBonusShiftsGang.Text.Trim();
                        grdLabour["LINERESPCODE", intRow].Value = cboBonusShiftsResponseCode.Text.Trim();
                        grdLabour["SHIFTS_WORKED", intRow].Value = txtShifts.Text.Trim();
                        grdLabour["AWOP_SHIFTS", intRow].Value = txtAwop.Text.Trim();
                        grdLabour["STRIKE_SHIFTS", intRow].Value = txtStrikeShifts.Text.Trim();
                        grdLabour["DrillerInd", intRow].Value = cboDrillerInd.Text.Trim();
                        grdLabour["DRILLERSHIFTS", intRow].Value = txtDrillerShifts.Text.Trim();
                        grdLabour["TEAMLEADERIND", intRow].Value = cboTLeader.Text.Trim();

                        for (int i = 0; i <= grdLabour.Columns.Count - 1; i++)
                        {
                            grdLabour[i, intRow].Style.BackColor = Color.Lavender;
                        }
                        //    }
                        //    else
                        //    {
                        //        MessageBox.Show("Invalid password.", "Error", MessageBoxButtons.OK);
                        //    }
                        //}

                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabGangLinking":
                    #region tabGangLink

                    if (cboGangLinkSafetyInd.Text.Trim().Length > 0 &&
                        cboGangLinkGangType.Text.Trim().Length > 0 &&
                        cboGangLinkWorkplace.Text.Trim().Length > 0 &&
                        cboGangLinkGang.Text.Trim().Length > 0)
                    {
                        intRow = grdGangLink.CurrentCell.RowIndex;

                        string strGang = Convert.ToString(grdGangLink["Gang", intRow].Value);
                        string strWorkplace = Convert.ToString(grdGangLink["Workplace", intRow].Value);
                        string strGangType = Convert.ToString(grdGangLink["GangType", intRow].Value);
                        string strSafetyInd = Convert.ToString(grdGangLink["Safetyind", intRow].Value);

                        strSQL = "Update ganglink set  Workplace = '" + cboGangLinkWorkplace.Text.Trim() +
                                 "' , Gangtype = '" + cboGangLinkGangType.Text.Trim() +
                                 "' , Gang = '" + cboGangLinkGang.Text.Trim() +
                                 "' , SafetyInd = '" + cboGangLinkSafetyInd.Text.Trim() +
                                 "' where gang = '" + grdGangLink["Gang", intRow].Value +
                                 "' and workplace = '" + grdGangLink["Workplace", intRow].Value +
                                 "' and period = '" + grdGangLink["Period", intRow].Value +         //xxxxxxxxxxx
                                 "' and section = '" + grdGangLink["Section", intRow].Value +         //xxxxxxxxxxx
                                 "' and gangtype = '" + grdGangLink["Gangtype", intRow].Value + "'";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdGangLink["Gang", intRow].Value = cboGangLinkGang.Text.Trim();
                        grdGangLink["Workplace", intRow].Value = cboGangLinkWorkplace.Text.Trim();
                        grdGangLink["SafetyInd", intRow].Value = cboGangLinkSafetyInd.Text.Trim();
                        grdGangLink["Gangtype", intRow].Value = cboGangLinkGangType.Text.Trim();

                        for (int i = 0; i <= grdGangLink.Columns.Count - 1; i++)
                        {
                            grdGangLink[i, intRow].Style.BackColor = Color.Lavender;
                        }
                        //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                        foreach (string s in lstTableColumns)
                        {
                            if (dictGridValues[s] == grdGangLink[s, intRow].Value.ToString().Trim())
                            {

                            }
                            else
                            {
                                //Write out to audit log
                                writeAudit("GANGLINK", "U - Update", s, dictGridValues[s], grdGangLink[s, intRow].Value.ToString().Trim());

                            }

                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties

                    //HJ
                    if (cboEmplPenEmployeeNo.Text.Trim().Length != 0 &&
                        txtPenaltyValue.Text.Trim().Length != 0 && cboPenaltyInd.Text.Trim().Length != 0)
                    {

                        intRow = grdEmplPen.CurrentCell.RowIndex;
                        intColumn = grdEmplPen.CurrentCell.ColumnIndex;

                        if (cboEmplPenEmployeeNo.Text.Contains("-"))
                        {
                            strName = cboEmplPenEmployeeNo.Text.Substring(0, cboEmplPenEmployeeNo.Text.IndexOf("-")).Trim();
                        }
                        else
                        {
                            strName = cboEmplPenEmployeeNo.Text.Trim();
                        }

                        strSQL = "BEGIN transaction; Update EmployeePenalties set Period = '" + txtPeriod.Text.Trim() +
                                             "', Employee_No = '" + strName + "', PenaltyValue = '" + txtPenaltyValue.Text.Trim() +
                                             "', PenaltyInd = '" + cboPenaltyInd.Text.Trim() + "'" +
                                             " Where Section = '" + grdEmplPen["SECTION", intRow].Value.ToString().Trim() +
                                             "' and Period = '" + grdEmplPen["PERIOD", intRow].Value.ToString().Trim() +
                                             "' and Employee_No = '" + grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() +
                                             "' and PenaltyValue = '" + grdEmplPen["PENALTYVALUE", intRow].Value.ToString().Trim() +
                                             "' and PenaltyInd = '" + grdEmplPen["PENALTYIND", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        if (grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXXXXXXXXXXX")
                        {
                            grdEmplPen["Section", intRow].Value = txtSelectedSection.Text.Trim();
                            grdEmplPen["Section", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["Period", intRow].Value = txtPeriod.Text.Trim();
                            grdEmplPen["Period", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["Employee_No", intRow].Value = cboEmplPenEmployeeNo.Text.Trim();
                            grdEmplPen["Employee_No", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["PenaltyValue", intRow].Value = txtPenaltyValue.Text.Trim();
                            grdEmplPen["PenaltyValue", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["PenaltyInd", intRow].Value = cboPenaltyInd.Text.Trim();
                            grdEmplPen["PenaltyInd", intRow].Style.BackColor = Color.Lavender;

                            TB.InsertData(Base.DBConnectionString, strSQL);
                            clearAllCalcValues("Ganglink", txtSelectedSection.Text.Trim());
                            clearAllCalcValues("Miners", txtSelectedSection.Text.Trim());
                            clearAllCalcValues("Bonusshifts", txtSelectedSection.Text.Trim());

                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabConfig":
                    #region tabConfiguration

                    //HJ
                    if (grdConfigs[0, intRow].Value.ToString().Trim() != "XXX")
                    {
                        if (cboParameterName.Text.Trim().Length != 0 && cboParm1.Text.Trim().Length != 0 &&
                            cboParm2.Text.Trim().Length != 0 && cboParm3.Text.Trim().Length != 0 &&
                            cboParm4.Text.Trim().Length != 0 && cboParm5.Text.Trim().Length != 0 &&
                            cboParm6.Text.Trim().Length != 0 && cboParm7.Text.Trim().Length != 0)
                        {

                            intRow = grdConfigs.CurrentCell.RowIndex;
                            intColumn = grdConfigs.CurrentCell.ColumnIndex;

                            InputBoxResult intresult = InputBox.Show("Password: ");

                            if (intresult.ReturnCode == DialogResult.OK)
                            {
                                if (intresult.Text.Trim() == "Moses")
                                {

                                    General.updateConfigsRecord(Base.BaseConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType,
                                     cboParameterName.Text.Trim(), cboParm1.Text.Trim(), cboParm2.Text.Trim(), cboParm3.Text.Trim(), cboParm4.Text.Trim(),
                                     cboParm5.Text.Trim(), cboParm6.Text.Trim(), cboParm7.Text.Trim(), grdConfigs["ParameterName", intRow].Value.ToString().Trim(),
                                     grdConfigs["Parm1", intRow].Value.ToString().Trim(), grdConfigs["Parm2", intRow].Value.ToString().Trim(),
                                     grdConfigs["Parm3", intRow].Value.ToString().Trim(), grdConfigs["Parm4", intRow].Value.ToString().Trim());
                                    //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                                    foreach (string s in lstTableColumns)
                                    {
                                        if (dictGridValues[s] == grdConfigs[s, intRow].Value.ToString().Trim())
                                        {

                                        }
                                        else
                                        {
                                            //Write out to audit log
                                            writeAudit("CONFIGURATION", "U - Update", s, dictGridValues[s], grdConfigs[s, intRow].Value.ToString().Trim());

                                        }

                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Invalid password", "Error", MessageBoxButtons.OK);
                                }
                            }

                            grdConfigs["ParameterName", intRow].Value = cboParameterName.Text.Trim();
                            grdConfigs["ParameterName", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm1", intRow].Value = cboParm1.Text.Trim();
                            grdConfigs["Parm1", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm2", intRow].Value = cboParm2.Text.Trim();
                            grdConfigs["Parm2", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm3", intRow].Value = cboParm3.Text.Trim();
                            grdConfigs["Parm3", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm4", intRow].Value = cboParm4.Text.Trim();
                            grdConfigs["Parm4", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm5", intRow].Value = cboParm5.Text.Trim();
                            grdConfigs["Parm5", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm6", intRow].Value = cboParm6.Text.Trim();
                            grdConfigs["Parm6", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm7", intRow].Value = cboParm7.Text.Trim();
                            grdConfigs["Parm7", intRow].Style.BackColor = Color.Lavender;

                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data.", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabRates":
                    #region tabRates

                    //HJ
                    if (txtLowValue.Text.Trim().Length != 0 &&
                        txtHighValue.Text.Trim().Length != 0 && txtRate.Text.Trim().Length != 0)
                    {

                        InputBoxResult result = InputBox.Show("Password: ", "Rates Inputs are Password Protected!", "*", "0");

                        if (result.ReturnCode == DialogResult.OK)
                        {
                            if (result.Text.Trim() == "Moses")
                            {
                                intRow = grdRates.CurrentCell.RowIndex;
                                intColumn = grdRates.CurrentCell.ColumnIndex;

                                General.updateRatesRecord(Base.DBConnectionString, BusinessLanguage.BussUnit, 
                                                             txtMiningType.Text.Trim(),
                                                             txtBonusType.Text.Trim(),
                                                             txtPeriod.Text.ToString().Trim(), 
                                                             txtRateType.Text.Trim(), 
                                                             txtLowValue.Text.Trim(),
                                                             txtHighValue.Text.Trim(), txtRate.Text.Trim(),
                                                             grdRates["Low_Value", intRow].Value.ToString().Trim(), 
                                                             grdRates["High_Value", intRow].Value.ToString().Trim(),
                                                             grdRates["Rate", intRow].Value.ToString().Trim());
                                Application.DoEvents();

                                MessageBox.Show("All calculations will becleared.  Recalculations have to be done.", "Information", MessageBoxButtons.OK);
                                clearAllCalcValues("Ganglink", txtSelectedSection.Text.Trim());
                                clearAllCalcValues("Miners", txtSelectedSection.Text.Trim());
                                clearAllCalcValues("Bonusshifts", txtSelectedSection.Text.Trim());

                                grdRates["Low_Value", intRow].Value = txtLowValue.Text.Trim();
                                grdRates["Low_Value", intRow].Style.BackColor = Color.Lavender;
                                grdRates["High_Value", intRow].Value = txtHighValue.Text.Trim();
                                grdRates["High_Value", intRow].Style.BackColor = Color.Lavender;
                                grdRates["Rate", intRow].Value = txtRate.Text.Trim();
                                grdRates["Rate", intRow].Style.BackColor = Color.Lavender;
                                //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                                foreach (string s in lstTableColumns)
                                {
                                    if (dictGridValues[s] == grdRates[s, intRow].Value.ToString().Trim())
                                    {

                                    }
                                    else
                                    {
                                        //Write out to audit log
                                        writeAudit("RATES", "U - Update", s, dictGridValues[s], grdRates[s, intRow].Value.ToString().Trim());

                                    }

                                }

                            }
                            else
                            {
                                MessageBox.Show("Invalid Password.", "Information", MessageBoxButtons.OK);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabFactors":
                    #region tabFactors

                    //HJ
                    if (cboVarName.Text.Trim().Length != 0 && txtVarValue.Text.Trim().Length != 0)
                    {
                        intRow = grdFactors.CurrentCell.RowIndex;
                        intColumn = grdFactors.CurrentCell.ColumnIndex;

                        if (grdFactors[0, intRow].Value.ToString().Trim() != "XXX")
                        {

                            strSQL = "BEGIN transaction; Update Factors set VarName = '" + cboVarName.Text.Trim() +
                                             "', VarValue = '" + txtVarValue.Text.Trim() + "' Where " +
                                             " VarName = '" + grdFactors["VARNAME", intRow].Value.ToString().Trim() +
                                             "' and VarValue = '" + grdFactors["VARVALUE", intRow].Value.ToString().Trim() +
                                             "' and period = '" + BusinessLanguage.Period + "' ;Commit Transaction;";


                            grdFactors["VARNAME", intRow].Style.BackColor = Color.Lavender;
                            grdFactors["VARNAME", intRow].Value = cboVarName.Text.Trim();
                            grdFactors["VARVALUE", intRow].Style.BackColor = Color.Lavender;
                            grdFactors["VARVALUE", intRow].Value = txtVarValue.Text.Trim();

                            TB.InsertData(Base.DBConnectionString, strSQL);
                            //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                            foreach (string s in lstTableColumns)
                            {
                                if (dictGridValues[s] == grdFactors[s, intRow].Value.ToString().Trim())
                                {

                                }
                                else
                                {
                                    //Write out to audit log
                                    writeAudit("FACTORS", "U - Update", s, dictGridValues[s], grdFactors[s, intRow].Value.ToString().Trim());

                                }

                            }

                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion

                case "tabDrillers":

                    if (grdDrillers.Columns.Count > 2)
                    {
                        #region tabDrillers

                        if (cboAutoDrillerDrilInd.Text.Trim().Length > 0 &&
                            txtAutoDGang.Text.Trim().Length > 0 &&
                            txtAutoDWorkplace.Text.Trim().Length > 0 &&
                            txtAutoDrilShifts.Text.Trim().Length > 0)
                        {

                            if (Convert.ToInt32(cboAutoDrillerDrilInd.Text.ToString().Trim()) > 0)
                            {
                                intRow = grdDrillers.CurrentCell.RowIndex;

                                strSQL = "Update DRILLERS set DrillerShifts = '" + txtAutoDrilShifts.Text.Trim() +
                                         "' , DrillerInd = '" + cboAutoDrillerDrilInd.Text.Trim() +
                                         "' where GANG = '" + grdDrillers["Gang", intRow].Value.ToString().Trim() +
                                         "' and WORKPLACE = '" + grdDrillers["Workplace", intRow].Value.ToString().Trim() +
                                         "' and PERIOD = '" + grdDrillers["Period", intRow].Value.ToString().Trim() +      //xxxxxxxxxxxxx
                                         "' and EMPLOYEE_No = '" + grdDrillers["Employee_no", intRow].Value.ToString().Trim() + "'";

                                TB.InsertData(Base.DBConnectionString, strSQL);

                                grdDrillers["DrillerInd", intRow].Value = cboAutoDrillerDrilInd.Text.Trim();
                                grdDrillers["DrillerShifts", intRow].Value = txtAutoDrilShifts.Text.Trim();

                                for (int i = 0; i <= grdDrillers.Columns.Count - 1; i++)
                                {
                                    grdDrillers[i, intRow].Style.BackColor = Color.Lavender;
                                }
                            }
                            else
                            {
                                MessageBox.Show("Employee must have a driller indicator of 1 or 2", "Error", MessageBoxButtons.OK);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }

           
                        #endregion
                    }
                    else
                    {

                        MessageBox.Show("Please use the button 'Show Empl' to display the full detail of the selected driller.", "Information", MessageBoxButtons.OK);
                    }
                    break;
            }
        }

        private void clearAllCalcValues(string _Tablename, string _Section)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Update " + _Tablename + " set ");
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, _Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                sb.Append(row["CALC_NAME"].ToString() + " = '0',");
            }

            if (sb.Length > 25)
            {
                sb.Append(strWhere);

                string strTemp = Convert.ToString(sb.Replace(",where", " Where"));
                TB.InsertData(Base.DBConnectionString, strTemp);
            }
        }

        private void btnProcessAll_Click(object sender, EventArgs e)
        {

            int intCheckLocks = checkLockInputProcesses();
            if (intCheckLocks == 0)
            {
                openTab(tabProcess);

                //checkProcess();
                calcCrewsandGangs();
            }
            else
            {
                MessageBox.Show("Finish all input processes first, before trying to process all.", "Informations", MessageBoxButtons.OK);
            }

        }

        private void deleteAllColumns(string Tablename)
        {
            //xxxxxxxxxxxxxxxxxxx
            //Create the earnings table
            createTheFile(Tablename);

            //Add the calculation columns.
            createEarningsColumns(Tablename);

            List<string> lstColumnNames = new List<string>();

            //extract the latest data from the base file e.g. Ganglink, Bonusshifts and replace data in the earningsfile.

            DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename,
                           " where section = '" + txtSelectedSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period + "'");

            //Give the tempory file a name
            tb.TableName = Tablename + "EARN" + BusinessLanguage.Period.Trim();

            if (Tablename.ToUpper() == "BONUSSHIFTS")
            {
                #region Remove columns starting with DAY from BONUSSHIFTS
                //Remove all the columns starting with "day" from temporary file, because BONUSSHIFTSEARN does not carry the DAY columns
                foreach (DataColumn dc in tb.Columns)
                {
                    if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                    {
                        lstColumnNames.Add(dc.ColumnName.Trim());
                    }
                    else
                    {

                    }
                }

                foreach (string s in lstColumnNames)
                {
                    tb.Columns.Remove(s);
                    tb.AcceptChanges();
                }

                lstColumnNames.Clear();
                #endregion
            }

            //Save the data to be processed to the earnings table.
            TB.saveCalculations2(tb, Base.DBConnectionString, " where section = '" + txtSelectedSection.Text.Trim() + "'",
                                 tb.TableName.Trim());

            Application.DoEvents();
            //}
        }

        private void createTheFile(string Tablename)
        {
            //Check if earningstable exist - e.g. GangLinkEarn201108....if not...CREATE the table
            List<string> lstColumnNames = new List<string>();

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, Tablename + "EARN" + BusinessLanguage.Period.Trim());

            if (intCount > 0)
            {
            }
            else
            {
                //CREATE the earnings table:  GanglinkEarn201108
                //Extract the table into a temp file from the datafile e.g. GANGLINK, BONUSSHIFTS, DRILLERS etc.

                DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename,
                               "where section = '" + txtSelectedSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period + "'");

                //Give the tempory file a name
                tb.TableName = Tablename + "Earn" + BusinessLanguage.Period.Trim();

                if (Tablename.ToUpper() == "BONUSSHIFTS")
                {
                    #region Remove columns starting with DAY from BONUSSHIFTS
                    //Remove all the columns starting with "day" from temporary file, because BONUSSHIFTSEARN does not carry the DAY columns
                    foreach (DataColumn dc in tb.Columns)
                    {
                        if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                        {
                            lstColumnNames.Add(dc.ColumnName.Trim());
                        }
                        else
                        {

                        }
                    }

                    foreach (string s in lstColumnNames)
                    {
                        tb.Columns.Remove(s);
                        tb.AcceptChanges();
                    }

                    lstColumnNames.Clear();
                    #endregion
                }

                strSqlAlter.Remove(0, strSqlAlter.Length);

                //First create the base table.  Why, because all these columns should be NOT NULL.  
                //The Formulas SHOULD be NULL when created
                foreach (DataColumn dc in tb.Columns)
                {
                    if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                    {
                    }
                    else
                    {
                        lstColumnNames.Add(dc.ColumnName);
                    }
                }

                //Create the earningstable e.g. BONUSSHIFTSEARN201108T

                TB.createEarningsTable(Base.DBConnectionString, tb.TableName, Tablename, lstColumnNames);

            }
        }

        private void createEarningsColumns(string Tablename)
        {
            DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename + "EARN" + BusinessLanguage.Period);

            strSqlAlter.Remove(0, strSqlAlter.Length);
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(Base.DBName + BusinessLanguage.Period,
                                      Tablename + "EARN", Base.AnalysisConnectionString);

            foreach (DataRow row in tableformulas.Rows)
            {
                if (tb.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                }
                else
                {
                    strSqlAlter = strSqlAlter.Append(" ; Alter table " + Tablename + "EARN" + BusinessLanguage.Period + " add " +
                                                     row["CALC_NAME"].ToString().Trim() + " varchar(50) NULL");
                }
            }

            if (strSqlAlter.ToString().Trim().Length > 0)
            {
                StringBuilder bld = new StringBuilder();
                bld.Append("BEGIN transaction;" + strSqlAlter.ToString().Substring(1).Trim() + ";COMMIT transaction;");
                TB.InsertData(Base.DBConnectionString, bld.ToString().Trim());
                Application.DoEvents();
            }
            else
            {
            }
        }



        private void deleteAllCalcColumns(string Tablename)
        {
            strSqlAlter.Remove(0, strSqlAlter.Length);
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                TB.removeColumn(Base.DBConnectionString, Tablename, row["CALC_NAME"].ToString());
            }
        }

        private void deleteAllCalcColumns(string Tablename, DataTable Table)
        {
            //remove the column from the database.
            strSqlAlter.Remove(0, strSqlAlter.Length);

            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                if (Table.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                    TB.removeColumn(Base.DBConnectionString, Tablename, row["CALC_NAME"].ToString());
                }

            }
        }

        private void deleteAllCalcColumnsFromTempTable(string Tablename, DataTable Table)
        {
            //remove the column from the database.
            strSqlAlter.Remove(0, strSqlAlter.Length);

            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                if (Table.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                    Table.Columns.Remove(row["CALC_NAME"].ToString().Trim());
                }
            }

            Table.AcceptChanges();
        }

        private void Calcs(string tablename, string phasename, string Delete)
        {
            if (Delete == "Y")
            {
                deleteAllColumns(tablename);
            }

            TB.insertProcess(Base.AnalysisConnectionString, Base.DBName + BusinessLanguage.Period, tablename + "EARN", phasename, txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "N", "N", (string)DateTime.Now.ToLongTimeString(), Convert.ToString(++intProcessCounter));

        }

        private void openTab(TabPage tp)
        {
            this.tabInfo.SelectedTab = tp;

            Application.DoEvents();

        }

        private void calcCrewsandGangs()
        {
            string strTableName = "";

            for (int i = 1; i <= 4; i++)
            {
                strTableName = "GangLink" + Convert.ToString(i).Trim();
                switch (i)
                {
                    case 1:
                        //btnPhase1.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 1", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Application.DoEvents();
                        Calcs("GangLink", "Ganglink10", "Y");
                        break;

                    case 2:
                        //btnPhase1.BackColor = Color.LightGreen;
                        //btnPhase2.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 2", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Application.DoEvents();
                        Calcs("GangLink", "Ganglink20", "Y");
                        break;

                    case 3:
                        //btnPhase2.BackColor = Color.LightGreen;
                        //btnPhase3.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 3", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Application.DoEvents();
                        Calcs("GangLink", "Ganglink30", "Y");
                        break;

                    case 4:
                        //btnPhase3.BackColor = Color.LightGreen;
                        //btnPhase4.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 4", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Header", "Base Calc Process", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());

                        //Application.DoEvents();
                        Calcs("GangLink", "Ganglink40", "Y");
                        break;
                }

                //executeFormulas(strTableName);
            }


            //btnPhase4.BackColor = Color.LightGreen;
            //Application.DoEvents();

        }

        private void calcCrewsandGangs(int counter)
        {
            string strTableName = "";

            for (int i = counter; i <= counter; i++)
            {
                strTableName = "GangLink" + Convert.ToString(i).Trim();
                switch (i)
                {
                    case 1:
                        Calcs("GangLink", "Ganglink10", "Y");
                        break;

                    case 2:
                        Calcs("GangLink", "Ganglink20", "N");
                        break;

                    case 3:
                        Calcs("GangLink", "Ganglink30", "N");
                        break;

                    case 4:
                        Calcs("GangLink", "Ganglink40", "N");
                        break;
                }


            }



            Application.DoEvents();

        }

        private void executeCostSheetFormulas(string TableName)
        {

            string strSQL = "BEGIN transaction; Delete from monitor ; commit transaction;";
            TB.InsertData(Base.DBConnectionString, strSQL);
            string strprevPeriod = TableName;
            strSQL = "BEGIN transaction; insert into monitor values('" + Base.DBName + "','" + strprevPeriod + "','N','0','" + txtSelectedSection.Text.Trim() + "','0','0'); commit transaction; ";
            TB.InsertData(Base.DBConnectionString, strSQL);

        }

        #region Open Tabs

        private void btnLockCalendar_Click(object sender, EventArgs e)
        {
            openTab(tabCalendar);
        }

        private void btnLockSurvey_Click(object sender, EventArgs e)
        {
            openTab(tabSurvey);
        }

        private void btnLockDrillers_Click(object sender, EventArgs e)
        {
            openTab(tabDrillers);
        }

        private void btnLockBonusShifts_Click(object sender, EventArgs e)
        {
            openTab(tabLabour);
        }
 
        private void btnLockGangLink_Click(object sender, EventArgs e)
        {
            openTab(tabGangLinking);
        }

        private void btnLockMiners_Click(object sender, EventArgs e)
        {
            openTab(tabMiners);
        }

        private void btnLockOffday_Click(object sender, EventArgs e)
        {
            openTab(tabOffday);
        }

        private void btnLockEmplPen_Click(object sender, EventArgs e)
        {
            openTab(tabEmplPen);
        }
        #endregion

        private void btnCrewLevel_Click(object sender, EventArgs e)
        {
            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {
                calcCrewsandGangs();

                evaluateStatus();
            }
            else
            {
                MessageBox.Show("Finish all input processes first, before trying to process all.", "Informations", MessageBoxButtons.OK);
            }
        }

        private void btnEmplTeamCalcHeader_Click(object sender, EventArgs e)
        {
            evaluateStatus();
        }

        private void saveXXXTeamShifts(DataTable TeamShifts)
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            #region TeamPrint
            foreach (DataRow rr in TeamShifts.Rows)
            {

                strSQL.Append("insert into TeamShifts values('" + rr["SECTION"].ToString().Trim() +
                              "','" + rr["CONTRACT"].ToString().Trim() + "','" + rr["WORKPLACE"].ToString().Trim() + "','" +
                              rr["GANG"].ToString().Trim() + "','" + rr["WAGECODE"].ToString().Trim() + "','" + rr["LINERESPCODE"].ToString().Trim() + "','" +
                              rr["EMPLOYEE_NO"].ToString().Trim() + "','" + rr["INITIALS"].ToString().Trim() + "','" +
                              rr["SURNAME"].ToString().Trim() + "','" + rr["REGISTER"].ToString().Trim() + "','" +
                              rr["DATEFROM"].ToString().Trim() + "','" + rr["EMPLOYEEPRODUCTIONBONUS"].ToString().Trim() + "','" +
                              rr["EMPLOYEEDRESSINGBONUS"].ToString().Trim() + "','" + rr["EMPLOYEEAWOPPENALTYBONUS"].ToString().Trim() + "','" +
                              rr["EMPLOYEEAWOPDRESSNGPENALTYBONUS"].ToString().Trim() + "','" + rr["EMPLOYEEHYDROBONUS"].ToString().Trim() + "','" +
                              rr["EMPLOYEESTOPEPROCESSBONUS"].ToString().Trim() + "')");



            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            #endregion

        }

        private void saveXXXTeamProd(DataTable Teamprod)
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            #region TeamPrint
            foreach (DataRow rr in Teamprod.Rows)
            {
                //"CREATE TABLE TEAMPROD (SECTION char(50), CONTRACT Char(50), WORKPLACE Char(50), " +
                //    "GANG Char(50),WPNAME Char(50),WPSHIFTS Char(50),WPSHIFTSTOTAL Char(50), WPSQM Char(50), " +
                //    "WPFOOTWALL Char(50),WPSTOPEWIDTH Char(50),WPSTOPEWIDTHRATE Char(50), WPSTOPEWIDTHBONUS Char(50), " +
                //    "WPCONTRACTORBONUS Char(50),WPTOTALBONUS Char(50))";

                strSQL.Append("insert into TeamProd values('" + rr["SECTION"].ToString().Trim() + "','" + rr["CONTRACT"].ToString().Trim() +
                              "','" + rr["WORKPLACE"].ToString().Trim() + "','" +
                              rr["GANG"].ToString().Trim() + "','" + rr["CREWNO"].ToString().Trim() + "','" + rr["WPNAME"].ToString().Trim() + "','" + rr["WPSHIFTS"].ToString().Trim() +
                              "','" + rr["WPSHIFTSTOTAL"].ToString().Trim() + "','" + rr["WPSQM"].ToString().Trim() + "','" +
                              rr["WPFOOTWALL"].ToString().Trim() + "','" + rr["WPSTOPEWIDTH"].ToString().Trim() +
                              "','" + rr["WPSTOPEWIDTHRATE"].ToString().Trim() + "','" + rr["WPSTOPEWIDTHBONUS"].ToString().Trim() +
                              "','" + rr["WPCONTRACTORBONUS"].ToString().Trim() + "','" + rr["WPTOTALBONUS"].ToString().Trim() + "');");
            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            #endregion

        }

        private void WriteCSV(DataTable dt)
        {
            StreamWriter sw;
            string filePath = strServerPath + ":\\Crew.csv";


            sw = File.CreateText(filePath);

            try
            {
                // write the data in each row & column
                int intcounter = 0;
                foreach (DataRow row in dt.Rows)
                {
                    // recreate an empty Stringbuilder through each row iteration.
                    StringBuilder rowToWrite = new StringBuilder();

                    for (int counter = 0; counter <= dt.Columns.Count - 1; counter++)
                    {
                        if (intcounter == 0)
                        {
                            foreach (DataColumn column in dt.Columns)
                            {
                                //rowToWrite.Append("'" + column.ColumnName + "'");
                                rowToWrite.Append("'" + column.ColumnName + "'");
                            }
                            rowToWrite.Replace("''", "','");
                            rowToWrite.Replace("'", "");

                            rowToWrite.Append("\r\n");
                            sw.Write(rowToWrite);
                            rowToWrite.Remove(0, rowToWrite.Length);
                        }
                        intcounter = intcounter + 1;
                        rowToWrite.Append("'" + row[counter] + "'");
                    }

                    rowToWrite.Replace("''", "','");
                    rowToWrite.Replace("'", "");

                    rowToWrite.Append("\r\n");
                    sw.Write(rowToWrite);
                }
            }
            catch
            {
                //("An error occurred while attempting to build the CSV file. " + e.Message);
            }
            finally
            {
                sw.Close();
            }
        }

        private void btnApplyOffdays_Click(object sender, EventArgs e)
        {
            if (txtSelectedSection.Text.Trim() == "***")
            {
                MessageBox.Show("Please select a section.", "Information", MessageBoxButtons.OK);
            }
            else
            {
                UpdateClockedShifts();
            }
        }

        private void btnDeleteRow_Click_1(object sender, EventArgs e)
        {

            int intRow = 0;
            int intColumn = 0;

            string strSQL = "";

            switch (tabInfo.SelectedTab.Name)
            {
                case "tabMiners":
                    #region tabminers
                    intRow = grdMiners.CurrentCell.RowIndex;
                    intColumn = grdMiners.CurrentCell.ColumnIndex;
                    string strName = "";
                    string strDesignation = "";

                    if (cboNames.Text.Contains("-"))
                    {
                        strName = cboNames.Text.Substring(0, cboNames.Text.IndexOf("-")).Trim();
                    }
                    else
                    {
                        strName = cboNames.Text.Trim();
                    }

                    if (cboDesignation.Text.Contains("-"))
                    {
                        strDesignation = cboDesignation.Text.Substring(0, cboDesignation.Text.IndexOf("-")).Trim();
                    }
                    else
                    {
                        strDesignation = cboDesignation.Text.Trim();
                    }

                    if (grdMiners["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXX")
                    {

                        strSQL = "BEGIN transaction; Delete from Miners " +
                                 " Where Section = '" + grdMiners["Section", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdMiners["Period", intRow].Value.ToString().Trim() +
                                 "' and GANG = '" + grdMiners["GANG", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + strName.Trim() +
                                 "' and Designation = '" + strDesignation.Trim() +
                                 "' and safetyind = '" + grdMiners["SafetyInd", intRow].Value.ToString().Trim() +
                                 "' and Payshifts = '" + grdMiners["Payshifts", intRow].Value.ToString().Trim() +
                                 "' and Awop_Shifts = '" + grdMiners["Awop_Shifts", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateMiners();

                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion

                case "tabGangLinking":

                    #region tabGangLink
                    intRow = grdGangLink.CurrentCell.RowIndex;
                    intColumn = grdGangLink.CurrentCell.ColumnIndex;

                    DialogResult result = MessageBox.Show("Sure you would like to delete gang: " + grdGangLink["Gang", intRow].Value.ToString().Trim() +
                        "' on workplace: '" + grdGangLink["Workplace", intRow].Value.ToString().Trim() + "?", "Question", MessageBoxButtons.YesNo);

                    switch (result)
                    {

                        case DialogResult.Yes:
                            strSQL = "BEGIN transaction; Delete from Ganglink " +
                                 " Where Section = '" + grdGangLink["Section", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdGangLink["Period", intRow].Value.ToString().Trim() +
                                 "' and Gang = '" + grdGangLink["Gang", intRow].Value.ToString().Trim() +
                                 "' and Workplace = '" + grdGangLink["Workplace", intRow].Value.ToString().Trim() +
                                 "' and SafetyInd = '" + grdGangLink["SafetyInd", intRow].Value.ToString().Trim() +
                                 "' and GangType = '" + grdGangLink["GangType", intRow].Value.ToString().Trim() +
                                 "';Commit Transaction;";

                            TB.InsertData(Base.DBConnectionString, strSQL);
                            evaluateGangLinking();
                            break;


                        case DialogResult.No:
                            break;


                    }

                    break;

                    #endregion
                    break;

                case "tabAbnormal":
                    #region tabAbnormal
                    intRow = grdAbnormal.CurrentCell.RowIndex;
                    intColumn = grdAbnormal.CurrentCell.ColumnIndex;
                    rowindex = grdAbnormal.CurrentCell.RowIndex;


                    if (grdAbnormal["ABNORMALVALUE", intRow].Value.ToString().Trim() != "XXX")
                    {

                        strSQL = "BEGIN transaction; Delete from Abnormal " +
                                     " Where Section = '" + grdAbnormal["Section", intRow].Value.ToString().Trim() +
                                     "' and Period = '" + grdAbnormal["Period", intRow].Value.ToString().Trim() +
                                     "' and Workplace = '" + grdAbnormal["Workplace", intRow].Value.ToString().Trim() +
                                     "' and AbnormalLevel = '" + grdAbnormal["AbnormalLevel", intRow].Value.ToString().Trim() +
                                     "' and AbnormalType = '" + grdAbnormal["AbnormalType", intRow].Value.ToString().Trim() +
                                     "' and AbnormalValue = '" + grdAbnormal["AbnormalValue", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateAbnormal();
                        grdAbnormal.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion


                case "tabOffday":
                    #region tabOffdays

                    if (cboOffDaysGang.Text.Trim().Length > 0 &&
                        cboOffDaysSection.Text.Trim().Length > 0)
                    {
                        strSQL = "delete from Offdays where gang = '" + grdOffDays["Gang", intRow].Value +
                                 "' and section = '" + grdOffDays["Section", intRow].Value +
                                 "' and OffDayValue = '" + grdOffDays["OffdayValue", intRow].Value + "'";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateOffDays();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion


                case "tabEmplPen":
                    #region tabEmployeePenalty

                    intRow = grdEmplPen.CurrentCell.RowIndex;
                    intColumn = grdEmplPen.CurrentCell.ColumnIndex;

                    if (grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXX")
                    {

                        strSQL = "BEGIN transaction; Delete from EmployeePenalties " +
                                 " Where Section = '" + grdEmplPen["Section", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdEmplPen["Period", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + grdEmplPen["Employee_no", intRow].Value.ToString().Trim() +
                                 "' and Workplace = '" + grdEmplPen["Workplace", intRow].Value.ToString().Trim() +
                                 "' and PenaltyInd = '" + grdEmplPen["PenaltyInd", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateEmployeePenalties();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion

                case "tabDrillers":
                    #region tabDrillers

                    if (cboAutoDrillerDrilInd.Text.Trim().Length > 0 &&
                        txtAutoDGang.Text.Trim().Length > 0 &&
                        txtAutoDWorkplace.Text.Trim().Length > 0 &&
                        txtAutoDrilShifts.Text.Trim().Length > 0)
                    {
                        intRow = grdDrillers.CurrentCell.RowIndex;


                        strSQL = "delete from DRILLERS where DrillerShifts = '" + txtAutoDrilShifts.Text.Trim() +
                                 "' and DrillerInd = '" + cboAutoDrillerDrilInd.Text.Trim() +
                                 "' and GANG = '" + grdDrillers["Gang", intRow].Value +
                                 "' and WORKPLACE = '" + grdDrillers["Workplace", intRow].Value +
                                 "' and EMPLOYEE_No = '" + grdDrillers["Employee_no", intRow].Value + "'";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateDrillers();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion


            }
        }

        protected virtual void FrontDecorator(System.Web.UI.HtmlTextWriter writer)
        {
            writer.WriteFullBeginTag("HTML");
            writer.WriteFullBeginTag("Head");
            writer.RenderBeginTag(System.Web.UI.HtmlTextWriterTag.Style);
            writer.Write("<!--");

            StreamReader sr = File.OpenText(strServerPath + ":\\koos.html");
            String input;
            while ((input = sr.ReadLine()) != null)
            {
                writer.WriteLine(input);
            }
            sr.Close();
            writer.Write("-->");
            writer.RenderEndTag();
            writer.WriteEndTag("Head");
            writer.WriteFullBeginTag("Body");
        }

        protected virtual void RearDecorator(System.Web.UI.HtmlTextWriter writer)
        {
            writer.WriteEndTag("Body");
            writer.WriteEndTag("HTML");
        }

        private void printHTML(DataTable dt, string TabName)
        {
            if (dt.Columns.Count > 0)
            {
                string OPath = "c:\\koos.html";
                try
                {

                    StreamWriter SW = new StreamWriter(OPath);
                    //StringWriter SW = new StringWriter();
                    System.Web.UI.HtmlTextWriter HTMLWriter = new System.Web.UI.HtmlTextWriter(SW);
                    System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();

                    grid.DataSource = dt;
                    grid.DataBind();

                    using (SW)
                    {
                        using (HTMLWriter)
                        {

                            HTMLWriter.WriteLine("HARMONY - Tshepong Mine - " + TabName);
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteLine("==============================");
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteBreak();

                            grid.RenderControl(HTMLWriter);
                            //RearDecorator(HTMLWriter);

                        }
                    }

                    SW.Close();
                    HTMLWriter.Close();


                    System.Diagnostics.Process P = new System.Diagnostics.Process();
                    P.StartInfo.WorkingDirectory = strServerPath + ":\\Program Files\\Internet Explorer";
                    P.StartInfo.FileName = "IExplore.exe";
                    P.StartInfo.Arguments = "C:\\koos.html";
                    P.Start();
                    P.WaitForExit();


                }
                catch (Exception exx)
                {
                    MessageBox.Show("Could not create " + OPath.Trim() + ".  Create the directory first." + exx.Message, "Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Your spreadsheet could not be created.  No columns found in datatable.", "Error Message", MessageBoxButtons.OK);
            }

        }

        private void btnLoad_Click_1(object sender, EventArgs e)
        {
            if (listBox3.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select the number of measuring shifts", "Information", MessageBoxButtons.OK);
            }
            else
            {
                if (txtSelectedSection.Text.Trim().Length == 0)
                {
                    MessageBox.Show("Please select a section and the correct month measuring shifts for the section.", "Information", MessageBoxButtons.OK);
                }
                else
                {
                    string selectedSection = txtSelectedSection.Text.Trim();
                    string grdSection = grdCalendar["SECTION", intFiller].Value.ToString().Trim();
                    if (selectedSection == grdSection)
                    {
                        Base.updateCalendarRecord(Base.DBConnectionString, BusinessLanguage.BussUnit, txtMiningType.Text.Trim(),
                                                         txtBonusType.Text.Trim(), txtSelectedSection.Text.Trim(),
                                                         txtPeriod.Text.ToString().Trim(),
                                                         (Convert.ToDateTime(dateTimePicker1.Text)).ToString("yyyy-MM-dd"),
                                                         (Convert.ToDateTime(dateTimePicker2.Text)).ToString("yyyy-MM-dd"),
                                                         listBox3.SelectedItem.ToString().Trim());
                        Application.DoEvents();
                    }

                    else
                    {
                        MessageBox.Show("Selected section not the same as grid section.", "Informations", MessageBoxButtons.OK);
                    }

                    //Extract Calendar again and insert into 
                    Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar", " where period = '" + BusinessLanguage.Period + "'");
                    grdCalendar.DataSource = Calendar;
                }
            }
        }

        private void btnCostsheetPhase1_Click(object sender, EventArgs e)
        {
             
        }

        private void CalcCostSheetEmployees()
        {
            
             
        }

        private void label83_Click(object sender, EventArgs e)
        {
            extractDBTableNames(listBox1);
        }

        private void btnLockPaysend_Click(object sender, EventArgs e)
        {
            if (Base.DBTables.Contains("PAYROLL"))
            {
            }
            else
            {
                if (myConn.State == ConnectionState.Open)
                {
                }
                else
                {
                    myConn.Open();
                }

                //Create a table
                Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");
                if (intCount > 0)
                {
                }
                else
                {
                    TB.createPayrollTable(Base.DBConnectionString);
                }
            }

            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, txtSelectedSection.Text.Trim());
            paysend.Show();


        }

        private void btnEmployeeCostsheet_Click(object sender, EventArgs e)
        {
            Calcs("Miners", "Miners", "N");
        }

        private void btnPrint_Click_1(object sender, EventArgs e)
        {
            switch (tabInfo.SelectedTab.Name)
            {
                case "tabMiners":
                    #region tabMiners

                    DataTable dt = Base.extractPrintData(Base.DBConnectionString, "Miners", strWhere);
                    deleteAllCalcColumns("Miners", dt);
                    if (dt.Rows.Count > 0)
                    {
                        dt.Columns.Remove("BUSSUNIT");
                        dt.Columns.Remove("MININGTYPE");
                        dt.Columns.Remove("BONUSTYPE");
                        dt.Columns.Remove("SAFETYIND");
                        dt.Columns.Remove("SHIFTS_WORKED");
                        dt.AcceptChanges();

                        printHTML(dt, "Miners");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabGangLinking":
                    #region tabGangLinking

                    dt = Base.extractPrintData(Base.DBConnectionString, "GangLink", strWhere);
                    deleteAllCalcColumns("GangLink", dt);
                    dt.AcceptChanges();
                    if (dt.Rows.Count > 0)
                    {

                        printHTML(dt, "GangLinking");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabAbnormal":
                    #region tabAbnormal

                    dt = Base.extractPrintData(Base.DBConnectionString, "Abnormal", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Abnormal");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabLabour":
                    #region tabLabour

                    dt = Base.extractPrintData(Base.DBConnectionString, "BonusShifts", strWhere);
                    deleteAllCalcColumns("BonusShifts", dt);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "BonusShifts");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabSurvey":
                    #region tabSurvey

                    dt = Base.extractPrintData(Base.DBConnectionString, "Survey", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Survey");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties

                    dt = Base.extractPrintData(Base.DBConnectionString, "EmployeePenalties", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "EmployeePenalties");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }
                    break;
                    #endregion

                case "tabOffday":
                    #region tabOffdays

                    dt = Base.extractPrintData(Base.DBConnectionString, "Offdays", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Offdays");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabCalendar":
                    #region tabCalendar

                    dt = Base.extractPrintData(Base.DBConnectionString, "Calendar", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Calendar");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabClockShifts":
                    #region tabClockShifts

                    dt = Base.extractPrintData(Base.DBConnectionString, "ClockedShifts", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "ClockedShifts");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabRates":
                    #region tabRates

                    dt = Base.extractPrintData(Base.DBConnectionString, "Rates", "");
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Rates");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabMonitor":
                    #region tabRates

                    dt = Base.extractPrintData(Base.DBConnectionString, "Monitor", "");
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Monitor");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

            }
        }

        private void calcStopeData()
        {
            Base.Period = txtPeriod.Text.Trim();
            //Base.Period = "200909";

            SqlConnection stopeConn = Base.StopeConnection;
            stopeConn.Open();

            try
            {
                DataTable ContractTotals = TB.getContractCrewOfficialBonus(Base.StopeConnectionString, "STOPING", txtSelectedSection.Text.Trim());

                stopeConn.Close();

                TB.updateDSShiftbossCrewBonus(Base.DBConnectionString, ContractTotals);
            }
            catch { }


        }

        private void btnBaseCalcsHeader_Click(object sender, EventArgs e)
        {
            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {
                //Check if the a calculator is currently running
                Int16 intCount1 = TB.checkTableExist(Base.DBConnectionString, "BonusShiftsEARN");
                Int16 intCount2 = TB.checkTableExist(Base.DBConnectionString, "GanglinkEARN");
                Int16 intCount3 = TB.checkTableExist(Base.DBConnectionString, "SupportLinkEARN");
                Int16 intCount4 = TB.checkTableExist(Base.DBConnectionString, "DrillersEARN");
                Int16 intCount5 = TB.checkTableExist(Base.DBConnectionString, "MinersEARN");
                Int16 intCount6 = TB.checkTableExist(Base.DBConnectionString, "SectionEarningsEARN");

                if (intCount1 > 0 || intCount2 > 0 || intCount3 > 0 || intCount4 > 0 || intCount5 > 0 || intCount6 > 0)
                {
                    MessageBox.Show("A calculator is currently running for this bonus scheme: " + BusinessLanguage.MiningType +
                                    " " + BusinessLanguage.BonusType);
                }
                else
                {
                    startCalcProcess();

                }

            }
            else
            {
                MessageBox.Show("Finish all input processes first, before trying to process all.", "Informations", MessageBoxButtons.OK);
            }
           

        }

        private void startCalcProcess()
        {
            this.Cursor = Cursors.WaitCursor;
            btnx.Visible = true;
            btnx.Enabled = true;
            btnx.Text = "Run";
            TB.deleteProcess(Base.AnalysisConnectionString, Base.DBName + BusinessLanguage.Period);
            //clear the monitor table
            TB.deleteAllExcept(Base.DBConnectionString, "Monitor");
            Calcs("BonusShifts", "BonusShiftsearn09", "Y");
            Calcs("GangLink", "Ganglinkearn05", "Y");
            Calcs("GangLink", "Ganglinkearn10", "N");
            Calcs("GangLink", "Ganglinkearn20", "N");
            Calcs("GangLink", "Ganglinkearn30", "N");
            Calcs("GangLink", "Ganglinkearn40", "N");
            Calcs("GangLink", "Ganglinkearn50", "N");
            Calcs("GangLink", "Ganglinkearn60", "N");
            Calcs("GangLink", "Ganglinkearn70", "N");
            Calcs("Drillers", "Drillersearn10", "Y");
            Calcs("Drillers", "Drillersearn20", "N");
            Calcs("Drillers", "Drillersearn30", "N");
            Calcs("BonusShifts", "BonusShiftsearn10", "N");
            Calcs("SectionEarnings", "SectionEarningsearn10", "Y");
            Calcs("SectionEarnings", "SectionEarningsearn20", "N");
            Calcs("BonusShifts", "BonusShiftsearn20", "N");
            Calcs("BonusShifts", "BonusShiftsearn30", "N");
            Calcs("BonusShifts", "BonusShiftsearn40", "N");
            Calcs("BonusShifts", "BonusShiftsearn50", "N");
            Calcs("BonusShifts", "BonusShiftsearn60", "N");
            Calcs("BonusShifts", "BonusShiftsearn65", "N");
            Calcs("BonusShifts", "BonusShiftsearn70", "N");
            Calcs("Miners", "Minersearn1", "Y");
            Calcs("Miners", "Minersearn2", "N");
            Calcs("Miners", "Minersearn3", "N");
            Calcs("Miners", "Minersearn10", "N");
            Calcs("Exit", "Exit", "N");
            btnBaseCalcs.BackColor = Color.Orange;
            btnGangLinkCalcs.BackColor = Color.Orange;
            btnMinersCalc.BackColor = Color.Orange;
            btnBonusShiftsCalcs.BackColor = Color.Orange;

            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "Ganglinkearn10", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");;
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "Ganglinkearn70", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "BonusShiftsearn60", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "Minersearn10", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "Exit", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");

            Base.backupDatabase3(Base.DBConnectionString, Base.DBName, Base.BackupPath);
            this.Cursor = Cursors.Arrow;
        }

        private void btnBaseCalcs_Click(object sender, EventArgs e)
        {
           

          
        }

        private void btnGangLinkCalcs_Click(object sender, EventArgs e)
        {
             
        }

        private void grdActiveSheet_ColumnHeaderMouseClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int mouseY = MousePosition.Y;
                int mouseX = MousePosition.X;

                ctMenu.Show(this, new Point(mouseX, mouseY));

                columnnr = e.ColumnIndex;
                //string columname = grdActiveSheet.Columns[columnnr].Name;
                //DialogResult result = MessageBox.Show("Do you want to delete the column:  " + grdActiveSheet.Columns[columnnr].HeaderText + "?", "INFORMATION", MessageBoxButtons.YesNo);

                //if (result == DialogResult.Yes)
                //{
                //    //columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
                //    //TB.removeColumn(Base.DBConnectionString, TB.TBName, grdActiveSheet.Columns[columnnr].HeaderText);
                //    //DoDataExtract();
                //    //grdActiveSheet.DataSource = TB.getDataTable(TB.TBName);
                //    grdActiveSheet.Columns[columnnr].Visible = false;
                //}
                //else
                //{
                //    if (listBox1.SelectedItem.ToString().Trim() == "MONITOR")
                //    {

                //        string strSQL = "Begin transaction; Delete from monitor; commit transaction";
                //        TB.InsertData(Base.DBConnectionString, strSQL);
                //        Application.DoEvents();

                //    }
                //}
            }

            else
            {
                AConn = Analysis.AnalysisConnection;
                AConn.Open();
                DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, TB.TBName, Base.AnalysisConnectionString);

                foreach (DataRow dt in tempDataTable.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }



            }
        }

        private void grdCalendar_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dateTimePicker1.Value = Convert.ToDateTime(Calendar.Rows[e.RowIndex]["FSH"].ToString().Trim());
                dateTimePicker2.Value = Convert.ToDateTime(Calendar.Rows[e.RowIndex]["LSH"].ToString().Trim());
                intFiller = e.RowIndex;
            }

        }

        //private void cboAbnormalLevel_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    cboAbnormalType.Items.Clear();

        //    foreach (DataRow dr in Configs.Rows)
        //    {
        //        if (dr["PARAMETERNAME"].ToString().Trim() == "ABNORMAL" && dr["Miningtype"].ToString().Trim() == BusinessLanguage.MiningType
        //           && dr["BONUSTYPE"].ToString().Trim() == txtBonusType.Text.Trim() && dr["Parm1"].ToString().Trim() == cboAbnormalLevel.SelectedItem.ToString().Trim())
        //        {
        //            for (int i = 5; i <= Configs.Columns.Count - 1; i++)
        //            {
        //                if (dr[i].ToString().Trim() == "Q")
        //                {
        //                }
        //                else
        //                {
        //                    cboAbnormalType.Items.Add(dr[i].ToString().Trim()); 
        //                }
        //            }
        //        }
        //    }
        //}

        private void grdRates_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {

            }
            else
            {
                if (grdRates["RATE_TYPE", e.RowIndex].Value.ToString().Trim() == "XXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                    btnInsertRow.Enabled = true;

                }
                else
                {
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                    btnInsertRow.Enabled = true;
                }

                txtRateType.Text = grdRates["RATE_TYPE", e.RowIndex].Value.ToString().Trim();
                txtLowValue.Text = grdRates["LOW_VALUE", e.RowIndex].Value.ToString().Trim();
                txtHighValue.Text = grdRates["HIGH_VALUE", e.RowIndex].Value.ToString().Trim();
                txtRate.Text = grdRates["RATE", e.RowIndex].Value.ToString().Trim();
            }
        }

        private void tBViewColumnsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //if (listBox1.SelectedItems.Count == 0)
            //{
            //    MessageBox.Show("Please select the table to be processed","Confirm",MessageBoxButtons.OK);
            //}
            //else
            //{
            //    scrHideColumns hidecolumns = new scrHideColumns();
            //    hidecolumns.scrscrHideColumnsLoad(listBox1.SelectedItem.ToString().Trim(),Base.DBConnectionString.ConnectionString);
            //    hidecolumns.Show();
            //}
        }

        private void btnBonusShiftsCalcs_Click(object sender, EventArgs e)
        {
             
            
        }

        private void payrollSend_Click(object sender, EventArgs e)
        {
            //scrPayroll scrPay = new scrPayroll();
            //scrPay.Show();

            //if (Base.DBName.Trim() == "")
            //{
            //    MessageBox.Show("Select a database and table that contains data to be paysend", "Information", MessageBoxButtons.OK);

            //}
            //else
            //{
            //    if (TB.TBName == "")
            //    {
            //        MessageBox.Show("Select a table that contains data to be paysend", "Information", MessageBoxButtons.OK);
            //    }
            //    else
            //    {
            if (Base.DBTables.Contains("PAYROLL"))
            {
            }
            else
            {
                if (myConn.State == ConnectionState.Open)
                {
                }
                else
                {
                    myConn.Open();
                }

                //Create a table
                Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");
                if (intCount > 0)
                {
                }
                else
                {
                    TB.createPayrollTable(Base.DBConnectionString);
                }
            }

            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, txtSelectedSection.Text.Trim());
            paysend.Show();
            //}
            //}

        }

        private void emailInfo_Click(object sender, EventArgs e)
        {

        }

        private void basicGraph_Click(object sender, EventArgs e)
        {

        }

        private void drillDownGraph_Click(object sender, EventArgs e)
        {

        }

        private void dataFilter_Click(object sender, EventArgs e)
        {
            if (General.textTestSQL.ToString().Trim().Length > 0)
            {
                scrQuerySQL testsql = new scrQuerySQL();
                testsql.TestSQL(Base.DBConnection, General, Base.DBConnectionString);
                testsql.Show();
            }
            else
            {
                MessageBox.Show("No SQL to pass", "Information", MessageBoxButtons.OK);
            }
        }

        private void dataPrintTables_Click(object sender, EventArgs e)
        {

        }

        private void dataFormulasImportTable_Click(object sender, EventArgs e)
        {
            //Email error information to the standby person

            //OutlookIntegrationEx.MainForm ex = new OutlookIntegrationEx.MainForm();
            //ex.Show();

        }

        private void TBCreateSpreadsheet_Click(object sender, EventArgs e)
        {
            try
            {
                if (openDialog.ShowDialog() != DialogResult.OK) return;
                //grpData.Enabled = false;
                string filename = openDialog.FileName;
                FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                spreadsheet = new ExcelDataReader.ExcelDataReader(fs);
                fs.Close();

                if (spreadsheet.WorkbookData.Tables.Count > 0)
                {
                    switch (string.IsNullOrEmpty(Base.DBName))
                    {
                        case true:
                            MessageBox.Show("Create or select a database.", "DATABASE NEEDED!", MessageBoxButtons.OK);
                            break;

                        case false:
                            saveTheSpreadSheetToTheDatabase();
                            MessageBox.Show("Successfully Uploaded.", "Information", MessageBoxButtons.OK);
                            break;
                        default:

                            break;
                    }
                }

                //cboSheet.Items.Clear();
                //cboSheet.DisplayMember = "TableName";
                //foreach (DataTable dt in spreadsheet.WorkbookData.Tables)
                //    cboSheet.Items.Add(dt);

                //if (cboSheet.Items.Count == 0) return;

                //grpData.Enabled = true;
                //checker = true;
                //cboSheet.SelectedIndex = 0;
                //btnSave.Visible = true;
                //lblSheet.Visible = true;
                //cboSheet.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to read file: \n" + ex.Message);
            }
        }

        private void TBDeleteTable_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Delete table: " + TB.TBName + " ? ", "Confirm", MessageBoxButtons.YesNo);

            switch (result)
            {
                case DialogResult.Yes:
                    bool tableCreate = TB.dropDatabaseTable(Base.DBConnectionString);
                    extractDBTableNames(listBox1);
                    TB.deleteDataTableFromCollection(TB.DBName);
                    TB.TBName = "";
                    TBFormulas.Tablename = "";
                    loadInfo();
                    break;


                case DialogResult.No:
                    break;
            }
        }

        private void TBDeleteCalcColumns_Click(object sender, EventArgs e)
        {
            DialogResult result1 = MessageBox.Show("Confirm DELETE of calculated columns from table: " + TBFormulas.Tablename + "?", "", MessageBoxButtons.YesNo);

            switch (result1)
            {
                case DialogResult.Yes:

                    DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, TB.TBName, Base.AnalysisConnectionString);
                    foreach (DataRow row in tableformulas.Rows)
                    {
                        TB.removeColumn(Base.DBConnectionString, TB.TBName, row["CALC_NAME"].ToString());

                    }
                    loadInfo();
                    break;

                case DialogResult.No:
                    break;
            }
        }

        private void TBDeleteAllTables_Click(object sender, EventArgs e)
        {
            foreach (string s in listBox1.Items)
            {
                TB.TBName = s.Trim();
                bool tableCreate = TB.dropDatabaseTable(Base.DBConnectionString);
            }
            extractDBTableNames(listBox1);
            loadInfo();
        }

        private void DBCreate_Click(object sender, EventArgs e)
        {

        }

        private void createNewDatabase(string Databasename)
        {

        }

        private void DBBackup_Click(object sender, EventArgs e)
        {
            ////The database-tables and formulas will be stored on spreadsheets.

            //if (listBox1.Items.Count == 0)
            //{
            //    MessageBox.Show("No tables to backup", "Backup Failure", MessageBoxButtons.OK);
            //}
            //else
            //{
            //    foreach (string s in listBox1.Items)
            //    {
            //        TB.TBName = s.Trim();
            //        saveTheSpreadSheet();
            //    }
            //}

            ////Extract the formulas of the database
            //extractDatabaseFormulas();
            //TB.TBName = "";

            //Base.backupDatabase3(Base.DBConnectionString, Base.DBName, "D:\\iCalc\\Backups\\Databases");

            //MessageBox.Show("Backup Done to:  D:\\iCalc\\Backups\\Databases ", "Information", MessageBoxButtons.OK);
        }

        private void extractDatabaseFormulas()
        {

        }

        private void DBDeleteList_Click(object sender, EventArgs e)
        {

        }

        private void listDB()
        {

        }

        private void DBList_Click(object sender, EventArgs e)
        {

        }

        private void evaluateStatusButtons()
        {
            btnInsertRow.Enabled = false;
            btnUpdate.Enabled = false;
            btnDeleteRow.Enabled = false;
            btnLoad.Enabled = false;
            btnPrint.Enabled = false;
            btnLock.Enabled = false;

            panelInsert.BackColor = Color.Cornsilk;
            panelUpdate.BackColor = Color.Cornsilk;
            panelDelete.BackColor = Color.Cornsilk;
            panelPreCalcReport.BackColor = Color.Cornsilk;
        }

        private void btnx_Click_1(object sender, EventArgs e)
        {

            btnx.Text = "Running";
            btnx.Enabled = false;
            btnRefresh.Visible = true;
            execute();
            refreshExecution();

        }

        private void refreshExecution()
        {         
            calcTime.Enabled = true;   
        }

        private void execute()
        {

            System.Diagnostics.Process P = new System.Diagnostics.Process();

            switch (BusinessLanguage.Env)
            {
                case "Production":
                    strName = "TshepongStpP";
                    P.StartInfo.WorkingDirectory = @"z:\Harmony\Tshepong\Production\Core";
                    P.StartInfo.FileName = strName + ".exe";


                    pictBox.Visible = true;
                    pictBox2.Visible = true;
                    calcTime.Enabled = true;

                    P.Start();
                    P.Close();
                    break;

                case "Test":
                    strName = "TshepongT";
                    P.StartInfo.WorkingDirectory = "C:\\OEM2";
                    P.StartInfo.FileName = strName + ".exe";

                    pictBox.Visible = true; 
                    pictBox2.Visible = true;  
                    calcTime.Enabled = true; 

                    P.Start();
                    P.Close();
                    break;

                case "Development":

                    strName = "TshepongD";
                    P.StartInfo.WorkingDirectory = @"C:\iCalc\Harmony\Tshepong\Core";
                    P.StartInfo.FileName = strName;

                    pictBox.Visible = true; 
                    pictBox2.Visible = true; 
                    calcTime.Enabled = true; 

                    P.Start();
                    P.Close();
                    break;
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            evaluateStatus();
            evaluateStatusButtons();
        }

        private void btnTeamPrint_Click(object sender, EventArgs e)
        {
            lstStopeReports.Visible = true;

        }

        private void TBExport_Click_1(object sender, EventArgs e)
        {
            saveTheSpreadSheet();
        }

        private void cboNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Search for the coyno in the Labour datatable
            DataTable temp = new DataTable();
            if (Clocked.Rows.Count > 0)
            {
                IEnumerable<DataRow> query1 = from locks in Clocked.AsEnumerable()
                                              where locks.Field<string>("EMPLOYEE_NO").TrimEnd() == cboNames.Text.Trim()
                                              select locks;


                temp = query1.CopyToDataTable<DataRow>();
            }

            if (temp.Rows.Count > 0)
            {
                cboMinersEmpName.Text = temp.Rows[0]["Employee_Name"].ToString().Trim();
            }
            else
            {
                cboMinersEmpName.Text = "xxx";
            }

            if (Labour.Rows.Count > 0)
            {
                IEnumerable<DataRow> query2 = from locks in Labour.AsEnumerable()
                                              where locks.Field<string>("EMPLOYEE_NO").TrimEnd() == cboNames.Text.Trim()
                                              select locks;


                temp = query2.CopyToDataTable<DataRow>();
            }

            if (temp.Rows.Count > 0)
            {
                txtADTeamShifts.Text = temp.Rows[0]["SHIFTS_WORKED"].ToString().Trim();
                txtADTeamShifts.Text = temp.Rows[0]["SHIFTS_WORKED"].ToString().Trim();
                txtAwops.Text = temp.Rows[0]["AWOP_SHIFTS"].ToString().Trim();
            }
            else
            {
                txtPayShifts.Text = "0";
                txtAwops.Text = "0";
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void btnChangePeriod_Click(object sender, EventArgs e)
        {
            //Gets the name of all open forms in application
            foreach (Form form in Application.OpenForms)
            {
                if (form is scrLogon)
                {
                    form.Show(); //Show the form
                    break;
                }
            }
            exitValue = 2;//Change exit value

            this.Close(); //Close the current window

        }

        private void scrTeamS_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (exitValue == 0)
            {
                DialogResult result = MessageBox.Show("Have you saved your data? If not sure, please SAVE.", "REMINDER", MessageBoxButtons.YesNo);

                switch (result)
                {
                    case DialogResult.Yes:
                        //this.Close();
                        //scrMain main = new scrMain();
                        //main.MainLoad(BusinessLanguage, DB, Survey, Labour, Miners, Designations, Occupations, Clocked, EmplList, EmplPen, Configs);
                        //main.ShowDialog();
                        myConn.Close();
                        AAConn.Close();
                        AConn.Close();
                        //this.Close();
                        exitValue = 1;
                        Application.Exit();
                        break;

                    case DialogResult.No:
                        e.Cancel = true;
                        break;
                }
                if (exitValue == 2)
                {
                    exitValue = 1;
                    this.Close();
                }
            }
        }


        private void btnAttendance_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            evaluateLabour();
            if (Labour.Rows.Count == 0)
            {
                MessageBox.Show("No Labour records to print for the section: " + txtSelectedSection.Text.Trim(), "Information", MessageBoxButtons.OK);
            }
            else
            {
                //DataTable temp = Labour.Copy();
                //deleteAllCalcColumnsFromTempTable("BonusShifts", temp);
                //temp.Columns.Remove("TEAMLEADERIND");
                //TB.createAttendanceTable(Base.DBConnectionString, temp);

                //MetaReportRuntime.App mm = new MetaReportRuntime.App();
                //mm.Init(strMetaReportCode);
                //mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
                //mm.StartReport("STPTMTA4000");
                DataTable temp = Labour.Copy();
                deleteAllCalcColumnsFromTempTable("BonusShifts", temp);

                temp.Columns.Remove("TMLEADERIND");
                temp.AcceptChanges();

                //create a view of calendar
                string strSQL = " Drop view CalendarV;";

                Base.InsertData(Base.DBConnectionString, strSQL);

                strSQL = "create view CalendarV as SELECT * from Calendar" +
                                " where period = '" + BusinessLanguage.Period.Trim() + "';";

                Base.InsertData(Base.DBConnectionString, strSQL);

                TB.createAttendanceTable_withPeriod(Base.DBConnectionString, temp);

                MetaReportRuntime.App mm = new MetaReportRuntime.App();
                mm.Init(strMetaReportCode);
                mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
                mm.StartReport("DEVTMTA");

            }
            this.Cursor = Cursors.Arrow;
        }

        private void btnSearchEmployNr_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = true;
            txtSearchGang.Visible = false;
            txtSearchEmplName.Visible = false;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            txtSearchEmplyNr.Focus();
        }

        private void btnEmployName_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = false;
            txtSearchGang.Visible = false;
            txtSearchEmplName.Visible = true;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NAME"], ListSortDirection.Ascending);
            txtSearchEmplName.Focus();
        }

        private void btnSearchGang_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = false;
            txtSearchGang.Visible = true;
            txtSearchEmplName.Visible = false;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["GANG"], ListSortDirection.Ascending);
            txtSearchGang.Focus();
        }

        private void txtSearchEmplyNr_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            searchEmplNr = txtSearchEmplyNr.Text.ToString();
            searchEmplName = "";
            searchEmplGang = "";
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang, grdLabour); //Calls the metod

        }

        private void txtSearchEmplName_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NAME"], ListSortDirection.Ascending);
            searchEmplNr = "";
            searchEmplName = txtSearchEmplName.Text.ToString();
            searchEmplGang = "";
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang, grdLabour); //Calls the metod

        }

        private void txtSearchGang_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["GANG"], ListSortDirection.Ascending);
            searchEmplNr = "";
            searchEmplName = "";
            searchEmplGang = txtSearchGang.Text.ToString();
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang, grdLabour); //Calls the metod
        }

        public void searchBonus(string nr, string name, string gang, DataGridView Grid)
        {
            //Sets the details passed to lower case
            nr = nr.ToLower();
            name = name.ToLower();
            gang = gang.ToLower();

            //Gets the length
            int nrLenght = nr.Length;
            int nameLenght = name.Length;
            int gangLenght = gang.Length;

            // Ensuring length are always 1 and not 0 as
            // "" can not be tested.
            if (nrLenght == 0)
            {
                nrLenght = 1;
            }
            if (nameLenght == 0)
            {
                nameLenght = 1;
            }
            if (gangLenght == 0)
            {
                gangLenght = 1;
            }

            //Iterate through all the rows in the grid
            for (int i = 0; i < Grid.Rows.Count - 1; i++)
            {
                //Gets the values of the grid in the different columns
                string nrColumn = Grid.Rows[i].Cells["Employee_No"].Value.ToString();  //Cells from grid count from left starting at 0
                string nameColumn = Grid.Rows[i].Cells[1].Value.ToString();
                string gangColumn = Grid.Rows[i].Cells["Gang"].Value.ToString();

                //Sets the values from grid to lowercase for testing
                nrColumn = nrColumn.ToLower();
                nameColumn = nameColumn.ToLower();
                gangColumn = gangColumn.ToLower();

                //Gets the same amount from the grid string as was entertered bty the user to 
                //ensure the string can be tested
                nrColumn = nrColumn.Substring(1, nrLenght);//Start at 1 to throw away the aphabetic nr
                nameColumn = nameColumn.Substring(0, nameLenght);
                gangColumn = gangColumn.Substring(0, gangLenght);

                //Compares the different strings
                if (nr == nrColumn) //Employee nr
                {
                    //Empty the string not used
                    nameColumn = "";
                    gangColumn = "";
                    Grid.ClearSelection(); // Clears all past selection
                    Grid.Rows[i].Selected = true; //Selects the current row
                    Grid.FirstDisplayedScrollingRowIndex = i; //Jumps automatically to the row
                    break; //breaks the loop
                }

                if (gang == gangColumn) //Gang
                {
                    nrColumn = "";
                    nameColumn = "";
                    Grid.ClearSelection();
                    Grid.Rows[i].Selected = true;
                    Grid.FirstDisplayedScrollingRowIndex = i;
                    break;
                }
            }
        }

     

        private void dataPrintFormulas_Click(object sender, EventArgs e)
        {
            DataTable dt = Base.dataPrintFormulasBonusShifts(Base.AnalysisConnectionString, Base.DBName, "Production");
            if (dt.Rows.Count > 0)
            {
                printHTML(dt, "Formulas on Production");
            }
            else
            {
                MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
            }
        }

        private void auditByTable_Click(object sender, EventArgs e)
        {
            DataTable audit = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Audit", " where tablename = 'Ganglink'");
            string[] auditcolumns = new string[10];

            string test = audit.Rows[0]["PK"].ToString().Trim();
            int testlength = test.Length;

            for (int i = 0; i <= 9; i++)
            {
                int tstLength = test.IndexOf(">");
                if (tstLength != -1)
                {
                    auditcolumns[i] = test.Substring(0, tstLength).Replace("<", "").Trim();
                    test = test.Substring(test.IndexOf(">") + 1);
                }

            }





        }

        private void btnEmplyeAudit_Click(object sender, EventArgs e)
        {


            #region extract the sheet name and FSH and LSH of the extract
            string FilePath = "C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\ADTeam_201004.xls";
            string[] sheetNames = GetExcelSheetNames(FilePath);
            string sheetName = sheetNames[0];
            #endregion

            #region import Clockshifts
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();

            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 8.0;'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
            da.Fill(dt);

            #region remove invalid records
            // Delete records that does not conform to configurations
            //foreach (DataRow row in dt.Rows)
            //{
            //    if ((row["GANG NAME"].ToString().Substring(5, 1) == "A" || row["GANG NAME"].ToString().Substring(5, 1) == "B" ||
            //        row["GANG NAME"].ToString().Substring(5, 1) == "C" || row["GANG NAME"].ToString().Substring(5, 1) == "D" ||
            //        row["GANG NAME"].ToString().Substring(5, 1) == "E" || row["WAGE CODE"].ToString() == "245M003" ||
            //        row["WAGE CODE"].ToString() == "400M009" || row["WAGE CODE"].ToString() == "245M001" ||
            //        row["WAGE CODE"].ToString() == "246M004" || row["WAGE CODE"].ToString() == "400M009")
            //        && (row["GANG NAME"].ToString().Substring(0, 5) == txtSelectedSection.Text.Trim()))
            //    {
            //    }
            //    else
            //    {
            //        //row.Delete();
            //    }

            //}

            //dt.AcceptChanges();

            //extract the column names with length less than 3.  These columns must be deleted.
            string[] columnNames = new String[dt.Columns.Count];

            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                if (dt.Columns[i].ColumnName.Length <= 2)
                {
                    columnNames[i] = dt.Columns[i].ColumnName;
                }
            }

            for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
            {
                if (string.IsNullOrEmpty(columnNames[i]))
                {

                }
                else
                {
                    dt.Columns.Remove(columnNames[i].ToString().Trim());
                    dt.AcceptChanges();
                }
            }

            dt.Columns.Remove("INDUSTRY NUMBER");
            dt.AcceptChanges();
            #endregion

            string strSheetFSH = string.Empty;
            string strSheetLSH = string.Empty;

            //Extract the dates from the spreadsheet - the name of the spreadsheet contains the the start and enddate of the extract
            string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
            string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

            //Correct the dates and calculate the number of days extracted.
            if (strSheetFSHx.Substring(6, 1) == "-")
            {
                strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
            }

            if (strSheetLSHx.Substring(6, 1) == "-")
            {
                strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
            }

            DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
            DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

            //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
            intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);
            noOFDay = intNoOfDays;

            if (intNoOfDays <= 44)
            {
                for (int j = intNoOfDays + 1; j <= 44; j++)
                {
                    dt.Columns.Add("DAY" + j);
                }
            }
            else
            {

            }

            #region Change the column names
            //Change the column names to the correct column names.
            Dictionary<string, string> dictNames = new Dictionary<string, string>();
            DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                 "Select * from varnames");
            dictNames.Clear();

            dictNames = TB.loadDict(varNames, dictNames);
            int counter = 0;


            //If it is a column with a date as a name.
            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.Substring(0, 1) == "2")
                {
                    if (counter == 0)
                    {
                        strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                        column.ColumnName = "DAY" + counter;
                        counter = counter + 1;

                    }
                    else
                    {
                        if (column.Ordinal == dt.Columns.Count - 1)
                        {

                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;
                        }
                    }


                }
                else
                {
                    if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                    {
                        column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                    }

                }
            }

            //Add the extra columns
            dt.Columns.Add("FSH");
            dt.Columns.Add("LSH");
            dt.Columns.Add("SECTION");
            dt.AcceptChanges();


            foreach (DataRow row in dt.Rows)
            {
                row["FSH"] = strSheetFSH;
                row["LSH"] = strSheetLSH;
                row["MININGTYPE"] = "STOPE";
                if (row["GANG"].ToString().Length > 0)
                {
                    row["SECTION"] = row["GANG"].ToString().Substring(0, 5);
                }
                else
                {
                    row["SECTION"] = "XXX";
                }

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }
            #endregion

            //Write to the database
            // TB.saveCalculations2(dt, Base.DBConnectionString, strWhere, "CLOCKEDSHIFTS");

            // Application.DoEvents();

            // grdClocked.DataSource = dt;
            #endregion

            #region Calculate the shifts per employee en output to bonusshifts

            //string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
            //                "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where (section = '"
            //                + txtSelectedSection.Text.Trim() + "' or WAGE_DESCRIPTION = 'STOPER')";

            string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";

            // BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
            fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix); 

            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            sheetfhs = SheetFSH;
            sheetlhs = SheetLSH;
            intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -44)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 44;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                // DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //switch (result)
                //{
                //    case DialogResult.OK:
                //        extractAndCalcShifts(intStartDay, intStopDay);
                //        break;

                //    case DialogResult.Cancel:
                //        break;

                //}

                #endregion

            #endregion

                #region Extract the ganglinking of the current section
                ////Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                ////before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                ////This is done in the methord extractGangLink

                //DataTable temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "GANGLINK", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current ganglinking for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:
                //            extractGangLink();
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    extractGangLink();
                //}

                //cboMinersGangNo.Items.Clear();
                //lstNames = TB.loadDistinctValuesFromColumn(Labour, "Gang");
                //if (lstNames.Count > 1)
                //{

                //    foreach (string s in lstNames)
                //    {
                //        if (cboMinersGangNo.Items.Contains(s))
                //        { }
                //        else
                //        {
                //            cboMinersGangNo.Items.Add(s.Trim());
                //        }
                //    }
                //}

                #endregion

                #region Extract the miners of the current section
                //Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                //before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                //This is done in the method extractMiners

                //temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "MINERS", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current MINERS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:
                //            extractMiners();
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    extractMiners();
                //}
                #endregion

                fillFixTable(fixShifts, sheetfhs, sheetlhs, intNoOfDays, intStartDay, intStopDay);
                this.Cursor = Cursors.Arrow;
                //}
            }

        }

        public void fillFixTable(DataTable clockedTable, DateTime SheetFSH, DateTime SheetLSH, int intNoOfDays, int DayStart, int DayEnd)
        {
            //Calculate the shifts in the clockedshifts table and insert all in a fixed
            //table that cannot be changed by the user!

            string SQLTable = "IF OBJECT_ID(N'emplshiftfix', N'U')IS NOT NULL DROP TABLE EMPLSHIFTFIX create table EMPLSHIFTFIX (employeeno char(20),shiftsfix char(20)) truncate table EMPLSHIFTFIX";
            Base.VoidQuery(Base.DBConnectionString, SQLTable);

            #region Calculate the shifts per employee en output to bonusshifts

            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -44)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 44;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                int intSubstringLength = 0;
                int intShiftsWorked = 0;
                int intAwopShifts = 0;
                int shiftsCheck = 0;
                StringBuilder sqlInsertFixShifts = new StringBuilder("BEGIN TRANSACTION; ");

                foreach (DataRow row in clockedTable.Rows)
                {
                    foreach (DataColumn column in clockedTable.Columns)
                    {
                        if ((column.ColumnName.Substring(0, 3) == "DAY"))
                        {

                            if (column.ColumnName.ToString().Length == 4)
                            {
                                intSubstringLength = 1;
                            }
                            else
                            {
                                intSubstringLength = 2;
                            }

                            if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                               Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
                            {

                                if (row[column].ToString().Trim() == "U" || row[column].ToString().Trim() == "u" || row[column].ToString().Trim() == "q" || row[column].ToString().Trim() == "Q" || row[column].ToString().Trim() == "W" || row[column].ToString().Trim() == "w")
                                {
                                    intShiftsWorked = intShiftsWorked + 1;
                                    shiftsCheck = 1;
                                }
                                else
                                {
                                    if (row[column].ToString().Trim() == "A")
                                    {
                                        intAwopShifts = intAwopShifts + 1;
                                    }
                                    else { }

                                }
                            }
                            else
                            {
                                row[column] = "*";
                            }
                        }
                        else
                        {
                            if (column.ColumnName == "BONUSTYPE")
                            {
                                row["BONUSTYPE"] = "TEAM";
                            }
                        }
                    }//foreach datacolumn

                    row["SHIFTS_WORKED"] = intShiftsWorked;

                    string emplNr = row["employee_no"].ToString();
                    workedShiftsFixedClockedShift = intShiftsWorked;
                    intShiftsWorked = 0;
                    intAwopShifts = 0;
                    if (shiftsCheck == 1)
                    {
                        sqlInsertFixShifts.Append("INSERT INTO EMPLSHIFTFIX VALUES ('" + emplNr.Trim() + "','" + workedShiftsFixedClockedShift.ToString().Trim() + "');");
                    }
                }

                sqlInsertFixShifts.Append(" COMMIT TRANSACTION");


                Base.VoidQuery(Base.DBConnectionString, sqlInsertFixShifts.ToString());

                //DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //switch (result)
                //{
                //    case DialogResult.OK:
                //        extractAndCalcShifts(intStartDay, intStopDay);
                //        break;

                //    case DialogResult.Cancel:
                //        break;

                //}

                #endregion

            #endregion

            }
        }

        private void lstBErrorLog_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedNr = lstBErrorLog.SelectedItem.ToString();
            if (selectedNr != "Employee Nr")
            {

                selectedNr = selectedNr.Remove(0, 1);
                int last = selectedNr.LastIndexOf("-");
                selectedNr = selectedNr.Remove(last - 1).Trim();
                txtSearchEmplyNr.Visible = true;
                txtSearchEmplyNr.Text = selectedNr;
            }
        }

        private void hideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            grdActiveSheet.Columns[columnnr].Visible = false;


        }


        public void changeRights()
        {
            InputBoxResult pass = InputBox.Show("Password: ");

            string paas = pass.Text.ToString();
            if (pass.ReturnCode == DialogResult.OK)
            {

                if (paas == "admin")
                {
                    txtShifts.Enabled = true;
                    txtPayShifts.Enabled = true;

                }
                else
                {
                    MessageBox.Show("You do not have the right to change this box");
                    txtShifts.Enabled = false;
                    txtPayShifts.Enabled = false;
                    txtShifts.Text = "";
                    txtPayShifts.Text = "";

                }
            }

        }




        private void txtPayShifts_Click(object sender, EventArgs e)
        {
            //changeRights();

        }

        private void txtShifts_Click(object sender, EventArgs e)
        {
            //changeRights();
        }

        private void btnSurveySummary_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.StartReport("SurveySumSTM");
            this.Cursor = Cursors.Arrow;
        }

        private void txtSearchEmplyNr2_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdPayroll.Sort(grdPayroll.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            searchEmplNr2 = txtSearchEmplyNr2.Text.ToString();
            searchEmplGang2 = "";
            searchBonus(searchEmplNr2, "", searchEmplGang2, grdPayroll); //Calls the metod

        }

        private void txtSearchGang2_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdPayroll.Sort(grdPayroll.Columns["GANG"], ListSortDirection.Ascending);
            searchEmplNr2 = "";
            searchEmplGang = txtSearchGang2.Text.ToString();
            searchBonus(searchEmplNr2, "", searchEmplGang2, grdPayroll); //Calls the metod
        }

        private void btnSearchEmployNr2_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr2.Visible = true;
            txtSearchGang2.Visible = false;

            txtSearchEmplyNr2.Text = "";
            txtSearchGang2.Text = "";
            grdPayroll.Sort(grdPayroll.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            txtSearchEmplyNr2.Focus();
        }

        private void btnSearchGang2_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr2.Visible = false;
            txtSearchGang2.Visible = true;

            txtSearchEmplyNr2.Text = "";
            txtSearchGang2.Text = "";
            grdPayroll.Sort(grdPayroll.Columns["GANG"], ListSortDirection.Ascending);
            txtSearchGang2.Focus();
        }

        private void evaluateAbnormal()
        {
            // Display die Abnormal info
            Abnormal.Rows.Clear();

            loadAbnormal();

        }

        private void loadAbnormal()
        {
            //Check if ABNORMAL exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "ABNORMAL");

            if (intCount > 0)
            {
                //YES

                Abnormal = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "ABNORMAL", strWhere);

                if (Abnormal.Rows.Count == 0)
                {
                    string strSQL = "Begin Transaction; Delete from abnormal where section = '" + txtSelectedSection.Text.Trim() + "';" +
                                    " Select '" + BusinessLanguage.BussUnit + "' as BUSSUNIT, '" +
                                    BusinessLanguage.MiningType + "' as MININGTYPE, '" +
                                    BusinessLanguage.BonusType + "' as BONUTTYPE, '" +
                                    txtSelectedSection.Text.Trim() + "' AS SECTION, '" +
                                    BusinessLanguage.Period + "' AS PERIOD,WORKPLACE ," +
                                    "'XXX' AS ABNORMALLEVEL,'XXX' AS ABNORMALTYPE , '0'  AS ABNORMALVALUE " +
                                    "from Survey where section = '" + txtSelectedSection.Text.Trim() + "'; Commit Transaction;  ";

                    DataTable TempDataTable = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                    foreach (DataRow _row in TempDataTable.Rows)
                    {
                        if (string.IsNullOrEmpty(_row[0].ToString()))
                        {
                        }
                        else
                        {
                            Abnormal.Rows.Add(_row.ItemArray);
                        }
                    }

                    TB.saveCalculations2(Abnormal, Base.DBConnectionString, strWhere, "ABNORMAL");

                }
                else
                {

                }
            }
            else
            {

            }

            string sqlWP = "SELECT distinct WORKPLACE FROM PRODUCTION WHERE SECTION = '" + txtSelectedSection.Text.ToString().Trim() + 
                           "' and period = '" + BusinessLanguage.Period + "'";

            DataTable wpnr = new DataTable();

            wpnr = TB.createDataTableWithAdapter(Base.DBConnectionString, sqlWP);
            cboAbnormalWorkplace.Items.Clear();

            foreach (DataRow wp in wpnr.Rows)
            {
                string wpload = wp[0].ToString().Trim();
                cboAbnormalWorkplace.Items.Add(wpload);
            }

            grdAbnormal.DataSource = Abnormal;

            hideColumnsOfGrid("grdAbnormal");

        }


        private void calcTime_Tick(object sender, EventArgs e)
        {
            btnRefresh_Click("Method", null);
        }

        private void grdAbnormal_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {

            }
            else
            {

                cboAbnormalWorkplace.Text = grdAbnormal["WORKPLACE", e.RowIndex].Value.ToString().Trim();
                cboAbnormalLevel.Text = grdAbnormal["ABNORMALLEVEL", e.RowIndex].Value.ToString().Trim();
                txtAbnormalValue.Text = grdAbnormal["ABNORMALVALUE", e.RowIndex].Value.ToString().Trim();
                cboAbnormalType.Text = grdAbnormal["ABNORMALTYPE", e.RowIndex].Value.ToString().Trim();


                btnUpdate.Enabled = true;
                btnDeleteRow.Enabled = true;
            }

            Cursor.Current = Cursors.Arrow;

        }

        private void cboAbnormalLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboAbnormalType.Items.Clear();

            foreach (DataRow dr in Configs.Rows)
            {
                if (dr["PARAMETERNAME"].ToString().Trim() == "ABNORMAL" && dr["Miningtype"].ToString().Trim() == BusinessLanguage.MiningType
                   && dr["BONUSTYPE"].ToString().Trim() == txtBonusType.Text.Trim() && dr["Parm1"].ToString().Trim() == cboAbnormalLevel.SelectedItem.ToString().Trim())
                {
                    for (int i = 5; i <= Configs.Columns.Count - 1; i++)
                    {
                        if (dr[i].ToString().Trim() == "Q")
                        {
                        }
                        else
                        {
                            cboAbnormalType.Items.Add(dr[i].ToString().Trim());
                        }
                    }
                }
            }
        }

        static void ReadReceipts(string path)
        {
            //create the mail message


            MailMessage mail = new MailMessage("vaatjie@gmail.com", "support@icalcsolutions.co.za", "Tshepong Stope Team", "Stope Team Bonus.");

            Attachment attachment = new Attachment(path); //create the attachment
            mail.Attachments.Add(attachment);	//add the attachment
            SmtpClient client = new SmtpClient(); //your real server goes here
            client.Credentials = CredentialCache.DefaultNetworkCredentials;
            client.Host = "smtp.gmail.com";
            client.Port = 587;
            client.EnableSsl = true;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential("vaatjie@gmail.com", "annel01");

            try
            {
                client.Timeout = 10000000;
                client.Send(mail);

                MessageBox.Show("Mail was sent succesfull!");
            }
            catch (Exception)
            {
                MessageBox.Show("Mail was not succesfull!");
                throw;
            }

        }

        private bool createZipFolder(string path, string databasename)
        {
            path = Base.BackupPath.Replace(Base.BackupPath.Substring(0, 2), "C:") + "\\" + databasename + DateTime.Today.ToString("yyyyMMdd");
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void FastZipCompress(string pathDBBackup, string zipname)
        {
            FastZip fZip = new FastZip();

            fZip.CreateZip("C:\\icalc\\" + zipname + ".zip", pathDBBackup.Replace("xxx.bak", ""), false, ".bak$");
        }


        private void BackupDB(string connectionstring, string dbname, string dbPath)
        {
            Cursor.Current = Cursors.Arrow;
            bool check = false;
            check = Base.backupDatabase3(connectionstring, dbname, dbPath);

            //Copy the file to the C:\drive
            if (check == true)
            {
                //MessageBox.Show("Source = " + dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2) + "\\ICALC", "X:") + 
                //                dbname + DateTime.Today.ToString("yyyyMMdd") + ".bak", "Information", MessageBoxButtons.OK);

                Path = dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2), "C:") + dbname +
                       DateTime.Today.ToString("yyyyMMdd") + " \\\\";

                createZipFolder(Path, dbname);

                //MessageBox.Show("dest = " + Path + dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak", "Information", MessageBoxButtons.OK);
                check = BusinessLanguage.copyBackupFile(dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2) +
                        "\\ICALC", "Z:") + dbname + DateTime.Today.ToString("yyyyMMdd") + ".bak",
                        Path + dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak");

                if (check == true)
                {
                    string filename = dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak";
                    FastZipCompress(Path + "\\", dbname + DateTime.Today.ToString("yyyyMMdd"));
                    DialogResult checks = MessageBox.Show("Backup Done to : " + Path, "Information", MessageBoxButtons.YesNo);

                }
                else
                {
                    MessageBox.Show("Copy unsuccessfull from : " + dbPath.Substring(0, 2) + "   Copy unsuccessfull to :" + dbPath.Replace(dbPath.Substring(0, 2), "C:"), "Information", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Backup unsuccessfull to : " + dbPath.Replace(dbPath.Substring(0, 2), "C:"), "Information", MessageBoxButtons.OK);
            }

            Cursor.Current = Cursors.Arrow;

        }
        private void defaultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BackupDB(Base.DBConnectionString, Base.DBName, Base.BackupPath);
        }

        private void btnMetervs_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.StartReport("MetersVSPayoutStope");
            this.Cursor = Cursors.Arrow;
        }

        private void cboDrillerInd_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtDrillerShifts.Focus();
        }





        private void btnOpenAcess_Click(object sender, EventArgs e)
        {
            if (btnOpenAcess.Text == "Open Access")
            {

                InputBoxResult result = InputBox.Show("Password: ", "To Change Bonusshifts permision are needed!", "*", "0");

                if (result.ReturnCode == DialogResult.OK)
                {
                    if (result.Text.Trim() == "Moses")
                    {
                        pannelWorking.Enabled = true;
                        txtShifts.Enabled = true;
                        btnOpenAcess.Text = "Close Access";
                    }
                }
            }
            else
            {
                pannelWorking.Enabled = false;
                btnOpenAcess.Text = "Open Access";
            }
        }

        private void btnManualSend_Click(object sender, EventArgs e)
        {
            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, "MANUAL");
            paysend.Show();
        }

        private void cboRefresh_SelectedIndexChanged(object sender, EventArgs e)
        {
                refreshShifts();
        }


        private void refreshShifts()
        {
            pictBox.Visible = true;

            #region extract the sheet name and FSH and LSH of the extract
            //MessageBox.Show("maak nou instance van excel");
            ATPMain.VkExcel excel = new ATPMain.VkExcel(false);


            bool XLSX_exists = File.Exists("C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx");
            bool XLS_exists = File.Exists("C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls");

            if (XLSX_exists.Equals(true))
            {
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx", "BONTS2011");
                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            if (XLS_exists.Equals(true))
            {
                //MessageBox.Show("nou in xls filepath");
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls", "BONTS2011");

                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            excel.stopExcel();

            string FilePath = "";

            string FilePath_XLSX = "C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";

            string FilePath_XLS = "C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";

            XLSX_exists = File.Exists(FilePath_XLSX);
            XLS_exists = File.Exists(FilePath_XLS);

            if (XLS_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";
            }

            if (XLSX_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";
            }




            //excel.GetExcelSheets();
            string[] sheetNames = GetExcelSheetNames(FilePath);
            string sheetName = sheetNames[0];

            string testString = sheetName.Substring(0, 3).ToString().Trim();


            if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
            {
                sheetName = sheetNames[1];
            }

            if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
            {
                sheetName = sheetNames[2];
            }

            if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
            {
                sheetName = sheetNames[3];
            }
            #endregion

            #region import Clockshifts
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();

            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 8.0;'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
            da.Fill(dt);
            IEnumerable<DataRow> query1 = from locks in dt.AsEnumerable()
                                          where locks.Field<string>("MINING PROCESS").TrimEnd() != "Development"
                                          select locks;



            //Temp will contain a list of the gangs for the section
            DataTable Tempdt = query1.CopyToDataTable<DataRow>();

            dt = Tempdt.Copy();
            #region remove invalid records

            //extract the column names with length less than 3.  These columns must be deleted.
            string[] columnNames = new String[dt.Columns.Count];

            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                if (dt.Columns[i].ColumnName.Length <= 2)
                {
                    columnNames[i] = dt.Columns[i].ColumnName;
                }
            }

            for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
            {
                if (string.IsNullOrEmpty(columnNames[i]))
                {

                }
                else
                {
                    dt.Columns.Remove(columnNames[i].ToString().Trim());
                    dt.AcceptChanges();
                }
            }

            dt.Columns.Remove("INDUSTRY NUMBER");
            dt.AcceptChanges();
            #endregion

            string strSheetFSH = string.Empty;
            string strSheetLSH = string.Empty;

            //Extract the dates from the spreadsheet - the name of the spreadsheet contains the the start and enddate of the extract
            string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
            string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

            //Correct the dates and calculate the number of days extracted.
            string tested = strSheetFSHx.Substring(6, 1);
            if (strSheetFSHx.Substring(6, 1) == "-")
            {
                strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
            }
            else
            {
                strSheetFSH = strSheetFSHx;
            }


            if (strSheetLSHx.Substring(6, 1) == "-")
            {
                strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
            }
            else
            {
                strSheetLSH = strSheetLSHx;
            }

            DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
            DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

            //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
            int intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);

            if (intNoOfDays <= 44)
            {
                for (int j = intNoOfDays + 1; j <= 44; j++)
                {
                    dt.Columns.Add("DAY" + j);
                }
            }
            else
            {

            }

            #region Change the column names
            //Change the column names to the correct column names.
            Dictionary<string, string> dictNames = new Dictionary<string, string>();
            DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                 "Select * from varnames");
            dictNames.Clear();

            dictNames = TB.loadDict(varNames, dictNames);
            int counter = 0;

            //If it is a column with a date as a name.
            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.Substring(0, 1) == "2")
                {
                    if (counter == 0)
                    {
                        strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                        column.ColumnName = "DAY" + counter;
                        counter = counter + 1;

                    }
                    else
                    {
                        if (column.Ordinal == dt.Columns.Count - 1)
                        {

                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;
                        }
                    }


                }
                else
                {
                    if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                    {
                        column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                    }

                }
            }

            //Add the extra columns
            dt.Columns.Add("FSH");
            dt.Columns.Add("LSH");
            dt.Columns.Add("SECTION");
            dt.Columns.Add("EMPLOYEETYPE");
            dt.AcceptChanges();


            foreach (DataRow row in dt.Rows)
            {
                row["FSH"] = strSheetFSH;
                row["LSH"] = strSheetLSH;
                row["MININGTYPE"] = "STOPING";
                row["SECTION"] = row["GANG"].ToString().Substring(0, 5);
                row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }
            #endregion
            //exportToExcel("c:\\", dt);
            //Write to the database
            TB.saveCalculations2(dt, Base.DBConnectionString, "", "CLOCKEDSHIFTS");

            Application.DoEvents();
            Clocked = dt.Copy();
            grdClocked.DataSource = Clocked;

            #endregion

            #region Calculate the shifts per employee en output to bonusshifts

            string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
                            "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where section = '" +
                            txtSelectedSection.Text.Trim() + "'";

            string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";


            if (BusinessLanguage.MiningType == "STOPE")
            {
                //strSQL = strSQL.Trim() + " and bonustype = 'Stoping' ";
            }
            else
            {
                //if (BusinessLanguage.MiningType == "DEVELOPMENT")
                //{
                strSQL = strSQL.Trim();
                //+ " and bonustype = 'Development' ";
                //}
            }

            fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix); 
            BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
            //exportToExcel("c:\\", BonusShifts);
            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            sheetfhs = SheetFSH;
            sheetlhs = SheetLSH;
            int intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            int intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            int intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -44)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 44;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                DialogResult result = MessageBox.Show("Do you want to REFRESH the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                switch (result)
                {
                    case DialogResult.OK:
                        extractAndCalcShiftsForRefresh(intStartDay, intStopDay);
                        break;

                    case DialogResult.Cancel:
                        break;

                }

                #endregion

            #endregion

                #region Extract the ganglinking of the current section
                //Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                //before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                //This is done in the methord extractGangLink

                //DataTable temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "GANGLINK", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current ganglinking for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:

                //            Base.dropTrigger(Base.DBConnectionString, "Ganglink");
                //            extractGangLink();
                //            Base.createTrigger(Base.DBConnectionString, "Ganglink");
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    Base.dropTrigger(Base.DBConnectionString, "Ganglink");
                //    extractGangLink();
                //    Base.createTrigger(Base.DBConnectionString, "Ganglink");
                //}

                //cboMinersGangNo.Items.Clear();
                //lstNames = TB.loadDistinctValuesFromColumn(Labour, "Gang");
                //if (lstNames.Count > 1)
                //{

                //    foreach (string s in lstNames)
                //    {
                //        if (cboMinersGangNo.Items.Contains(s))
                //        { }
                //        else
                //        {
                //            cboMinersGangNo.Items.Add(s.Trim());
                //        }
                //    }
                //}

                #endregion

                #region Extract the miners of the current section
                //Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                //before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                //This is done in the method extractMiners

                //temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "MINERS", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current MINERS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:
                //            Base.dropTrigger(Base.DBConnectionString, "Miners");
                //            extractMiners();
                //            Base.createTrigger(Base.DBConnectionString, "Miners");
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    Base.dropTrigger(Base.DBConnectionString, "Miners");
                //    extractMiners();
                //    Base.createTrigger(Base.DBConnectionString, "Miners");

                //}
                #endregion

                #region Extract the ganglinking

                //strSQL = "BEGIN transaction; Delete from ganglink where section = '" + txtSelectedSection.Text.Trim() + "';Select '" +
                //         BusinessLanguage.BussUnit + "' as BUSSUNIT,'" + BusinessLanguage.MiningType + "' AS MININGTYPE,'" +
                //         BusinessLanguage.BonusType + "' AS BONUSTYPE,'" + txtSelectedSection.Text.Trim() +
                //         "' AS SECTION,t1.period AS PERIOD,t1.workplace,t1.gang AS GANG," +
                //         "'0'  AS SAFETYIND,'DEVELOPMENT' AS GANGTYPE " +
                //         " from production as t1" +
                //         " where t1.section = '" + txtSelectedSection.Text.Trim() + "' UNION Select '" +
                //         BusinessLanguage.BussUnit + "' as BUSSUNIT,'" + BusinessLanguage.MiningType + "' AS MININGTYPE,'" +
                //         BusinessLanguage.BonusType + "' AS BONUSTYPE,'" + txtSelectedSection.Text.Trim() +
                //         "' AS SECTION,t1.period AS PERIOD,t1.workplace,t1.gang AS GANG," +
                //         "'0'  AS SAFETYIND,'RIGGING' AS GANGTYPE " +
                //         " from production as t1" +
                //         " where t1.section = '" + txtSelectedSection.Text.Trim() + "' UNION Select '" +
                //         BusinessLanguage.BussUnit + "' as BUSSUNIT,'" + BusinessLanguage.MiningType + "' AS MININGTYPE,'" +
                //         BusinessLanguage.BonusType + "' AS BONUSTYPE,'" + txtSelectedSection.Text.Trim() +
                //         "' AS SECTION,t1.period AS PERIOD,t1.workplace,t1.gang AS GANG," +
                //         "'0'  AS SAFETYIND,'EQUIPPING' AS GANGTYPE " +
                //         " from production as t1" +
                //         " where t1.section = '" + txtSelectedSection.Text.Trim() + "' UNION Select '" +
                //         BusinessLanguage.BussUnit + "' as BUSSUNIT,'" + BusinessLanguage.MiningType + "' AS MININGTYPE,'" +
                //         BusinessLanguage.BonusType + "' AS BONUSTYPE,'" + txtSelectedSection.Text.Trim() +
                //         "' AS SECTION,t1.period AS PERIOD,t1.workplace,t1.gang AS GANG," +
                //         "'0'  AS SAFETYIND,'TRAMMING' AS GANGTYPE " +
                //         " from production as t1" +
                //         " where t1.section = '" + txtSelectedSection.Text.Trim() + "'; Commit Transaction;";

                //DataTable TmpGanglink = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                //foreach (DataRow row in TmpGanglink.Rows)
                //{
                //    if (row["GANGTYPE"].ToString() == "RIGGING")
                //    {
                //        row["GANG"] = row["GANG"].ToString().Substring(0, 5) + "C" + row["GANG"].ToString().Substring(6);

                //    }
                //    else
                //    {
                //        if (row["GANGTYPE"].ToString() == "TRAMMING")
                //        {
                //            row["GANG"] = row["GANG"].ToString().Substring(0, 5) + "E" + row["GANG"].ToString().Substring(6);

                //        }
                //        else
                //        {
                //            if (row["GANGTYPE"].ToString() == "EQUIPPING")
                //            {
                //                row["GANG"] = row["GANG"].ToString().Substring(0, 5) + "B" + row["GANG"].ToString().Substring(6);

                //            }
                //            else
                //            {
                //                row["GANGTYPE"] = "DEVELOPMENT";
                //            }
                //        }
                //    }
                //}



                //if (TmpGanglink.Rows.Count > 0)
                //{

                //    TB.saveCalculations2(TmpGanglink, Base.DBConnectionString, strWhere, "GANGLINK");
                //    Application.DoEvents();

                //}
                //else
                //{
                //    MessageBox.Show("No records for ganglinking were extracted for section: " + txtSelectedSection.Text.Trim(), "Information", MessageBoxButtons.OK);
                //}

                #endregion



                this.Cursor = Cursors.Arrow;
                File.Delete(FilePath);
                pictBox.Visible = false;
            }
        }

        private void refreshShifts2()
        {
            pictBox.Visible = true;

            #region extract the sheet name and FSH and LSH of the extract
            //MessageBox.Show("maak nou instance van excel");
            ATPMain.VkExcel excel = new ATPMain.VkExcel(false);


            bool XLSX_exists = File.Exists("C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx");
            bool XLS_exists = File.Exists("C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls");

            if (XLSX_exists.Equals(true))
            {
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx", "BONTS2011");
                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            if (XLS_exists.Equals(true))
            {
                //MessageBox.Show("nou in xls filepath");
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls", "BONTS2011");

                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            excel.stopExcel();

            string FilePath = "";

            string FilePath_XLSX = "C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";

            string FilePath_XLS = "C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";

            XLSX_exists = File.Exists(FilePath_XLSX);
            XLS_exists = File.Exists(FilePath_XLS);

            if (XLS_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";
            }

            if (XLSX_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Tshepong\\Development\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";
            }




            //excel.GetExcelSheets();
            string[] sheetNames = GetExcelSheetNames(FilePath);
            string sheetName = sheetNames[0];

            string testString = sheetName.Substring(0, 3).ToString().Trim();


            if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
            {
                sheetName = sheetNames[1];
            }

            if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
            {
                sheetName = sheetNames[2];
            }

            if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
            {
                sheetName = sheetNames[3];
            }
            #endregion

            #region import Clockshifts
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();

            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 8.0;'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
            da.Fill(dt);
            IEnumerable<DataRow> query1 = from locks in dt.AsEnumerable()
                                          where locks.Field<string>("MINING PROCESS").TrimEnd() != "Development"
                                          select locks;



            //Temp will contain a list of the gangs for the section
            DataTable Tempdt = query1.CopyToDataTable<DataRow>();

            dt = Tempdt.Copy();
            #region remove invalid records

            //extract the column names with length less than 3.  These columns must be deleted.
            string[] columnNames = new String[dt.Columns.Count];

            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                if (dt.Columns[i].ColumnName.Length <= 2)
                {
                    columnNames[i] = dt.Columns[i].ColumnName;
                }
            }

            for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
            {
                if (string.IsNullOrEmpty(columnNames[i]))
                {

                }
                else
                {
                    dt.Columns.Remove(columnNames[i].ToString().Trim());
                    dt.AcceptChanges();
                }
            }

            dt.Columns.Remove("INDUSTRY NUMBER");
            dt.AcceptChanges();
            #endregion

            string strSheetFSH = string.Empty;
            string strSheetLSH = string.Empty;

            //Extract the dates from the spreadsheet - the name of the spreadsheet contains the the start and enddate of the extract
            string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
            string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

            //Correct the dates and calculate the number of days extracted.
            string tested = strSheetFSHx.Substring(6, 1);
            if (strSheetFSHx.Substring(6, 1) == "-")
            {
                strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
            }
            else
            {
                strSheetFSH = strSheetFSHx;
            }


            if (strSheetLSHx.Substring(6, 1) == "-")
            {
                strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
            }
            else
            {
                strSheetLSH = strSheetLSHx;
            }

            DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
            DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

            //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
            int intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);

            if (intNoOfDays <= 44)
            {
                for (int j = intNoOfDays + 1; j <= 44; j++)
                {
                    dt.Columns.Add("DAY" + j);
                }
            }
            else
            {

            }

            #region Change the column names
            //Change the column names to the correct column names.
            Dictionary<string, string> dictNames = new Dictionary<string, string>();
            DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                 "Select * from varnames");
            dictNames.Clear();

            dictNames = TB.loadDict(varNames, dictNames);
            int counter = 0;

            //If it is a column with a date as a name.
            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.Substring(0, 1) == "2")
                {
                    if (counter == 0)
                    {
                        strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                        column.ColumnName = "DAY" + counter;
                        counter = counter + 1;

                    }
                    else
                    {
                        if (column.Ordinal == dt.Columns.Count - 1)
                        {

                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;
                        }
                    }


                }
                else
                {
                    if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                    {
                        column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                    }

                }
            }

            //Add the extra columns
            dt.Columns.Add("FSH");
            dt.Columns.Add("LSH");
            dt.Columns.Add("SECTION");
            dt.Columns.Add("EMPLOYEETYPE");
            dt.AcceptChanges();


            foreach (DataRow row in dt.Rows)
            {
                row["FSH"] = strSheetFSH;
                row["LSH"] = strSheetLSH;
                row["MININGTYPE"] = "STOPING";
                row["SECTION"] = row["GANG"].ToString().Substring(0, 5);
                row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }
            #endregion
            //exportToExcel("c:\\", dt);
            //Write to the database
            TB.saveCalculations2(dt, Base.DBConnectionString, "", "CLOCKEDSHIFTS");

            Application.DoEvents();
            Clocked = dt.Copy();
            grdClocked.DataSource = Clocked;

            #endregion

            #region Calculate the shifts per employee en output to bonusshifts

            string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
                            "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where section = '" +
                            txtSelectedSection.Text.Trim() + "'";

            string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";


            if (BusinessLanguage.MiningType == "STOPE")
            {
                //strSQL = strSQL.Trim() + " and bonustype = 'Stoping' ";
            }
            else
            {
                //if (BusinessLanguage.MiningType == "DEVELOPMENT")
                //{
                strSQL = strSQL.Trim();
                //+ " and bonustype = 'Development' ";
                //}
            }

            fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix); 
            BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
            //exportToExcel("c:\\", BonusShifts);
            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            sheetfhs = SheetFSH;
            sheetlhs = SheetLSH;
            int intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            int intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            int intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -44)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 44;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                DialogResult result = MessageBox.Show("Do you want to REFRESH the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                switch (result)
                {
                    case DialogResult.OK:
                        extractAndCalcShiftsForRefresh(intStartDay, intStopDay);
                        break;

                    case DialogResult.Cancel:
                        break;

                }

                #endregion

            #endregion

                #region Extract the ganglinking of the current section
                //Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                //before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                //This is done in the methord extractGangLink

                //DataTable temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "GANGLINK", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current ganglinking for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:

                //            Base.dropTrigger(Base.DBConnectionString, "Ganglink");
                //            extractGangLink();
                //            Base.createTrigger(Base.DBConnectionString, "Ganglink");
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    Base.dropTrigger(Base.DBConnectionString, "Ganglink");
                //    extractGangLink();
                //    Base.createTrigger(Base.DBConnectionString, "Ganglink");
                //}

                //cboMinersGangNo.Items.Clear();
                //lstNames = TB.loadDistinctValuesFromColumn(Labour, "Gang");
                //if (lstNames.Count > 1)
                //{

                //    foreach (string s in lstNames)
                //    {
                //        if (cboMinersGangNo.Items.Contains(s))
                //        { }
                //        else
                //        {
                //            cboMinersGangNo.Items.Add(s.Trim());
                //        }
                //    }
                //}

                #endregion

                #region Extract the miners of the current section
                //Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                //before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                //This is done in the method extractMiners

                //temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "MINERS", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current MINERS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:
                //            Base.dropTrigger(Base.DBConnectionString, "Miners");
                //            extractMiners();
                //            Base.createTrigger(Base.DBConnectionString, "Miners");
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    Base.dropTrigger(Base.DBConnectionString, "Miners");
                //    extractMiners();
                //    Base.createTrigger(Base.DBConnectionString, "Miners");

                //}
                #endregion

                #region Extract the ganglinking

                //strSQL = "BEGIN transaction; Delete from ganglink where section = '" + txtSelectedSection.Text.Trim() + "';Select '" +
                //         BusinessLanguage.BussUnit + "' as BUSSUNIT,'" + BusinessLanguage.MiningType + "' AS MININGTYPE,'" +
                //         BusinessLanguage.BonusType + "' AS BONUSTYPE,'" + txtSelectedSection.Text.Trim() +
                //         "' AS SECTION,t1.period AS PERIOD,t1.workplace,t1.gang AS GANG," +
                //         "'0'  AS SAFETYIND,'DEVELOPMENT' AS GANGTYPE " +
                //         " from production as t1" +
                //         " where t1.section = '" + txtSelectedSection.Text.Trim() + "' UNION Select '" +
                //         BusinessLanguage.BussUnit + "' as BUSSUNIT,'" + BusinessLanguage.MiningType + "' AS MININGTYPE,'" +
                //         BusinessLanguage.BonusType + "' AS BONUSTYPE,'" + txtSelectedSection.Text.Trim() +
                //         "' AS SECTION,t1.period AS PERIOD,t1.workplace,t1.gang AS GANG," +
                //         "'0'  AS SAFETYIND,'RIGGING' AS GANGTYPE " +
                //         " from production as t1" +
                //         " where t1.section = '" + txtSelectedSection.Text.Trim() + "' UNION Select '" +
                //         BusinessLanguage.BussUnit + "' as BUSSUNIT,'" + BusinessLanguage.MiningType + "' AS MININGTYPE,'" +
                //         BusinessLanguage.BonusType + "' AS BONUSTYPE,'" + txtSelectedSection.Text.Trim() +
                //         "' AS SECTION,t1.period AS PERIOD,t1.workplace,t1.gang AS GANG," +
                //         "'0'  AS SAFETYIND,'EQUIPPING' AS GANGTYPE " +
                //         " from production as t1" +
                //         " where t1.section = '" + txtSelectedSection.Text.Trim() + "' UNION Select '" +
                //         BusinessLanguage.BussUnit + "' as BUSSUNIT,'" + BusinessLanguage.MiningType + "' AS MININGTYPE,'" +
                //         BusinessLanguage.BonusType + "' AS BONUSTYPE,'" + txtSelectedSection.Text.Trim() +
                //         "' AS SECTION,t1.period AS PERIOD,t1.workplace,t1.gang AS GANG," +
                //         "'0'  AS SAFETYIND,'TRAMMING' AS GANGTYPE " +
                //         " from production as t1" +
                //         " where t1.section = '" + txtSelectedSection.Text.Trim() + "'; Commit Transaction;";

                //DataTable TmpGanglink = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                //foreach (DataRow row in TmpGanglink.Rows)
                //{
                //    if (row["GANGTYPE"].ToString() == "RIGGING")
                //    {
                //        row["GANG"] = row["GANG"].ToString().Substring(0, 5) + "C" + row["GANG"].ToString().Substring(6);

                //    }
                //    else
                //    {
                //        if (row["GANGTYPE"].ToString() == "TRAMMING")
                //        {
                //            row["GANG"] = row["GANG"].ToString().Substring(0, 5) + "E" + row["GANG"].ToString().Substring(6);

                //        }
                //        else
                //        {
                //            if (row["GANGTYPE"].ToString() == "EQUIPPING")
                //            {
                //                row["GANG"] = row["GANG"].ToString().Substring(0, 5) + "B" + row["GANG"].ToString().Substring(6);

                //            }
                //            else
                //            {
                //                row["GANGTYPE"] = "DEVELOPMENT";
                //            }
                //        }
                //    }
                //}



                //if (TmpGanglink.Rows.Count > 0)
                //{

                //    TB.saveCalculations2(TmpGanglink, Base.DBConnectionString, strWhere, "GANGLINK");
                //    Application.DoEvents();

                //}
                //else
                //{
                //    MessageBox.Show("No records for ganglinking were extracted for section: " + txtSelectedSection.Text.Trim(), "Information", MessageBoxButtons.OK);
                //}

                #endregion



                this.Cursor = Cursors.Arrow;
                File.Delete(FilePath);
                pictBox.Visible = false;
            }
        }




        private void extractAndCalcShiftsForRefresh(int DayStart, int DayEnd)
        {
            int intSubstringLength = 0;
            int intShiftsWorked = 0;
            int intAwopShifts = 0;
            int shiftsCheck = 0;
            BonusShifts.Columns.Add("TMLEADERIND");

            foreach (DataRow row in BonusShifts.Rows)
            {
                foreach (DataColumn column in BonusShifts.Columns)
                {
                    if ((column.ColumnName.Substring(0, 3) == "DAY"))
                    {
                        if (column.ColumnName.ToString().Length == 4)
                        {
                            intSubstringLength = 1;
                        }
                        else
                        {
                            intSubstringLength = 2;
                        }

                        if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                           Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
                        {
                            if (row[column].ToString().Trim() == "U" || row[column].ToString().Trim() == "u" || row[column].ToString().Trim() == "q" || row[column].ToString().Trim() == "Q" || row[column].ToString().Trim() == "W" || row[column].ToString().Trim() == "w")
                            {
                                intShiftsWorked = intShiftsWorked + 1;
                                shiftsCheck = 1;
                            }
                            else
                            {
                                if (row[column].ToString().Trim() == "A")
                                {
                                    intAwopShifts = intAwopShifts + 1;
                                }
                                else { }

                            }
                        }
                        else
                        {
                            row[column] = "*";
                        }
                    }
                    else
                    {
                        if (column.ColumnName == "BONUSTYPE")
                        {
                            row["BONUSTYPE"] = txtBonusType.Text.ToString();
                        }
                    }
                }//foreach datacolumn

                row["SHIFTS_WORKED"] = intShiftsWorked;
                row["AWOP_SHIFTS"] = intAwopShifts;
                row["TMLEADERIND"] = "0";
                intShiftsWorked = 0;
                intAwopShifts = 0;
            }

            //// Query the Bonusshifts for each HOD's
           if(cboRefresh.Text.ToString().Trim() == "Bonusshifts - Refresh Shifts")
           {
               updateShifts(BonusShifts);
           }

           if (cboRefresh.Text.ToString().Trim() == "Bonusshifts - Refresh Employees")
           {
               insertEmployee(BonusShifts);
           }







            if (importdone == 0)
            {
                fillFixTable(fixShifts, sheetfhs, sheetlhs, noOFDay, DayStart, DayEnd);//Calls the method to load the fix clockedshiftstable
                importdone = 1;
            }

            Application.DoEvents();
        }

        public void updateShifts(DataTable BonusShifts)
        {

            foreach (DataRow row in BonusShifts.Rows)
            {
                IEnumerable<DataRow> query1 = from rec in BonusShifts.AsEnumerable()
                                              where rec.Field<string>("EMPLOYEE_NO").Trim() == row["EMPLOYEE_NO"].ToString().Trim()
                                              where rec.Field<string>("Gang").Trim() == row["GANG"].ToString().Trim()
                                              where rec.Field<string>("WAGECODE").Trim() == row["WAGECODE"].ToString().Trim()
                                              select rec;

                DataTable testTB = query1.CopyToDataTable<DataRow>();

                if (testTB.Rows.Count == 1)
                {
                    string update = "Update Bonusshifts set shifts_worked = '" + (Convert.ToInt32(testTB.Rows[0]["Shifts_Worked"].ToString().Trim()) +
                                                "', Awop_shifts  = '" + testTB.Rows[0]["Awop_Shifts"].ToString().Trim() +
                                                "' where employee_no = '" + testTB.Rows[0]["EMPLOYEE_NO"].ToString().Trim() + "' AND Gang = '" + testTB.Rows[0]["GANG"].ToString().Trim() + "'");
                    //Convert.ToInt32(BonusShifts.Rows[0]["Shifts_Worked"].ToString().Trim()
                    TB.InsertData(Base.DBConnectionString, update);
                }
                else
                {

                }
            }
        }

        public void insertEmployee(DataTable BonusShifts)
        {
            DataTable newMembers = new DataTable();
            DataTable bonusShiftCurrent = new DataTable();
            bonusShiftCurrent = TB.createDataTableWithAdapter(Base.DBConnectionString,"select * from bonusshifts where section = '"+txtSelectedSection.Text.ToString().Trim()+"'");
            DataTable bonusShiftCurrent2 = new DataTable();
            bonusShiftCurrent2 = TB.createDataTableWithAdapter(Base.DBConnectionString, "select * from bonusshifts where section = '" + txtSelectedSection.Text.ToString().Trim() + "'");
           
            foreach (DataRow row in BonusShifts.Rows)
            {
                IEnumerable<DataRow> query1 = from rec in BonusShifts.AsEnumerable()
                                              where rec.Field<string>("EMPLOYEE_NO").Trim() == row["EMPLOYEE_NO"].ToString().Trim()
                                              where rec.Field<string>("Gang").Trim() == row["GANG"].ToString().Trim()
                                              where rec.Field<string>("WAGECODE").Trim() == row["WAGECODE"].ToString().Trim()
                                              where rec.Field<string>("SECTION").Trim() == txtSelectedSection.Text.ToString().Trim()
                                              select rec;

                DataTable testTB = query1.CopyToDataTable<DataRow>();

                 bool alreadyOn = false;
                string TEST = testTB.Rows[0]["EMPLOYEE_NO"].ToString().Trim();

              

                foreach (DataRow current in bonusShiftCurrent.Rows)
                {
                    string TEST2 = current["EMPLOYEE_NO"].ToString().Trim();
                    if (current["EMPLOYEE_NO"].ToString().Trim() == testTB.Rows[0]["EMPLOYEE_NO"].ToString().Trim() && current["GANG"].ToString().Trim() == testTB.Rows[0]["GANG"].ToString().Trim() && current["WAGECODE"].ToString().Trim() == testTB.Rows[0]["WAGECODE"].ToString().Trim())
                    {
                        alreadyOn = true;
                    }
                }

                if (alreadyOn == false)
                {
                    foreach (DataRow newMem in testTB.Rows)
                    {
                        DataRow fff = newMem;
                        bonusShiftCurrent2.Rows.Add(fff.ItemArray);
                    }

                   bonusShiftCurrent2.AcceptChanges();

                }
            }

            TB.saveCalculations2(bonusShiftCurrent2, Base.DBConnectionString, "where section ='" + txtSelectedSection.Text.ToString().Trim()+"'", "BONUSSHIFTS");
        }

        private void btnSectionSelect_Click(object sender, EventArgs e)
        {
            listBox2.SelectedIndex = 0;
            listBox2.Select();
            listBox2.Focus();
        }

        private void brtnSetCalendar_Click(object sender, EventArgs e)
        {
            openTab(tabCalendar);
        }

        private void btnGanglink_Click(object sender, EventArgs e)
        {
            openTab(tabGangLinking);
        }

        private void btnMiners_Click(object sender, EventArgs e)
        {
            openTab(tabMiners);
        }

        private void btnCalc_Click(object sender, EventArgs e)
        {
            btnBaseCalcsHeader_Click("me", e);
        }

        private void cboMinersEmpName_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboMinersEmpName.Text = cboMinersEmpName.SelectedItem.ToString().Trim();

            string strNames = "select * from CLOCKEDSHIFTS";

            Workers = TB.createDataTableWithAdapter(Base.DBConnectionString, strNames);

            foreach (DataRow Work in Workers.Rows)
            {
                if (Work["EMPLOYEE_NAME"].ToString().Trim() == cboMinersEmpName.Text.ToString().Trim())
                {
                    cboNames.Text = Work["EMPLOYEE_NO"].ToString().Trim();
                }
            }            
        }

        private void tabInfo_MouseClick(object sender, MouseEventArgs e)
        {
            
        }

        private void btnRefreshGanglink_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            extractGangLinkData();
            TB.saveCalculations2(GangLink, Base.DBConnectionString, " where bussunit = '999' ", "GangLink");
            this.Cursor = Cursors.Arrow;
            MessageBox.Show("Ganglinking has been refreshed", "Information", MessageBoxButtons.OK);

        }

        private void extractGangLinkData()
        {
            string strSQL = "select distinct '" + BusinessLanguage.BussUnit + "' as BUSSUNIT,'" + BusinessLanguage.MiningType + "' as MININGTYPE,'" +
                             BusinessLanguage.BonusType + "' as BONUSTYPE,SECTION, PERIOD, WORKPLACE, 'XXX' as GANG,'0' as SAFETYIND, 'XXX' as GANGTYPE  " +
                            " from Survey " +
                            " where Section = '" + txtSelectedSection.Text.Trim() + "' AND PERIOD = '" + BusinessLanguage.Period + "'";


            DataTable tempDataTable = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

            lstNames = TB.loadDistinctValuesFromColumn(GangLink, "WORKPLACE");

            foreach (DataRow _row in tempDataTable.Rows)
            {
                if (string.IsNullOrEmpty(_row[0].ToString()))
                {
                }
                else
                {
                    if (lstNames.Contains(_row["WORKPLACE"].ToString().Trim()))
                    {
                    }
                    else
                    {
                        GangLink.Rows.Add(_row.ItemArray);
                    }
                }
            }
        }
        private List<string> extractGangWorkplaces(string Gang)
        {
            DataTable temp = new DataTable();
            List<string> lstTemp = new List<string>();


            if (GangLink.Rows.Count > 0)
            {
                IEnumerable<DataRow> query1 = from locks in GangLink.AsEnumerable()
                                              where locks.Field<string>("GANG").TrimEnd() == Gang
                                              select locks;

                try
                {
                    temp = query1.CopyToDataTable<DataRow>();
                }
                catch
                {

                }
            }

            if (temp.Rows.Count > 0)
            {
                lstTemp = TB.loadDistinctValuesFromColumn(temp, "WORKPLACE");
            }
            else
            {

            }

            return lstTemp;
        }
        private void btnDrillersRefresh_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            //Create a backuplist of lstGangs
            //lstBackup will contain all the gangs - STOPING, RIGGING, EQUIPPING etc.

            #region link gangs from ganglink
            if (GangLink.Rows.Count > 0)
            {
                DataTable temp = Drillers.Copy();
                temp.Clear();

                List<string> lstBackup = TB.loadDistinctValuesFromColumn(GangLink, "GANG");

                //Extract the gang numbers of the gangs already linked on DRILLERS
                List<string> lstTemp = new List<string>();
                lstTemp = TB.loadDistinctValuesFromColumn(Drillers, "Gang");

                List<string> lstDrillerWP = new List<string>();

                lstDrillerWP = TB.loadDistinctValuesFromColumn(Drillers, "Workplace");

                for (int i = 0; i <= lstBackup.Count - 1; i++)
                {
                    if (lstBackup[i].ToString().Trim().Length == 10 && lstBackup[i].ToString().Trim().Substring(4,2) == "ST")
                    {
                        //Import only STOPING gangs and not support gangs
                        if (lstTemp.Contains(lstBackup[i].ToString().Trim()))
                        {
                            //Extract the gang's workplace numbers
                            List<string> lstGangWorkplaces = new List<string>();
                            lstGangWorkplaces = extractGangWorkplaces(lstBackup[i].ToString().Trim());
                            foreach (string s in lstGangWorkplaces)
                            {
                                if (lstDrillerWP.Contains(s))
                                {
                                }
                                else
                                {
                                    if (lstBackup[i].ToString().Trim().Length == 10 && lstBackup[i].ToString().Trim().Substring(8, 1) == "R")
                                    {
                                        DataRow dr = temp.NewRow();

                                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
                                        dr["MININGTYPE"] = BusinessLanguage.MiningType;
                                        dr["BONUSTYPE"] = BusinessLanguage.BonusType;
                                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                                        dr["PERIOD"] = txtPeriod.Text.Trim();
                                        dr["WORKPLACE"] = s;
                                        dr["GANG"] = lstBackup[i].ToString().Trim();
                                        dr["EMPLOYEE_NO"] = "XXX";
                                        dr["DRILLERIND"] = "1";
                                        dr["DRILLERSHIFTS"] = "0";

                                        temp.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                        else
                        {
                            //Extract the gang's workplace numbers
                            List<string> lstGangWorkplaces = new List<string>();
                            lstGangWorkplaces = extractGangWorkplaces(lstBackup[i].ToString().Trim());
                            //Insert the new daygang into the drillers table for each workplace in the list

                            foreach (string s in lstGangWorkplaces)
                            {
                                if (lstBackup[i].ToString().Trim().Length == 10 && lstBackup[i].ToString().Trim().Substring(8, 1) == "R")
                                {
                                    DataRow dr = temp.NewRow();

                                    dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
                                    dr["MININGTYPE"] = BusinessLanguage.MiningType;
                                    dr["BONUSTYPE"] = BusinessLanguage.BonusType;
                                    dr["SECTION"] = txtSelectedSection.Text.Trim();
                                    dr["PERIOD"] = txtPeriod.Text.Trim();
                                    dr["WORKPLACE"] = s;
                                    dr["GANG"] = lstBackup[i].ToString().Trim();
                                    dr["EMPLOYEE_NO"] = "XXX";
                                    dr["DRILLERIND"] = "1";
                                    dr["DRILLERSHIFTS"] = "0";

                                    temp.Rows.Add(dr);
                                }
                            }
                        }
                }
                else
                {
                }
                }

                //Create a invalid delete that will execute in the savecalculation2 method.
                string strDelete = " where Bussunit = '999'";
                TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "DRILLERS");
                Application.DoEvents();
                evaluateDrillers();
            }

            else
            {
                MessageBox.Show("No records in Bonusshifts.  Please re-import clocked shifts.", "Information", MessageBoxButtons.OK);
            }
            #endregion

            //#region link the support gangs
            ////Read the supportlink table and insert the tramming and haulage and rigging gangs
            //if (SupportLink.Rows.Count > 0)
            //{
            //    DataTable temp = Drillers.Copy();
            //    temp.Clear();

            //    List<string> lstBackup = TB.loadDistinctValuesFromColumn(Labour, "GANG");

            //    //Extract the gang numbers of the gangs already linked on DRILLERS
            //    List<string> lstTemp = new List<string>();
            //    lstTemp = TB.loadDistinctValuesFromColumn(Drillers, "Gang");

            //    for (int i = 0; i <= lstBackup.Count - 1; i++)
            //    {
            //        if (lstTemp.Contains(lstBackup[i].ToString()))
            //        {

            //        }
            //        else
            //        {
            //            //Extract the gang's workplace numbers from ganglink
            //            List<string> lstGangWorkplaces = new List<string>();
            //            lstGangWorkplaces = extractGangWorkplaces(lstBackup[i].ToString().Trim());
            //            if (lstGangWorkplaces.Count == 0)
            //            {
            //                //It is an unlink gang or a support gang because workplaces were not found on ganglink
            //                //Try to extract the gangs dayshift gang from the supportlink table
            //                //If found, add the gang, else ignore the gang
            //                List<string> lstDayShiftList = new List<string>();
            //                lstDayShiftList = extractDayShiftGangs(lstBackup[i].ToString().Trim());

            //                foreach (string ss in lstDayShiftList)
            //                {
            //                    List<string> lstGangWorkplaces2 = new List<string>();
            //                    lstGangWorkplaces2 = extractGangWorkplaces(ss);

            //                    foreach (string wp in lstGangWorkplaces2)
            //                    {
            //                        DataRow dr = temp.NewRow();

            //                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
            //                        dr["MININGTYPE"] = BusinessLanguage.MiningType;
            //                        dr["BONUSTYPE"] = BusinessLanguage.BonusType;
            //                        dr["SECTION"] = txtSelectedSection.Text.Trim();
            //                        dr["PERIOD"] = txtPeriod.Text.Trim();
            //                        dr["WORKPLACE"] = wp;
            //                        dr["GANG"] = lstBackup[i].ToString().Trim();
            //                        dr["EMPLOYEE_NO"] = "XXX";
            //                        dr["DRILLERIND"] = "1";
            //                        dr["DRILLERSHIFTS"] = "0";

            //                        temp.Rows.Add(dr);
            //                    }
            //                }
            //            }

            //            //Create a invalid delete that will execute in the savecalculation2 method.

            //        }
            //    }

            //    string strDelete = " where Bussunit = '999'";
            //    TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "DRILLERS");
            //    evaluateDrillers();


            //}

            //#endregion

            this.Cursor = Cursors.Arrow;
        }


        private void lstGangs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstGangs.SelectedIndex >= 0)
            {
                cboGangLinkGang.Text = lstGangs.SelectedItem.ToString();
            }
        }

        private void grdDrillers_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {

            }
            else
            {
                if (grdDrillers.Columns.Count > 0 && grdDrillers.Columns.Count > 2)
                {
                    txtAutoDGang.Text = grdDrillers["GANG", e.RowIndex].Value.ToString().Trim();
                    txtAutoDWorkplace.Text = grdDrillers["WORKPLACE", e.RowIndex].Value.ToString().Trim();
                    txtAutoDriller.Text = grdDrillers["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                    cboAutoDrillerDrilInd.Text = grdDrillers["DRILLERIND", e.RowIndex].Value.ToString().Trim();
                    txtAutoDrilShifts.Text = grdDrillers["DRILLERSHIFTS", e.RowIndex].Value.ToString().Trim();

                    IEnumerable<DataRow> query1 = from locks in Survey.AsEnumerable()
                                                  where locks.Field<string>("WORKPLACE").TrimEnd() == grdDrillers["WORKPLACE", e.RowIndex].Value.ToString().Trim()
                                                  select locks;

                    try
                    {
                        DataTable temp = query1.CopyToDataTable<DataRow>();
                        txtAutoDWorkplace.Text = grdDrillers["WORKPLACE", e.RowIndex].Value.ToString().Trim() + " - " +
                                                 temp.Rows[0]["DESCRIPTION"].ToString().Trim();
                    }
                    catch
                    {
                        txtAutoDWorkplace.Text = grdDrillers["WORKPLACE", e.RowIndex].Value.ToString().Trim();
                    }


                    query1 = from locks in Labour.AsEnumerable()
                             where locks.Field<string>("GANG").TrimEnd() == grdDrillers["GANG", e.RowIndex].Value.ToString().Trim()
                             select locks;

                    try
                    {
                        DataTable temp = query1.CopyToDataTable<DataRow>();
                        lstDrillers.Items.Clear();
                        foreach (DataRow row in temp.Rows)
                        {

                            lstDrillers.Items.Add(row["EMPLOYEE_NO"].ToString().Trim() + "-" + row["Employee_name"].ToString().Trim());

                        }
                    }
                    catch
                    {
                        MessageBox.Show("This datatable contains no values");
                    }

                    #region Trigger output
                    //load the CURRENT values into dictionaries before the update 
                    // that was loaded in tabInfo_SelectedIndexChanged
                    dictPrimaryKeyValues.Clear();
                    dictGridValues.Clear();

                    foreach (string s in lstPrimaryKeyColumns)
                    {
                        if (e.RowIndex < 0)
                        {
                        }
                        else
                        {
                            dictPrimaryKeyValues.Add(s, grdDrillers[s, e.RowIndex].Value.ToString().Trim());
                        }
                    }

                    foreach (string s in lstTableColumns)
                    {
                        if (e.RowIndex < 0)
                        {
                        }
                        else
                        {
                            dictGridValues.Add(s, grdDrillers[s, e.RowIndex].Value.ToString().Trim());
                        }
                    }
                    #endregion
                }
                else
                {
                    txtAutoDriller.Text = grdDrillers["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                }
            }

            

        }

        private void cboEmplPenEmployeeNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Search for the coyno in the Labour datatable
            DataTable temp = new DataTable();
            if (Clocked.Rows.Count > 0)
            {
                IEnumerable<DataRow> query1 = from locks in Clocked.AsEnumerable()
                                              where locks.Field<string>("EMPLOYEE_NO").TrimEnd() == cboEmplPenEmployeeNo.Text.Trim()
                                              select locks;


                temp = query1.CopyToDataTable<DataRow>();
            }

            if (temp.Rows.Count > 0)
            {
                cboEmplPenEmployeeName.Text = temp.Rows[0]["Employee_Name"].ToString().Trim();
            }
            else
            {
                cboEmplPenEmployeeName.Text = "XXX";
            }
        }

        private void cboEmplPenEmployeeName_SelectedIndexChanged(object sender, EventArgs e)
        {

            #region Get employee no of shifts
            //Search for the coyno in the Labour datatable
            DataTable temp = new DataTable();
            if (Clocked.Rows.Count > 0)
            {
                IEnumerable<DataRow> query1 = from locks in Labour.AsEnumerable()
                                              where locks.Field<string>("EMPLOYEE_NAME").TrimEnd() == cboEmplPenEmployeeName.SelectedItem.ToString().Trim()
                                              select locks;


                temp = query1.CopyToDataTable<DataRow>();
            }

            if (temp.Rows.Count > 0)
            {
                cboEmplPenEmployeeNo.Text = temp.Rows[0]["EMPLOYEE_NO"].ToString().Trim();

            }
            else
            {
                cboEmplPenEmployeeNo.Text = "-";
            }

            #endregion

        }

        private void grdFactors_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {

            }
            else
            {
                txtVarValue.Text = grdFactors["VarValue", e.RowIndex].Value.ToString().Trim();
                cboVarName.Text = grdFactors["VarName", e.RowIndex].Value.ToString().Trim();

                btnUpdate.Enabled = true;
                btnDeleteRow.Enabled = false;
                btnInsertRow.Enabled = false;
            }

            #region Trigger output
            //load the CURRENT values into dictionaries before the update 
            dictPrimaryKeyValues.Clear();
            dictGridValues.Clear();

            foreach (string s in lstPrimaryKeyColumns)
            {
                dictPrimaryKeyValues.Add(s, grdFactors[s, e.RowIndex].Value.ToString().Trim());
            }

            foreach (string s in lstTableColumns)
            {
                dictGridValues.Add(s, grdFactors[s, e.RowIndex].Value.ToString().Trim());
            }
            #endregion

            Cursor.Current = Cursors.Arrow;
        }

        private void lstStopeReports_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstStopeReports.SelectedItem.ToString().Trim() == "Per Gang")
            {
                this.Cursor = Cursors.WaitCursor;
                MetaReportRuntime.App mm = new MetaReportRuntime.App();
                mm.Init(strMetaReportCode);
                mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
                mm.StartReport("STPTM4000Team");
                this.Cursor = Cursors.Arrow;
            }
            else
            {
                this.Cursor = Cursors.WaitCursor;
                MetaReportRuntime.App mm = new MetaReportRuntime.App();
                mm.Init(strMetaReportCode);
                mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
                mm.StartReport("STPTM4000TeamPERworkplace");
                this.Cursor = Cursors.Arrow;
            }

            lstStopeReports.Visible = false;
        }

        private void btnSupportPrint_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTM4000Support");
            this.Cursor = Cursors.Arrow;
        }

        private void printCostsheet_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTM4000CS");
            this.Cursor = Cursors.Arrow;
        }

        private void printAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTMTeamAuthorization");
            this.Cursor = Cursors.Arrow;      
        }

        private void btnOtherTeamAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTM4000TrammingAuthorization");
            this.Cursor = Cursors.Arrow;    
        }

        private void btnRiggingTeamAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTM4000RiggingAuthorization");
            this.Cursor = Cursors.Arrow;
        }

        private void btnEquippingTeamAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTM4000EquippingAuthorization");
            this.Cursor = Cursors.Arrow;
        }

        private void btnHaulageTeamAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTM4000HaulageAuthorization");
            this.Cursor = Cursors.Arrow;
        }
        private void btnCostsheetAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTM_CAS");
            this.Cursor = Cursors.Arrow;
        }

        private void btnDrillerPrint_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTM4000Driller");
            this.Cursor = Cursors.Arrow;

        }

        private void btnDrillerAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Tshepong\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPDrillerAUTH");
            this.Cursor = Cursors.Arrow;
        }

        private void label4_Click(object sender, EventArgs e)
        {
            TB.InsertData(Base.AnalysisConnectionString,"Delete from process where period = '" + BusinessLanguage.Period + 
                                                        "' and section = '" + txtSelectedSection.Text.Trim() + "'");
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void cboColumnNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> lstColumnValues = lstNames = TB.loadDistinctValuesFromColumn(newDataTable, cboColumnNames.SelectedItem.ToString());

            foreach (string s in lstColumnValues)
            {
                cboColumnValues.Items.Add(s.Trim());
            }

        }

        private void cboColumnValues_SelectedIndexChanged(object sender, EventArgs e)
        {

            IEnumerable<DataRow> query1 = from locks in newDataTable.AsEnumerable()
                                          where locks.Field<string>(cboColumnNames.SelectedItem.ToString()).TrimEnd() == cboColumnValues.SelectedItem.ToString()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            grdActiveSheet.DataSource = temp;

            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, TB.TBName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }
        }

        private void btnShowEmpl_Click(object sender, EventArgs e)
        {

            DataTable temp = new DataTable();
            IEnumerable<DataRow> query1 = from locks in Drillers.AsEnumerable()
                                          where locks.Field<string>("Employee_no").TrimEnd() == txtAutoDriller.Text.Trim()
                                          select locks;

            try
            {
                temp = query1.CopyToDataTable<DataRow>();
            }
            catch
            {

            }

            grdDrillers.DataSource = temp;
        }

        private void btnShowAll_Click(object sender, EventArgs e)
        {
            evaluateDrillers();
        }

        private void btnDrilOtherEmp_Click(object sender, EventArgs e)
        {
            createOtherEmpList();
        }

        private void createOtherEmpList()
        {
            List<string> lstBackup = TB.loadDistinctValuesFromColumn(Labour, "EMPLOYEE_NO");

            DataTable Temp = new DataTable();
            lstDrillers.Items.Clear();
            Temp = TB.createDataTableWithAdapter(Base.DBConnectionString,
                       "SELECT * FROM BONUSSHIFTS");

            foreach (DataRow row in Temp.Rows)
            {
                lstDrillers.Items.Add(row["EMPLOYEE_NO"].ToString().Trim() + "-" + row["Employee_name"].ToString().Trim());
            }

            lstDrillers.Refresh();
        }

    }
}
