
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Net;
using System.Data.Common;
using System.Web;
using System.Net.Mail;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing.Drawing2D;
using SAPbobsCOM;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;

namespace KSM_CostingModule
{
    class clsMain
    {
        #region Variables
        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.Company oCompany = null;

        string[] supportedFormats = new string[] { "dd/MM/yy", "dd/MM/yyyy", "dd-MM-yy", "dd-MM-yyyy", "MM/dd/yyyy", "MM/dd/yy", "ddMMMyyyy", "ddMMyyyy", "yyyyMMdd" };

        string psPath = "", sPath = "";
        SAPbouiCOM.MenuItem oMenuItem;
        SAPbouiCOM.Menus oMenus;
        SAPbouiCOM.MenuCreationParams oCreationPackage = null;
        public static SAPbobsCOM.Recordset oRec, oRecConfig, oRec1, oRecdata, oRecordset, oRecordSet, oRecM;
        public SAPbouiCOM.EventFilters oFilters;
        public SAPbouiCOM.EventFilter oFilter;

        public SAPbouiCOM.Form oForm;
        public static SAPbouiCOM.ComboBox oCombo;
        public SAPbouiCOM.Item oItem;
        public static string DBName, ServerName, UserName;
        public static string ProjectName = "Costing Module";
        public static string ProjectCode = "CM";
        private int IntCode;
        private SAPbobsCOM.UserTablesMD oUserTablesMD = null;
        private SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
        private SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
        public static bool CloseFlg, BtnFlag = false, DBSetupFlg = false;

        public static bool Success = false;

        CostingModule CostingModule = new CostingModule();

        #endregion

        #region Main
        public clsMain()
        {
            try
            {
                SetApplication();
                Setconnection();
                DBName = SBO_Application.Company.DatabaseName; //Database Name            
                ServerName = SBO_Application.Company.ServerName;//Server Name
                UserName = SBO_Application.Company.UserName;
                sPath = System.Windows.Forms.Application.StartupPath.ToString();


                SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
                SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);

                //if (DataBaseCreate() == true)
                //{
                //    //clsMain.SBO_Application.SetStatusBarMessage("Database Created successfully ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                //}
                AddMenu();
                SBO_Application.MessageBox("Add on " + ProjectName + " Connected !", 1, "Ok", "", "");
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }

        }
        #endregion

        #region Connection
        public void Setconnection()
        {
            try
            {
                string sCookie;
                string sConnectionContext;

                // First initialize the Company object
                oCompany = new SAPbobsCOM.Company();
                if (oCompany.Connected == true)
                    return;

                sCookie = oCompany.GetContextCookie();
                sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
                oCompany.SetSboLoginContext(sConnectionContext);
                oCompany.Connect();
            }
            catch (Exception ex)
            {
                oCompany = SBO_Application.Company.GetDICompany();
                //SBO_Application.MessageBox(ex.ToString().ToString(), 1, "ok", "", "");s
            }
        }
        private void SetFilters()
        {
            oFilters = new SAPbouiCOM.EventFilters();

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
            SBO_Application.SetFilter(oFilters);
        }
        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            //sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            try
            {
                // If there's no active application the connection will fail
                SboGuiApi.Connect(sConnectionString);
            }
            catch
            { //  Connection failed
                System.Windows.Forms.MessageBox.Show("No SAP Business One Application");
                System.Environment.Exit(0);
            }
            // get an initialized application object
            SBO_Application = SboGuiApi.GetApplication(-1);
        }

        #endregion

        #region Menu
        private void AddMenu()
        {
            try
            {
                oMenus = null;
                oMenuItem = null;
                oCreationPackage = null;
                oMenus = SBO_Application.Menus;
                oCreationPackage = ((SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                oMenuItem = oMenus.Item("43520");
                //oMenuItem = oMenus.Item("2048");
                if (!SBO_Application.Menus.Exists("mCMf"))
                {                    
                    AddMenu_Items(16, "", SAPbouiCOM.BoMenuType.mt_POPUP, "Costing Module", "mCMf", oMenus);
                    oMenuItem = SBO_Application.Menus.Item("mCMf");
                    oMenus = oMenuItem.SubMenus;
                    AddMenu_Items(0, "", SAPbouiCOM.BoMenuType.mt_STRING, "Costing Module", "mCM", oMenus);
                    //AddMenu_Items(1, "", SAPbouiCOM.BoMenuType.mt_STRING, "Quotation WhatsApp", "mQWh", oMenus);
                    //AddMenu_Items(2, "", SAPbouiCOM.BoMenuType.mt_STRING, "Opportunity WhatsApp", "mOWh", oMenus);
                    //AddMenu_Items(3, "", SAPbouiCOM.BoMenuType.mt_STRING, "BP Wise SMS", "mSMS", oMenus);
                }
            }
            catch { }
        }

        private void AddMenu_Items(int Position, string ImageName, SAPbouiCOM.BoMenuType Type, string MenuLabel, string MenuId, SAPbouiCOM.Menus oMenus)//, ref string ImageName,ref SAPbouiCOM.BoMenuType Type, ref string MenuLabel,ref string MenuId, ref SAPbouiCOM.Menus oMenus
        {
            try
            {
                oCreationPackage.Type = Type;
                oCreationPackage.String = MenuLabel;
                oCreationPackage.UniqueID = MenuId;
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = Position;
                if (SBO_Application.ClientType == SAPbouiCOM.BoClientType.ct_Desktop)
                {
                    oCreationPackage.Image = sPath + "\\" + ImageName;
                }
                if (!oMenus.Exists(MenuId))
                {
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void RemoveMenu()
        {
            oMenus = SBO_Application.Menus;
            oMenuItem = SBO_Application.Menus.Item("frmDB");
            RemoveMenu_Item(oMenus, "frmDB");

            oMenuItem = SBO_Application.Menus.Item("43520");
            RemoveMenu_Item(oMenus, "frmProject");

            oMenus = oMenuItem.SubMenus;
            oMenuItem = SBO_Application.Menus.Item("frmProject");

        }

        private void RemoveMenu_Item(SAPbouiCOM.Menus oMenus, string MenuId)
        {
            if (oMenus.Exists(MenuId) == true)
            {
                oMenuItem = SBO_Application.Menus.Item(MenuId);
                oMenus.Remove(oMenuItem);
            }
        }
        #endregion

        #region ItemEvents
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                GC.Collect();
                switch (FormUID)
                {
                    case "frmCostModule":
                        CostingModule.itemevent(ref pVal, ref BubbleEvent, FormUID);
                        break;
                }

            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.ToString(), BoMessageTime.bmt_Medium, true);
            }
        }
        #endregion

        #region MenuEvents
        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction == true)
                {
                    #region Delete Row
                    if (pVal.MenuUID == "1293")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmLPGAlloc":
                                CostingModule.MenuEvent(ref pVal, oForm.UniqueID, "DEL");
                                BubbleEvent = false;
                                break;
                        }
                    }
                    #endregion

                    #region Add Mode

                    if (pVal.MenuUID == "1282" || pVal.MenuUID == "1281")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        //if (pVal.MenuUID == "1281")
                        //{
                        //    oForm.Items.Item("btnSM133").Enabled = false;
                        //    BubbleEvent = false;
                        //}
                        //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.MenuUID == "1282")
                        //{
                        //    oForm.Items.Item("btnSM133").Enabled = false;
                        //    BubbleEvent = false;
                        //}
                    }
                    #endregion
                }
                if (pVal.BeforeAction == false)
                {
                    #region Add Mode
                    if (pVal.MenuUID == "1282")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmLPGAlloc":
                                CostingModule.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                                break;
                        }

                    }
                    #endregion

                    #region Find Mode
                    if (pVal.MenuUID == "1281")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmLPGAlloc":
                                BubbleEvent = false;
                                CostingModule.MenuEvent(ref pVal, oForm.UniqueID, "Find");
                                break;
                        }
                    }
                    #endregion

                    #region Navigation
                    if (pVal.MenuUID == "1288" || pVal.MenuUID == "1289" || pVal.MenuUID == "1290" || pVal.MenuUID == "1291")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmJWIssue":
                                //JI.MenuEvent(ref pVal, oForm.UniqueID, "NAV");
                                break;
                        }
                    }
                    #endregion

                    #region Add Row
                    if (pVal.MenuUID == "1292")
                    {
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        switch (oForm.UniqueID)
                        {
                            case "frmLPGAlloc":
                                CostingModule.MenuEvent(ref pVal, oForm.UniqueID, "AddR");
                                break;

                        }
                    }
                    #endregion

                    #region Costing Module

                    if (pVal.MenuUID == "mCM")
                    {
                        for (int i = 1; i < SBO_Application.Forms.Count; i++)
                        {
                            if (SBO_Application.Forms.Item(i).UniqueID == "frmCostModule")
                            {
                                SBO_Application.Forms.Item(i).Select();
                                return;
                            }
                        }
                        LoadFromXML("frmCostModule");
                        oForm = clsMain.SBO_Application.Forms.ActiveForm;
                        //oForm.DataBrowser.BrowseBy = "tCode";
                        CostingModule.MenuEvent(ref pVal, oForm.UniqueID, "Add");
                    }

                    #endregion

                }
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        #endregion

        #region Right Click Event
        public void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (eventInfo.FormUID)
                {
                    case "frmLPGAlloc":
                        CostingModule.RightClickEvent(ref eventInfo, ref BubbleEvent, eventInfo.FormUID);
                        break;

                }

            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        #endregion

        #region FormDataEvent
        public void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                GC.Collect();
                switch (BusinessObjectInfo.FormTypeEx)
                {
                    case "frmLPGAlloc":
                        CostingModule.FormDataEvent(ref BusinessObjectInfo, ref BubbleEvent);
                        break;

                }
            }
            catch { }
        }
        #endregion

        #region SAPEvents
        public static void LoadFromXML(string FileName)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();
            // load the content of the XML File
            string sPath = null;
            FileName = FileName + ".xml";
            sPath = System.Windows.Forms.Application.StartupPath.ToString();
            oXmlDoc.Load(sPath + @"\XML\" + FileName);
            // load the form to the SBO application in one batch
            string tmpStr;
            tmpStr = oXmlDoc.InnerXml;
            SBO_Application.LoadBatchActions(ref tmpStr);
            sPath = SBO_Application.GetLastBatchResults();
        }
        public static int GetNextDocNum(ref SAPbouiCOM.EditText oEdit, ref string TableName)
        {
            try
            {
                oEdit.Value = string.Empty;
                if (TableName.Trim() != string.Empty)
                {
                    SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRec.DoQuery("Select Max(\"DocEntry\") From \"" + TableName + "\"");
                    int MaxCode = Convert.ToInt32(oRec.Fields.Item(0).Value);
                    if (MaxCode == 0)
                    {
                        oEdit.Value = "1";
                    }
                    else
                    {
                        if (!oRec.EoF)
                        {
                            oEdit.Value = Convert.ToString(MaxCode + 1);
                        }
                        else
                        {
                            oEdit.Value = "1";
                        }
                    }
                    oRec = null;
                }
                else
                {
                    SBO_Application.StatusBar.SetText("Error on GetNextDocNum : Invalid TableName ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
                return 0;
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("Error on GetNextDocNum :" + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return -1;
            }
        }
        private void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition || EventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown || EventType == SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged)
            {
                //RemoveMenu();
                UnHandleEvents();
                SBO_Application = null;
                oCompany = null;
                foreach (var process in Process.GetProcessesByName("WhatsApp"))
                {
                    process.Kill();
                }
                System.Windows.Forms.Application.Exit();
            }
        }
        private void UnHandleEvents()
        {
            try
            {
                SBO_Application.ItemEvent -= new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SBO_Application.MenuEvent -= new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                SBO_Application.AppEvent -= new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SBO_Application.FormDataEvent -= new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
                SBO_Application.RightClickEvent -= new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
            }
            catch (Exception ex)
            {
            }
        }
        #endregion

        #region Other Methods

        public static void SetCode(String FormId, String OBJ, String Type)
        {
            try
            {
                SAPbouiCOM.Form oForm;
                SAPbouiCOM.EditText oEdit;
                string Table = null;
                DateTime now = DateTime.Now;
                oForm = SBO_Application.Forms.ActiveForm;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oEdit = oForm.Items.Item("tCode").Specific;
                Table = "@" + OBJ;

                GetNextDocNum(ref oEdit, ref Table);
                if (Type == "D")
                {
                    oRecordset = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordset.DoQuery("Select * From OUDO Where \"MngSeries\" = 'N' and \"Code\" = '" + OBJ + "' ");
                    oForm.Items.Item("tDocDate").Specific.value = DateTime.Now.ToString("yyyyMMdd");
                    if (oRecordset.RecordCount > 0)
                    {
                        oForm.Items.Item("tDocNum").Specific.value = oEdit.Value;
                    }
                    else
                    {
                        //oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber("DocEntry", OBJ);
                        //Event.Series.SeriesCombo(OBJ, "cSer");
                        oForm.Items.Item("tDocNum").Specific.value = oForm.BusinessObject.GetNextSerialNumber(oForm.Items.Item("cSer").Specific.value, OBJ);
                    }

                    //oForm.Items.Item("tCreator").Specific.value = clsMain.UserName;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Set Code :" + ex, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region Create DataBase

        private bool DataBaseCreate()
        {
            try
            {
                //AddMenu();
                oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery("Select * From \"@LICENSEMST\" Where \"Name\" = '" + clsMain.ProjectCode + "'");
                if (oRecordset.RecordCount == 0)
                {
                    DBSetupFlg = true;
                }
                else if (oRecordset.Fields.Item("U_DBCreate").Value == "Y")
                {
                    DBSetupFlg = true;
                }
            }
            catch
            {
                DBSetupFlg = true;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                oRecordset = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            if (DBSetupFlg == true)
            {
                try
                {
                    clsMain.SBO_Application.SetStatusBarMessage("Database structure creation in progress...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    CreateDataBase();
                    oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordset.DoQuery("Select * From \"@LICENSEMST\" Where \"Name\" = '" + clsMain.ProjectCode + "'");
                    if (oRecordset.RecordCount > 0)
                    {
                        oRecordset.DoQuery("Update \"@LICENSEMST\" Set \"U_DBCreate\" = 'N' Where \"Name\" = '" + clsMain.ProjectCode + "'");
                    }
                }
                catch (Exception ex)
                {
                    SBO_Application.SetStatusBarMessage(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                    oRecordset = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            return true;
        }

        private bool CreateDataBase()
        {
            try
            {
                #region Licence field creation

                CreateTable("LICENSEMST", "License Master", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                CreateFields("@LICENSEMST");


                #endregion

                #region SystemTable

                CreateFields("SystemTable");

                #endregion


                #region LPG_Allocation
                CreateTable("LPG_ALLOC", "LPG Allocations", SAPbobsCOM.BoUTBTableType.bott_Document);
                CreateTable("LPG_ALLOC1", "LPG Allocations Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                CreateTable("LPG_AD_QT", "LPG Additional Quota", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                CreateTable("LPG_ALLOC2", "LPG Quota Transfer", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);


                CreateFields("@LPG_ALLOC");

                string[] FindColoums = new string[2];
                FindColoums[0] = "DocNum";
                FindColoums[1] = "U_DocDate";
           

                string[] ChildTable = new string[3];
                ChildTable[0] = "LPG_ALLOC1";
                ChildTable[1] = "LPG_AD_QT";
                ChildTable[2] = "LPG_ALLOC2";

                CreateUserObject("LPG_ALLOC", "LPG Allocation Month Wise", "LPG_ALLOC", ChildTable, FindColoums, SAPbobsCOM.BoUDOObjType.boud_Document, "N");
                //BoUDOObjType.boud_Document.ToString()
                #endregion

                #region AutoFMS
                AutoFMS();
                #endregion

                return true;
            }
            catch (Exception ex)
            {
                // SBO_Application.StatusBar.SetText(ex.ToString().ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        private bool CreateFields(string TableName)
        {
            try
            {
                # region Licence Table
                if (TableName == "@LICENSEMST")
                {
                    FieldDetails("@LICENSEMST", "Module", "Module", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", "");
                    FieldDetails("@LICENSEMST", "DBCreate", "Database Creation", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "NY", "N");
                }
                # endregion

                #region SystemTable
                if (TableName == "SystemTable")
                {
                    FieldDetails("OOCR", "LPG_Alloc", "Route for LPG_Allocation(Y/N)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "NY", "N");
                }
                #endregion                

                # region LPG_Allocation
                if (TableName == "@LPG_ALLOC")
                {
                    FieldDetails("@LPG_ALLOC", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC", "Month", "Month", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 4, "", "");
                    FieldDetails("@LPG_ALLOC", "Browse", "Browse", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 254, "", "");
                    FieldDetails("@LPG_ALLOC", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "AI", "A");

                    FieldDetails("@LPG_ALLOC1", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", "");
                    FieldDetails("@LPG_ALLOC1", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", "");
                    FieldDetails("@LPG_ALLOC1", "BAR_BRG", "BAR_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "HLD_BRG", "HLD_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "DPR_BRG", "DPR_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "PRD_BRG", "PRD_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "BAR_BWA", "BAR_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "PRD_BWA", "PRD_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "MTH_BWA", "MTH_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "KAR_NPJ", "KAR_NPJ", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "MTH_NPJ", "MTH_NPJ", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "HLD_BWA", "HLD_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "MTH_BRG", "MTH_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "KAR_BWA", "KAR_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "MTH_DHI", "MTH_DHI", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "KAR_DHI", "KAR_DHI", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "BAR_BIR", "BAR_BIR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "HLD_BIR", "HLD_BIR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "DPR_BIR", "DPR_BIR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "PRD_BIR", "PRD_BIR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "DPR_BWA", "DPR_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "BAR_CHR", "BAR_CHR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "DPR_CHR", "DPR_CHR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "PRD_CHR", "PRD_CHR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC1", "HLD_CHR", "HLD_CHR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");

                    FieldDetails("@LPG_AD_QT", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", "");
                    FieldDetails("@LPG_AD_QT", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", "");
                    FieldDetails("@LPG_AD_QT", "BAR_BRG", "BAR_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "HLD_BRG", "HLD_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "DPR_BRG", "DPR_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "PRD_BRG", "PRD_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "BAR_BWA", "BAR_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "PRD_BWA", "PRD_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "MTH_BWA", "MTH_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "KAR_NPJ", "KAR_NPJ", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "MTH_NPJ", "MTH_NPJ", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "HLD_BWA", "HLD_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "MTH_BRG", "MTH_BRG", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "KAR_BWA", "KAR_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "MTH_DHI", "MTH_DHI", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "KAR_DHI", "KAR_DHI", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "BAR_BIR", "BAR_BIR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "HLD_BIR", "HLD_BIR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "DPR_BIR", "DPR_BIR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "PRD_BIR", "PRD_BIR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "DPR_BWA", "DPR_BWA", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "BAR_CHR", "BAR_CHR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "DPR_CHR", "DPR_CHR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "PRD_CHR", "PRD_CHR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "HLD_CHR", "HLD_CHR", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_AD_QT", "Date", "Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");

                    FieldDetails("@LPG_ALLOC2", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", "");
                    FieldDetails("@LPG_ALLOC2", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", "");
                    FieldDetails("@LPG_ALLOC2", "FrmRoute", "From Route", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");                   
                    FieldDetails("@LPG_ALLOC2", "ToRoute", "To Route", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC2", "TrnsQty", "Transfer Quantity", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");                   
                    FieldDetails("@LPG_ALLOC2", "TrnsDate", "Transfer Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "");
                    FieldDetails("@LPG_ALLOC2", "Commited", "Is Commited", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "NY", "N");


                }
                #endregion

                return true;
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Create Field: " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }

        private bool CreateTable(string TableName, string TableDesc, SAPbobsCOM.BoUTBTableType TableType)
        {
            try
            {
                int errCode;
                string ErrMsg = null;
                //oUserTablesMD = null;
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                //oUserTablesMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (oUserTablesMD == null)
                {
                    oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
                }
                if (oUserTablesMD.GetByKey(TableName) == true)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return true;
                }
                oUserTablesMD.TableName = TableName;
                oUserTablesMD.TableDescription = TableDesc;
                oUserTablesMD.TableType = TableType;
                long err = oUserTablesMD.Add();
                if (err != 0)
                {
                    clsMain.oCompany.GetLastError(out errCode, out ErrMsg);
                }
                if (err == 0)
                {
                    clsMain.SBO_Application.StatusBar.SetText("Table Created : " + TableDesc + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    //CreateFields(TableName);
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool FieldDetails(string TableName, string FieldName, string FieldDesc, SAPbobsCOM.BoFieldTypes FieldType, SAPbobsCOM.BoFldSubTypes FieldSubType, int FieldSize, string ValidValues, string DefaultVal)
        {
            if (FieldExist(TableName, FieldName) == false)
            {
                string ErrMsg;
                int errCode;
                int IRetCode;
                oUserFieldsMD = null;
                try
                {
                    GC.Collect();
                    oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
                    oUserFieldsMD.TableName = TableName;
                    oUserFieldsMD.Name = FieldName;
                    oUserFieldsMD.Description = FieldDesc;

                    oUserFieldsMD.Type = FieldType;
                    oUserFieldsMD.SubType = FieldSubType;
                    oUserFieldsMD.EditSize = FieldSize;

                    if (ValidValues != "")
                    {
                        switch (ValidValues)
                        {
                            case "NY":
                                oUserFieldsMD.ValidValues.Value = "Y";
                                oUserFieldsMD.ValidValues.Description = "Yes";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "N";
                                oUserFieldsMD.ValidValues.Description = "No";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "N";
                                break;

                            case "AI":
                                oUserFieldsMD.ValidValues.Value = "A";
                                oUserFieldsMD.ValidValues.Description = "Active";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "I";
                                oUserFieldsMD.ValidValues.Description = "Inactive";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "A";
                                break;

                            case "DocType":
                                oUserFieldsMD.ValidValues.Value = "13";
                                oUserFieldsMD.ValidValues.Description = "A/R Invoice";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "17";
                                oUserFieldsMD.ValidValues.Description = "Sales Order";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "23";
                                oUserFieldsMD.ValidValues.Description = "Sales Quotation";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "203";
                                oUserFieldsMD.ValidValues.Description = "A/R Downpayment Request";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "-1";
                                oUserFieldsMD.ValidValues.Description = "WhatsApp Bulk";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "13";
                                break;

                            case "Currency":
                                oUserFieldsMD.ValidValues.Value = "INR";
                                oUserFieldsMD.ValidValues.Description = "INR";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "NOT_INR";
                                oUserFieldsMD.ValidValues.Description = "Other Than INR";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "INR";
                                break;

                        }
                    }
                    if (DefaultVal != "")
                    {
                        oUserFieldsMD.DefaultValue = DefaultVal;
                    }
                    // Add the field to the table
                    IRetCode = oUserFieldsMD.Add();
                    if (IRetCode != 0)
                    {
                        if (IRetCode == -2035 || IRetCode == -1120)
                        {
                            clsMain.oCompany.GetLastError(out errCode, out ErrMsg);
                            return false;
                        }
                        else
                        {
                            clsMain.oCompany.GetLastError(out errCode, out ErrMsg);
                            clsMain.SBO_Application.SetStatusBarMessage("Error : " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                    else
                    {
                        clsMain.SBO_Application.SetStatusBarMessage("Field Created in : " + TableName + " As : " + FieldDesc, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    oUserFieldsMD = null;
                    ErrMsg = null; errCode = 0; IRetCode = 0;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            return true;
        }

        private bool CreateUserObject(string CodeID, string Name, string TableName, string[] ChildTableName, string[] FindColoums, SAPbobsCOM.BoUDOObjType ObjectType, string ManageSeries) //used for registration of user defined table
        {
            try
            {
                int lRetCode = 0, code = 0;
                string sErrMsg = null;
                oUserObjectMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (oUserObjectMD == null)
                    oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));
                if (oUserObjectMD.GetByKey(CodeID) == true)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                    oUserObjectMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return true;
                }
                oUserObjectMD.Code = CodeID;
                oUserObjectMD.Name = Name;
                oUserObjectMD.TableName = TableName;

                oUserObjectMD.ObjectType = ObjectType;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                if (ManageSeries == "Y")
                {
                    oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                else
                {
                    oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                }

                if (ChildTableName != null)
                {
                    for (int i = 0; i <= ChildTableName.Length - 1; i++)
                    {
                        if (ChildTableName[i] != null)
                        {
                            if (ChildTableName[i].Trim() != string.Empty)
                            {
                                oUserObjectMD.ChildTables.TableName = ChildTableName[i];
                                oUserObjectMD.ChildTables.Add();
                            }
                        }
                    }
                }
                if (FindColoums != null)
                {
                    for (int i = 0; i <= FindColoums.Length - 1; i++)
                    {
                        if (FindColoums[i] != null)
                        {
                            if (FindColoums[i].Trim() != string.Empty)
                            {
                                oUserObjectMD.FindColumns.ColumnAlias = FindColoums[i];
                                oUserObjectMD.FindColumns.Add();
                            }
                        }
                    }
                }
                // check for errors in the process
                lRetCode = oUserObjectMD.Add();

                if (lRetCode != 0)
                    if (lRetCode == -1)
                    { clsMain.oCompany.GetLastError(out lRetCode, out sErrMsg); }
                    else
                    { clsMain.oCompany.GetLastError(out lRetCode, out sErrMsg); }
                else
                {
                    SAPbobsCOM.Recordset oRecordset, oRec;
                    oRecordset = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    if (TableName != "LICAUTHMST")
                    {
                        oRecordset.DoQuery("Select \"Code\" from \"@OBJECTMST\" Where \"U_ObjCode\" = '" + TableName + "' ");
                        if (oRecordset.RecordCount > 0)
                        {
                            oRecordset.DoQuery("Select \"Code\" from \"@OBJECTMST\" Where \"U_Module\" like '%" + clsMain.ProjectCode + "%' And \"U_ObjCode\" = '" + TableName + "' ");
                            if (oRecordset.RecordCount == 0)
                            {
                                oRecordset.DoQuery("Update \"@OBJECTMST\" set \"U_Module\" = '" + clsMain.ProjectCode + "' + ',' + ISNULL(\"U_Module\",'') where \"U_ObjCode\" = '" + TableName + "'");
                            }
                        }
                        else
                        {
                            oRec.DoQuery("Select max(cast(\"Code\" as integer)) From \"@OBJECTMST\"");
                            if (oRec.RecordCount > 0)
                            {
                                IntCode = Convert.ToInt32(oRec.Fields.Item(0).Value) + 1;
                            }
                            else
                            {
                                IntCode = 1;
                            }
                            oRecordset.DoQuery("Insert into \"@OBJECTMST\" (\"Code\",\"Name\",\"U_ObjCode\",\"U_ObjName\",\"U_Module\") Values ('" + IntCode + "','" + IntCode + "','" + TableName + "','" + Name + "','" + clsMain.ProjectCode + "') ");
                        }
                        clsMain.SBO_Application.StatusBar.SetText("Object Registered : " + CodeID + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    oRec = null;
                    oRecordset = null;
                    IntCode = 0;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                    oUserObjectMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void InsertOperation()
        {
            try
            {
                SAPbobsCOM.GeneralService oGeneralService = default(SAPbobsCOM.GeneralService);
                SAPbobsCOM.GeneralDataParams oHeaderParams = default(SAPbobsCOM.GeneralDataParams);
                SAPbobsCOM.GeneralData oHeadTableRow = default(SAPbobsCOM.GeneralData);
                SAPbobsCOM.CompanyService sCmp = clsMain.oCompany.GetCompanyService();
                oRecordset = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery("Select DocEntry from [@VCOPMST] Where U_OpCode in('Inward','Pre-Dispatch','Sample','Re-Qc')");
                if (oRecordset.RecordCount == 0)
                {
                    oGeneralService = sCmp.GetGeneralService("VCOPMST");
                    // Set the params for receiving a specific record
                    //oHeaderParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    //oHeaderParams.SetProperty("DocEntry", "1") 'YOUR RECORD CODE
                    oHeadTableRow = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    oHeadTableRow.SetProperty("Code", "1");
                    //oHeadTableRow.SetProperty("DocEntry", 1)
                    oHeadTableRow.SetProperty("U_OpCode", "Inward");
                    oHeadTableRow.SetProperty("U_OpName", "Inward Operation");
                    oHeadTableRow.SetProperty("U_OpType", "1");
                    oGeneralService.Add(oHeadTableRow);

                    oHeadTableRow.SetProperty("Code", "2");
                    //your record code
                    oHeadTableRow.SetProperty("U_OpCode", "Pre-Dispatch");
                    oHeadTableRow.SetProperty("U_OpName", "Pre-Dispatch Operation");
                    oHeadTableRow.SetProperty("U_OpType", "1");
                    oGeneralService.Add(oHeadTableRow);

                    oHeadTableRow.SetProperty("Code", "3");
                    //your record code
                    oHeadTableRow.SetProperty("U_OpCode", "Sample");
                    oHeadTableRow.SetProperty("U_OpName", "Sample Operation");
                    oHeadTableRow.SetProperty("U_OpType", "1");
                    oGeneralService.Add(oHeadTableRow);

                    oHeadTableRow.SetProperty("Code", "4");
                    //your record code
                    oHeadTableRow.SetProperty("U_OpCode", "Re-Qc");
                    oHeadTableRow.SetProperty("U_OpName", "Re-Qc Operation");
                    oHeadTableRow.SetProperty("U_OpType", "1");
                    oGeneralService.Add(oHeadTableRow);
                }

                //oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                //oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                //Dim DocEntry As Double

                //oRec.DoQuery("Select ISNULL(Max(DocEntry),0)+1 from [@VCOPMST]")
                //DocEntry = oRec.Fields.Item(0).Value
                //oRecordset.DoQuery("Select Code from [@VCOPMST] Where U_OpCode = 'Inward'")
                //If oRecordset.RecordCount = 0 Then
                //    oRecordset.DoQuery("INSERT INTO [@VCOPMST] ([Code],[DocEntry],[Object],[U_OpCode],[U_OpName],[U_OpType]) Values ('" & DocEntry & "','" & DocEntry & "','VCOPMST','Inward','Inward Operation','1')")
                //    oRecordset.DoQuery("INSERT INTO [@VCOPMST] ([Code],[DocEntry],[Object],[U_OpCode],[U_OpName],[U_OpType]) Values ('" & DocEntry + 1 & "','" & DocEntry + 1 & "','VCOPMST','Pre-Dispatch','Pre-Dispatch Operation','1')")
                //    oRecordset.DoQuery("INSERT INTO [@VCOPMST] ([Code],[DocEntry],[Object],[U_OpCode],[U_OpName],[U_OpType]) Values ('" & DocEntry + 2 & "','" & DocEntry + 2 & "','VCOPMST','Sample','Sample Operation','1')")
                //    oRecordset.DoQuery("INSERT INTO [@VCOPMST] ([Code],[DocEntry],[Object],[U_OpCode],[U_OpName],[U_OpType]) Values ('" & DocEntry + 3 & "','" & DocEntry + 3 & "','VCOPMST','Re-PP','Re-PP Operation','1')")
                //End If
                //DocEntry = Nothing
                oGeneralService = null;
                oHeaderParams = null;
                oHeadTableRow = null;
                sCmp = null;
                oRecordset = null;
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Insert to OPMST : " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void AutoFMS()
        {
            oRecordset = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("Select \"CategoryId\" from OQCN Where \"CatName\" like '" + clsMain.ProjectName + "'");
            if (oRecordset.RecordCount == 0)
            {
                QueryCatGen();

            }
            QueryGen();
            AttchFMS();
            oRecordset = null;
        }

        private void QueryCatGen()
        {
            oRecordset = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery("Select (Max(\"CategoryId\")+1) from OQCN");
            oRecordset.DoQuery("INSERT INTO OQCN (\"CategoryId\",\"CatName\",\"PermMask\",\"DataSource\",\"UserSign\") Values (" + oRec.Fields.Item(0).Value + ",'" + clsMain.ProjectName + "','YNNNNNNNNNNNNNN','I','1')");
            //oRec.DoQuery("Select (Max(CategoryId)+1) from OQCN");
            //oRecordset.DoQuery("INSERT INTO OQCN ([CategoryId],[CatName],[PermMask],[DataSource],[UserSign]) Values (" + oRec.Fields.Item(0).Value + ",'GatePass FMS','YNNNNNNNNNNNNNN','I','1')");
        }

        private void QueryGen()
        {
            try
            {
                oRecordset = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec1 = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery("Select \"CategoryId\" from OQCN Where \"CatName\" = '" + clsMain.ProjectName + "'");

                oRec1.DoQuery("Select \"IntrnalKey\" from OUQR Where \"QName\" = 'FMS:Cost Center Group'");
                if (oRec1.RecordCount == 0)
                {
                    oRec1.DoQuery("Select (Max(\"IntrnalKey\")+1) from OUQR");
                    oRecordset.DoQuery("INSERT INTO OUQR (\"IntrnalKey\",\"QCategory\",\"QName\",\"QString\",\"QType\") Values (" + oRec1.Fields.Item(0).Value + "," + oRec.Fields.Item(0).Value + ",'FMS:Cost Center Group','select * from \"@OPRG\" T0','W')");
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void AttchFMS()
        {
            oRecordset = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec1 = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //Check Machine Serial No Attached or Not
            oRec.DoQuery("Select \"FormID\" from CSHS Where \"ItemID\" = '12' and \"FormID\" = '810'");
            if (oRec.RecordCount == 0)
            {
                oRec.DoQuery("Select (Max(\"IndexID\")+1) from CSHS");
                oRec1.DoQuery("Select \"IntrnalKey\" from OUQR Where \"QName\" = 'FMS:Cost Center Group'");
                oRecordset.DoQuery("INSERT INTO CSHS (\"FormID\",\"ItemID\",\"ColID\",\"ActionT\",\"QueryId\",\"IndexID\",\"Refresh\",\"FieldID\",\"FrceRfrsh\",\"ByField\") Values ('810','12','-1','2'," + oRec1.Fields.Item(0).Value + "," + oRec.Fields.Item(0).Value + ",'N','NULL','N','Y')");
            }
        }

        private void Alerts()
        {
            SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
            SAPbobsCOM.Recordset oRec1 = default(SAPbobsCOM.Recordset);
            SAPbobsCOM.Recordset oRecordset = default(SAPbobsCOM.Recordset);
            oRecordset = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec1 = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery("Select ISNULL(MAX(Code)+1,1) from OALT Where Code > 0");
            oRec1.DoQuery("Select IntrnalKey,QCategory from OUQR Where QName = 'PP Approval'");
            oRecordset.DoQuery("INSERT INTO OALT ([Code],[Name],[Type],[Priority],[Active],[NumOfParam],[ParamData],[Params],[NumOfDocs],[DocsData],[Docs],[UserText],[QueryId],[FrqncyType],[FrqncyIntr],[ExecDaY],[ExecTime],[UserSign],[History],[QCategory]) Values ('" + oRec.Fields.Item(0).Value + "','PP Approval','U','1','Y','0','NULL','NULL','0','NULL','NULL','NULL','" + oRec1.Fields.Item(0).Value + "','S','2','1','800','1','N','" + oRec1.Fields.Item(1).Value + "')");
            oRec = null;
            oRec1 = null;
            oRecordset = null;
        }

        private bool FieldExist(string TableName, string ColumnName)
        {
            SAPbobsCOM.Recordset oRecordSet = default(SAPbobsCOM.Recordset);
            oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + TableName + "' AND \"AliasID\" = '" + ColumnName + "'");
                if ((Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0))
                {
                    oRecordset = null;
                    return false;
                }
                else
                {
                    oRecordset = null;
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

    }
}
