using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;
using System.Globalization;

namespace KSM_CostingModule
{
    class CostingModule
    {
        #region Variable

        private SAPbouiCOM.Form oForm, oForm1, oFormCFL;
        private SAPbouiCOM.Button oBtn = null;
        private SAPbouiCOM.Item oItem, oItem1, oItem2, oItem3;
        private SAPbouiCOM.ComboBox oCombo;

        private SAPbouiCOM.Grid oGrid;

        private SAPbouiCOM.Matrix oMatrix1, oMatrix2;
        private Boolean ACTION = false;
        private SAPbobsCOM.Recordset oRecordSet, oRec1, oRec;
        private int Mode;
        private int i, DelLine;
        string Query = "", MatName = "";
        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;
        double RegularQuota = 0.0, AdditionaQuota = 0.0;
        string TransferBP = "", dtName = "";
        private SAPbouiCOM.DataTable DBDataTable;

        #endregion

        #region  Item event
        public bool itemevent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {
            try
            {
                oForm = clsMain.SBO_Application.Forms.GetForm(FormId, pVal.FormTypeCount);
                oMatrix1 = oForm.Items.Item("matDetail1").Specific;
                oMatrix2 = oForm.Items.Item("matDetail2").Specific;

                if (pVal.BeforeAction == true)
                {

                }
                switch (pVal.EventType)
                {

                    #region CFL
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "matDetail" && pVal.ColUID == "V_4")
                            {
                                CFLCondition("CFL_OCRD");
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                            oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                            string sCFL_ID = null;
                            sCFL_ID = oCFLEvento.ChooseFromListUID;
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            if (pVal.ItemUID == "matDetail" && pVal.ColUID == "V_4")
                            {
                                try
                                {
                                    //oMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.value = oDataTable.GetValue("CardCode", 0).ToString();
                                }
                                catch (Exception ex) { }
                               

                            }                            
                        }
                        break;

                    #endregion

                    #region FORM_LOAD
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        if (pVal.BeforeAction == false)
                        {
                            try
                            {

                            }
                            catch (Exception Ex)

                            { }
                        }
                        break;

                    #endregion

                    #region ITEM_PRESSED
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        if (pVal.BeforeAction)
                        {
                            try
                            {
                                if (pVal.BeforeAction == true && pVal.ItemUID == "1" && (pVal.FormMode == 3 || pVal.FormMode == 2))
                                {
                                    Mode = pVal.FormMode;
                                    if (Validation() == false)
                                    {
                                        BubbleEvent = false;
                                        return false;
                                        break;
                                    }
                                }

                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "btnRun")
                            {
                                GetFGItemDetails();
                                //GetDetails();
                            }

                        }

                        break;

                    #endregion

                    #region COMBO_SELECT
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        try
                        {
                            if (pVal.BeforeAction)
                            {

                            }
                            else
                            {
                                if (pVal.ItemUID == "tStyle")
                                {
                                    oCombo = oForm.Items.Item("tStyle").Specific;

                                    oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    oRecordSet.DoQuery(@"Select * from OITM where ifnull(""U_SUBSERIES"",'') = '" + oCombo.Selected.Value + "'");
                                    if (oRecordSet.RecordCount > 0)
                                    {
                                        oForm.Items.Item("tGender").Specific.value = oRecordSet.Fields.Item("U_CATEGORIES").Value;
                                        oForm.Items.Item("tColor1").Specific.value = oRecordSet.Fields.Item("U_COLOURS").Value;
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {

                            throw;
                        }

                        break;

                    #endregion

                    #region KeyDOWN

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        try
                        {
                            if (pVal.BeforeAction == true)
                            {
                                if (pVal.CharPressed == 9)
                                {
                                    if (pVal.ItemUID == "23_U_E")
                                    {
                                        if (string.IsNullOrEmpty(oForm.Items.Item("23_U_E").Specific.value))
                                        {
                                            clsMain.SBO_Application.SendKeys("+{F2}");
                                            BubbleEvent = false;
                                        }
                                    }
                                }
                            }
                            if (pVal.BeforeAction == false)
                            {

                            }
                        }
                        catch
                        { }

                        break;

                    # endregion

                    # region LOST FOCUS

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "matDetail2" && pVal.ColUID == "Col_1")
                            {
                                if (!string.IsNullOrEmpty(oMatrix2.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.value))
                                {

                                }
                            }
                        }
                        break;

                    #endregion

                    #region FORM_RESIZE

                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                        if (pVal.BeforeAction == false)
                        {
                            try
                            {
                                SAPbouiCOM.Item oitem1;
                                SAPbouiCOM.Item oitem2;

                                oitem1 = oForm.Items.Item("matDetail1");
                                oitem2 = oForm.Items.Item("matDetail2");

                                oitem1.Height = 100;
                                oItem2.Top = oitem1.Top + oitem1.Height + 20;
                                oitem2.Height = 180;
                            }
                            catch
                            {


                            }
                        }
                        break;

                    #endregion

                    #region FORM_RESIZE

                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.BeforeAction == false)
                        {
                            try
                            {
                                if (pVal.ItemUID == "matDetail1" && pVal.ColUID == "V_-1")
                                {
                                    if (oMatrix1.RowCount > 0)
                                    {
                                        for (int i = 1; i <= oMatrix1.RowCount; i++)
                                        {
                                            oMatrix1.CommonSetting.SetRowFontStyle(i, BoFontStyle.fs_Plain);
                                            if (oMatrix1.IsRowSelected(i))
                                            {
                                                GetDetails(oMatrix1.Columns.Item("Col_0").Cells.Item(i).Specific.value);
                                                oMatrix1.CommonSetting.SetRowFontStyle(i, BoFontStyle.fs_Bold);
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {


                            }
                        }
                        break;

                        #endregion

                }
                return BubbleEvent;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        # endregion

        #region  MenuEvent

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type)
        {
            bool bevent = true;
            try
            {
                oForm = clsMain.SBO_Application.Forms.Item(FormId);
                oMatrix1 = oForm.Items.Item("matDetail1").Specific;
                oMatrix2 = oForm.Items.Item("matDetail2").Specific;

                if (Type == "Add")
                {
                    FillComboStyle();
                    oMatrix1.AutoResizeColumns();
                    oMatrix2.AutoResizeColumns();
                }

                else if (Type == "Find")
                {
                    Enable_Disable(FormId, true, Type);
                    oForm.Items.Item("tDocNum").Click();

                }

                else if (Type == "AddR")
                {
                    if (MatName == "matDetail")
                    {
                        AddMatrixRow("matDetail", "V_4");
                    }
                    if (MatName == "matDetail1")
                    {
                        AddMatrixRow("matDetail1", "V_4");
                    }
                    if (MatName == "matDetail2")
                    {
                        AddMatrixRow("matDetail2", "V_4");
                    }
                }
                else if (Type == "DEL")
                {
                    if (MatName == "matDetail")
                    {
                        DeleteMatrixRow(MatName);
                    }
                    if (MatName == "matDetail1")
                    {
                        DeleteMatrixRow(MatName);
                    }
                    if (MatName == "matDetail2")
                    {
                        DeleteMatrixRow(MatName);
                    }
                }
                return bevent;
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Menu Event : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }

        #endregion

        #region FormDataEvent
        public void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                oForm = clsMain.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                string psFormId = BusinessObjectInfo.FormUID;
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        if (BusinessObjectInfo.BeforeAction == true)
                        {
                            oForm = clsMain.SBO_Application.Forms.Item(psFormId);

                            //oForm.Items.Item("matDetail").Enabled = true;
                        }
                        else
                        {
                            oForm.Items.Item("tDocNum").Enabled = false;
                            oForm.Items.Item("tDocDate").Enabled = false;
                            oForm.Items.Item("tYear").Enabled = false;
                            oForm.Items.Item("cMonth").Enabled = false;
                            //oForm.Items.Item("matDetail").Enabled = false;
                        }

                        break;
                }
            }
            catch { }
        }
        #endregion

        #region Other Event

        private bool Validation()
        {
            try
            {
                //oForm = clsMain.SBO_Application.Forms.GetFormByTypeAndCount(60110, 1);
                oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (string.IsNullOrEmpty(oForm.Items.Item("tDocDate").Specific.value))
                {
                    //clsMain.SBO_Application.SetStatusBarMessage("Please select Document Date", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    //oForm.Items.Item("tDocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    //return false;
                }                

            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Validation : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
            return true;
        }

        public void FillComboStyle()
        {
            oRec = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string Str = @"Select distinct ""U_SUBSERIES"" from OITM where ifnull(""U_SUBSERIES"",'')<>'' and ifnull(""U_SUBSERIES"",'') = 'G10-G2005' ";
            //string Str = @"Select distinct ""U_SUBSERIES"" from OITM where ifnull(""U_SUBSERIES"",'')<>'' ";
            oRec.DoQuery(Str);
            if (oRec.RecordCount > 0)
            {
                oCombo = oForm.Items.Item("tStyle").Specific;
                while (oCombo.ValidValues.Count > 0)
                {
                    oCombo.ValidValues.Remove(oCombo.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                if (oCombo.ValidValues.Count == 0)
                {
                    for (int i = 0; i < oRec.RecordCount; i++)
                    {
                        oCombo.ValidValues.Add("" + oRec.Fields.Item("U_SUBSERIES").Value + "", "" + oRec.Fields.Item("U_SUBSERIES").Value + "");
                        oRec.MoveNext();
                    }
                }
                oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
        }

        public void SetCode()
        {
            try
            {
                oForm = clsMain.SBO_Application.Forms.ActiveForm;
                //oForm.Freeze(true);
                //oForm.Items.Item("tDocNum").Enabled = true;
                //oForm.Items.Item("tDocDate").Enabled = true;
                //oForm.Items.Item("tYear").Enabled = true;
                //oForm.Items.Item("cMonth").Enabled = true;
                //oForm.Items.Item("matDetail").Enabled = true;
                //clsMain.SetCode(oForm.UniqueID, "LPG_ALLOC", "D");
                //oForm.PaneLevel = 1;
                //oForm.Items.Item("fldDetail").Click(BoCellClickType.ct_Regular);

                //oForm.Items.Item("tDocDate").Click();
                //oForm.Items.Item("tDocNum").Enabled = false;
                //oMatrix = oForm.Items.Item("matDetail").Specific;
                //oMatrix.Clear();
                //AddMatrixRow("matDetail", "V_4");
                FillComboStyle();


            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Set Code :" + ex, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void Enable_Disable(string FormUID, bool pbVal, string FormMode)
        {
            oForm = clsMain.SBO_Application.Forms.Item(FormUID);
            oForm.Freeze(true);
            if (FormMode == "Add")
            {
                //oForm.Items.Item("tRmk").Click();
                //oForm.Items.Item("tDocNum").Enabled = pbVal;                
                //oForm.Items.Item("tDocDate").Enabled = pbVal;
                //oForm.Items.Item("tYear").Enabled = pbVal;
                //oForm.Items.Item("cMonth").Enabled = pbVal;
                //oMatrix1.AutoResizeColumns();
            }
            else if (FormMode == "Find")
            {
                //oForm.Items.Item("tDocNum").Enabled = pbVal;
                //oForm.Items.Item("tDocDate").Enabled = pbVal;
                //oForm.Items.Item("tYear").Enabled = pbVal;
                //oForm.Items.Item("cMonth").Enabled = pbVal;
            }

            oForm.Freeze(false);
        }


        public void AddMatrixRow(string MatrixName, string ColName)
        {
            try
            {
                oForm.Freeze(true);
                oMatrix1 = oForm.Items.Item(MatrixName).Specific;
                if (oMatrix1.RowCount == 0)
                {
                    oMatrix1.AddRow();
                    ClearMatrixRow();
                }
                else
                {
                    if (oMatrix1.RowCount >= 1)
                    {
                        if (!string.IsNullOrEmpty(oMatrix1.Columns.Item(ColName).Cells.Item(oMatrix1.RowCount).Specific.Value))
                        {
                            oMatrix1.AddRow();
                            ClearMatrixRow();
                        }
                    }
                    else
                    {
                        //oMatrix1.AddRow();
                        //ClearMatrixRow();
                    }
                }
                oMatrix1.Columns.Item("V_-1").Cells.Item(oMatrix1.RowCount).Specific.Value = oMatrix1.RowCount;
                //oMatrix1.Columns.Item("Col_0").Cells.Item(oMatrix1.RowCount).Click();
                oForm.Freeze(false);
            }
            catch (Exception)
            {
                oForm.Freeze(false);
            }

        }

        private void ClearMatrixRow()
        {

            try
            {
                oForm.Freeze(true);
                if (oMatrix1.RowCount > 1)
                {
                    oMatrix1.Columns.Item("V_-1").Cells.Item(oMatrix1.RowCount).Specific.value = Convert.ToInt32(oMatrix1.Columns.Item("V_-1").Cells.Item(oMatrix1.RowCount - 1).Specific.value) + 1;
                    oMatrix1.ClearRowData(oMatrix1.RowCount);
                }
                else
                {
                    oMatrix1.Columns.Item("V_-1").Cells.Item(oMatrix1.RowCount).Specific.value = oMatrix1.RowCount;
                    oMatrix1.ClearRowData(oMatrix1.RowCount);
                }
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Clear Matrix Row : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private Boolean DeleteMatrixRow(string MatrixName)
        {
            try
            {
                oMatrix1 = oForm.Items.Item(MatrixName).Specific;
                if (oMatrix1.RowCount == 1)
                {
                    ClearMatrixRow();
                    oMatrix1.Columns.Item("V_-1").Cells.Item(oMatrix1.RowCount).Specific.Value = oMatrix1.RowCount;
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    return false;
                }
                else
                {
                    oMatrix1.DeleteRow(DelLine);
                    //clsMain.SBO_Application.SetStatusBarMessage("Please wait....", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    for (int i = 1; i <= oMatrix1.RowCount; i++)
                    {
                        oMatrix1.Columns.Item("V_-1").Cells.Item(i).Specific.Value = i;
                    }
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    return false;
                }
                // return true;
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Delete Row : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }

        public object CFLCondition(string CFL)
        {
            try
            {
                oCFL = oForm.ChooseFromLists.Item(CFL);
                oConds = clsMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                oRecordSet = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (CFL == "CFL_OCRD" || CFL == "CFL_OCRD1" || CFL == "CFL_OCRD2")
                {
                    oRecordSet.DoQuery("select \"CardCode\" from OCRD where \"CardType\" = 'C' and \"CardCode\" like ('CBTL%') Order by \"CardCode\"");
                    if (oRecordSet.RecordCount > 0)
                    {
                        for (int i = 1; i <= oRecordSet.RecordCount; i++)
                        {
                            oCond = oConds.Add();
                            oCond.Alias = "CardCode";
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = oRecordSet.Fields.Item(0).Value.ToString();
                            if (i != oRecordSet.RecordCount)
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            oRecordSet.MoveNext();
                        }
                        oCFL.SetConditions(oConds);
                        return true;
                    }
                    else
                    {
                        clsMain.SBO_Application.SetStatusBarMessage("No Record found ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        oCond = oConds.Add();
                        oCond.Alias = "";
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCond.CondVal = null;
                        oCFL.SetConditions(oConds);
                        return true;
                    }
                }

                //if (CFL == "CFL_OCRD2")
                //{
                //    oRecordSet.DoQuery("select \"CardCode\" from OCRD where \"CardType\" = 'S'");
                //    if (oRecordSet.RecordCount > 0)
                //    {
                //        for (int i = 1; i <= oRecordSet.RecordCount; i++)
                //        {
                //            oCond = oConds.Add();
                //            oCond.Alias = "CardCode";
                //            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //            oCond.CondVal = oRecordSet.Fields.Item(0).Value.ToString();
                //            if (i != oRecordSet.RecordCount)
                //                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                //            oRecordSet.MoveNext();
                //        }
                //        oCFL.SetConditions(oConds);
                //        return true;
                //    }
                //    else
                //    {
                //        clsMain.SBO_Application.SetStatusBarMessage("No Record found ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                //        oCond = oConds.Add();
                //        oCond.Alias = "";
                //        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //        oCond.CondVal = null;
                //        oCFL.SetConditions(oConds);
                //        return true;
                //    }

                //}
                return true;

            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("CFL : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }


        public void GetFGItemDetails()
        {
            try
            {
                dtName = "DataTable" + DateTime.Now.ToString();
                DBDataTable = oForm.DataSources.DataTables.Add(dtName);

                Query = "CALL GetFGItemDetails('" + oForm.Items.Item("tStyle").Specific.value + "')";

                DBDataTable.Clear();
                DBDataTable.ExecuteQuery(Query);

                if (DBDataTable.Rows.Count > 0)
                {
                    clsMain.SBO_Application.SetStatusBarMessage("Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                    oMatrix1.Clear();
                    oMatrix2.Clear();
                    oMatrix1.Columns.Item("V_-1").DataBind.Bind(dtName, "TempLine");
                    oMatrix1.Columns.Item("Col_0").DataBind.Bind(dtName, "Code");
                    oMatrix1.Columns.Item("Col_2").DataBind.Bind(dtName, "ItemName");
                    oMatrix1.Columns.Item("Col_6").DataBind.Bind(dtName, "PROCESS");
                    oMatrix1.Columns.Item("Col_4").DataBind.Bind(dtName, "UOM");
                    oMatrix1.Columns.Item("V_3").DataBind.Bind(dtName, "Quantity");
                    oMatrix1.Columns.Item("Col_1").DataBind.Bind(dtName, "COLOURS");
                    //oMatrix1.Columns.Item("Col_3").DataBind.Bind(dtName, "Rate");
                    //oMatrix1.Columns.Item("Col_5").DataBind.Bind(dtName, "Total");
                    oMatrix1.Columns.Item("Col_7").DataBind.Bind(dtName, "ProdQty");
                    oMatrix1.Columns.Item("Col_8").DataBind.Bind(dtName, "Std.Consumption");

                    if (DBDataTable.Rows.Count == 1)
                    {
                        if (DBDataTable.GetValue("Code", 0).ToString() == "0")
                        {
                            DBDataTable.Rows.Remove(0);
                        }
                    }
                    clsMain.SBO_Application.SetStatusBarMessage("Displaying Data. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                    oMatrix1.LoadFromDataSource();
                    oMatrix1.AutoResizeColumns();                    
                }

            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("GetDetails :" + ex, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
        public void GetDetails(string ItemCode)
        {
            try
            {
                dtName = "DataTable2" + DateTime.Now.ToString();
                DBDataTable = oForm.DataSources.DataTables.Add(dtName);

                Query = "CALL GetItemDetails('" + oForm.Items.Item("tStyle").Specific.value + "', '" + ItemCode + "')";

                DBDataTable.Clear();
                DBDataTable.ExecuteQuery(Query);

                if (DBDataTable.Rows.Count > 0)
                {
                    clsMain.SBO_Application.SetStatusBarMessage("Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                    oMatrix2.Clear();
                    oMatrix2.Columns.Item("V_-1").DataBind.Bind(dtName, "TempLine");
                    oMatrix2.Columns.Item("Col_0").DataBind.Bind(dtName, "Code");
                    oMatrix2.Columns.Item("Col_2").DataBind.Bind(dtName, "ItemName");
                    oMatrix2.Columns.Item("Col_6").DataBind.Bind(dtName, "PROCESS");
                    oMatrix2.Columns.Item("Col_4").DataBind.Bind(dtName, "UOM");
                    oMatrix2.Columns.Item("V_3").DataBind.Bind(dtName, "Norms");
                    oMatrix2.Columns.Item("Col_1").DataBind.Bind(dtName, "COLOURS");
                    oMatrix2.Columns.Item("Col_3").DataBind.Bind(dtName, "Rate");
                    oMatrix2.Columns.Item("Col_5").DataBind.Bind(dtName, "Total");
                    oMatrix2.Columns.Item("Col_7").DataBind.Bind(dtName, "ProdQty");
                    oMatrix2.Columns.Item("Col_8").DataBind.Bind(dtName, "Std.Consumption");

                    if (DBDataTable.Rows.Count == 1)
                    {
                        if (DBDataTable.GetValue("Code", 0).ToString() == "0")
                        {
                            DBDataTable.Rows.Remove(0);
                        }
                    }
                    clsMain.SBO_Application.SetStatusBarMessage("Displaying Data. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                    oMatrix2.LoadFromDataSource();
                    oMatrix2.AutoResizeColumns();

                    //oMatrix.CommonSetting.SetRowBackColor(1, 135);

                    



                    //SAPbobsCOM.Documents oDraft = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);

                    //if (!oDraft.GetByKey(5195))
                    //{
                    //    throw new Exception("Failed to retrieve the draft order." + clsMain.oCompany.GetLastErrorDescription());
                    //}

                    //SAPbobsCOM.Documents oSalesOrder = clsMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                    //// Copy the details from the draft order to the new sales order
                    //oSalesOrder.CardCode = oDraft.CardCode;
                    //oSalesOrder.DocDate = oDraft.DocDate;
                    //oSalesOrder.DocDueDate = DateTime.Now;
                    //oSalesOrder.TaxDate = oDraft.TaxDate;
                    //oSalesOrder.Comments = oDraft.Comments;

                    //for (int i = 0; i < oDraft.Lines.Count; i++)
                    //{
                    //    oDraft.Lines.SetCurrentLine(i);
                    //    oSalesOrder.Lines.ItemCode = oDraft.Lines.ItemCode;
                    //    oSalesOrder.Lines.Quantity = oDraft.Lines.Quantity;
                    //    oSalesOrder.Lines.Price = oDraft.Lines.Price;
                    //    oSalesOrder.Lines.UoMEntry = oDraft.Lines.UoMEntry;

                    //    oSalesOrder.Lines.Add();
                    //}

                    //if (oSalesOrder.Add() != 0)
                    //{
                    //    throw new Exception($"Failed to add sales order. Error: " + clsMain.oCompany.GetLastErrorDescription());
                    //}
                    //else
                    //{
                    //    Console.WriteLine("Sales order created successfully.");
                    //}

                    //oDraft.GetByKey(5196);
                    //oDraft.SaveDraftToDocument();

                }

            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("GetDetails :" + ex, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        #endregion

        #region  Right Click Event

        public bool RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent, string FormId)
        {
            try
            {
                if (eventInfo.FormUID == FormId && eventInfo.BeforeAction == true)
                {
                    MatName = null;

                    oForm.EnableMenu("1294", false);//Duplicate Row
                    oForm.EnableMenu("1299", false);//Close Row
                    oForm.EnableMenu("1284", false);//Cancle
                    oForm.EnableMenu("1287", false);//Close Row
                    oForm.EnableMenu("771", false);//Cut Row
                    oForm.EnableMenu("774", false);//Delete Row
                    oForm.EnableMenu("775", false);//Row
                    oForm.EnableMenu("8802", false);//Maximize Row
                    oForm.EnableMenu("8801", false);//
                    oForm.EnableMenu("1292", false);//Add Row
                    oForm.EnableMenu("1293", false);//Delete Row
                    oForm.EnableMenu("784", false);//copy table Row

                    oForm.EnableMenu("772", false);//Copy RoweventInfo.ItemUID == "matDetail"
                    oForm.EnableMenu("773", false);//Paste Row


                    if (eventInfo.ItemUID == "matDetail" || eventInfo.ItemUID == "matDetail1" || eventInfo.ItemUID == "matDetail2")
                    {
                        oForm.EnableMenu("784", true);//copy table Row
                        oForm.EnableMenu("772", true);//Copy Row
                        oForm.EnableMenu("773", true);//Paste Row
                        oForm.EnableMenu("1292", true);//Add Row
                        oForm.EnableMenu("1293", true);//Delete Row
                    }

                    MatName = eventInfo.ItemUID;
                    DelLine = eventInfo.Row;
                }
                return true;
            }
            catch (Exception ex)
            {
                clsMain.SBO_Application.SetStatusBarMessage("Right Click Event : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return false;
            }
        }

        # endregion

    }
}
