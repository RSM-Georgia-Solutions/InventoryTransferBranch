using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace InventoryTransferBranch.Forms
{
    [FormAttribute("InventoryTransferBranch.Forms.ReturnToGrpo", "Forms/ReturnToGrpo.b1f")]
    class ReturnToGrpo : UserFormBase
    {
        public ReturnToGrpo()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_7").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.VisibleAfter += new VisibleAfterHandler(this.Form_VisibleAfter);

        }

        private SAPbouiCOM.StaticText StaticText0;
        private UserDataSources userDataSources;

        private void OnCustomInitialize()
        {

            DiManager.Recordset.DoQuery("SELECT distinct U_ImportType FROM ORDN");

            while (!DiManager.Recordset.EoF)
            {
                var tmpType = DiManager.Recordset.Fields.Item("U_ImportType").Value.ToString();
                ComboBox0.ValidValues.Add(tmpType, tmpType);
                DiManager.Recordset.MoveNext();
            }

            SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(this.SBO_Application_ItemEvent_ChooseFromList);
        }

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            var bplId = userDataSources.Item("BplId").Value;
            Recordset recSet =
                (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes
                    .BoRecordset);
            recSet.DoQuery($"SELECT * FROM ODLN WHERE DocDate Between '{EditText0.Value}' And '{EditText1.Value}' And U_ImportType = '{ComboBox0.Selected.Value}' And BplId = {bplId} AND U_GrpoDocEntry is null");


            
            Documents deliveryObj = (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oDeliveryNotes);

            while (!recSet.EoF)
            {
                string returnDoc = recSet.Fields.Item("DocEntry").Value.ToString();
                deliveryObj.GetByKey(int.Parse(returnDoc));

                var deliveryXmlString = deliveryObj.GetAsXML();
                XElement DelXml = XElement.Parse(deliveryXmlString);
                List<string> costs = new List<string>();
                try
                {
                    foreach (var node in DelXml.Element("BO").Element("DLN1").Elements("row").Elements().Where(x => x.Name == "StockPrice"))
                    {
                        var cost = node.Value;
                        costs.Add(cost);
                    }
                }
                catch (Exception e)
                {

                }
                BusinessPartners businessPartner = (BusinessPartners)DiManager.Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                businessPartner.GetByKey(deliveryObj.CardCode);
                string whs = string.Empty;
                try
                {
                    whs = deliveryObj.UserFields.Fields.Item("U_WareHouse").Value.ToString();
                }

                catch (Exception e)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("WareHouse ველი დასამატებელია",
                        BoMessageTime.bmt_Short, true);
                }

                DiManager.Recordset.DoQuery(DiManager.QueryHanaTransalte($"SELECT OBPL.BPLName, OBPL.BPLId FROM OWHS JOIN OBPL ON OWHS.BPLid = OBPL.BPLId WHERE OWHS.WhsCode = N'{whs}'"));
                int branchId = int.Parse(DiManager.Recordset.Fields.Item("BPLid").Value.ToString());
                string branchName = DiManager.Recordset.Fields.Item("BPLName").Value.ToString();

                Documents goodsReceiptPo = (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes);
                goodsReceiptPo.CardCode = businessPartner.LinkedBusinessPartner;
                goodsReceiptPo.BPL_IDAssignedToInvoice = branchId;
                goodsReceiptPo.DocDate = deliveryObj.DocDueDate;
                for (int i = 0; i < deliveryObj.Lines.Count; i++)
                {
                    deliveryObj.Lines.SetCurrentLine(i);
                    goodsReceiptPo.Lines.ItemCode = deliveryObj.Lines.ItemCode;
                    goodsReceiptPo.Lines.Quantity = deliveryObj.Lines.Quantity;
                    goodsReceiptPo.Lines.UnitPrice = double.Parse(costs[i], CultureInfo.InvariantCulture);
                    goodsReceiptPo.Lines.WarehouseCode = whs;
                    goodsReceiptPo.Lines.Add();
                }

                deliveryObj.Close();
                var res = goodsReceiptPo.Add();
                if (res == 0)
                {
                    string docEntry = DiManager.Company.GetNewObjectKey();
                    string docNum;
                    DiManager.Company.GetNewObjectCode(out docNum);
                    try
                    {
                        deliveryObj.UserFields.Fields.Item("U_GrpoDocEntry").Value = docEntry;
                        deliveryObj.Update();
                    }
                    catch (Exception e)
                    {
                       var  err1 = DiManager.Company.GetLastErrorDescription();
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetSystemMessage(err1, BoMessageTime.bmt_Short);
                    }
                }
                var err = DiManager.Company.GetLastErrorDescription();
                if (string.IsNullOrWhiteSpace(err))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetSystemMessage("წარმატება", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetSystemMessage(err, BoMessageTime.bmt_Short);
                }

                recSet.MoveNext();
            }

        }

        private void SBO_Application_ItemEvent_ChooseFromList(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormTypeEx != "InventoryTransferBranch.Forms.ReturnToGrpo")
            {
                return;
            }
            if (pVal.ItemUID == "Item_5")
            {
                ChooseFromList(FormUID, pVal, "Item_5", "", "BplId", "BplName");
            }
        }

        private void ChooseFromList(string FormUID, ItemEvent pVal, string itemUId, string itemUIdDesc, string dataSourceId = "", string dataSourceDescId = "", bool isMatrix = false, string matrixUid = "")
        {
            if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
            {
                string val = null;
                string val2 = null;
                try
                {
                    IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((IChooseFromListEvent)(pVal));
                    string sCFL_ID = null;
                    sCFL_ID = oCFLEvento.ChooseFromListUID;
                    Form oForm = null;
                    oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = null;
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                    if (oCFLEvento.BeforeAction == false)
                    {
                        DataTable oDataTable = null;
                        oDataTable = oCFLEvento.SelectedObjects;

                        try
                        {
                            val = Convert.ToString(oDataTable.GetValue(0, 0));
                            val2 = Convert.ToString(oDataTable.GetValue(1, 0));
                        }
                        catch (Exception ex)
                        {

                        }
                        if (pVal.ItemUID == itemUId || pVal.ItemUID == matrixUid)
                        {
                            if (isMatrix)
                            {
                                //Grid0.DataTable.SetValue(itemUId, pVal.Row, val);
                                //Grid0.DataTable.SetValue(itemUIdDesc, pVal.Row, val2);
                            }
                            else if (pVal.ItemUID == itemUId)
                            {
                                var xz = SAPbouiCOM.Framework.Application.SBO_Application.Forms.GetForm("InventoryTransferBranch.Forms.ReturnToGrpo", 1);

                                xz.DataSources.UserDataSources.Item(dataSourceId).ValueEx = val2;
                                xz.DataSources.UserDataSources.Item(dataSourceId).Value = val;
                                if (!string.IsNullOrWhiteSpace(dataSourceDescId))
                                {
                                    xz.DataSources.UserDataSources.Item(dataSourceDescId).Value = val2;
                                }

                            }
                        }
                    }
                }
                catch (Exception e)
                {
                }
            }
        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private StaticText StaticText3;
        private ComboBox ComboBox0;
        private Button Button0;

        private void Form_VisibleAfter(SBOItemEventArg pVal)
        {
            userDataSources = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.DataSources.UserDataSources;
        }
    }
}
