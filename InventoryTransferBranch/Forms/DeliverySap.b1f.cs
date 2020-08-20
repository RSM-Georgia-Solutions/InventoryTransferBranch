using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using InventoryTransferBranch.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;

namespace InventoryTransferBranch
{
    [FormAttribute("140", "Forms/DeliverySap.b1f")]
    class DeliverySap : SystemFormBase
    {
        public DeliverySap()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            Button0 = ((Button)(GetItem("Item_3333").Specific));
            Button0.PressedAfter += new _IButtonEvents_PressedAfterEventHandler(Button0_PressedAfter);
            OnCustomInitialize();

        }

        public static void CloseDliveryUpdayeGrpo(int deliveryDocEntry, int grpoDocEntry)
        {
            Documents deliveryObj = (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
            deliveryObj.GetByKey(deliveryDocEntry);
            deliveryObj.UserFields.Fields.Item("U_GrpoDocEntry").Value = grpoDocEntry;
            var result = deliveryObj.Update();
            if (result == 0)
            {
                deliveryObj.Close();
            }
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {

        }

        private Button Button0;

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            Form deliveryForm = Application.SBO_Application.Forms.ActiveForm;


            if (deliveryForm.Type == 140)
            {
                Documents deliveryObj = (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                string deliveryDocEntry = deliveryForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0);
                if (string.IsNullOrWhiteSpace(deliveryDocEntry))
                {
                    Application.SBO_Application.SetStatusBarMessage("დოკუმენტი არაა დამატებული (Delivery)",
                        BoMessageTime.bmt_Short);
                    return;
                }

                deliveryObj.GetByKey(int.Parse(deliveryDocEntry));
                var grpoDocEntry = deliveryObj.UserFields.Fields.Item("U_GrpoDocEntry").Value.ToString();
                if (deliveryObj.DocumentStatus == BoStatus.bost_Close
                    || grpoDocEntry != "0")
                {
                    Application.SBO_Application.SetStatusBarMessage("უკვე გამოწერილია",
                        BoMessageTime.bmt_Short);
                    return;
                }

                BusinessPartners businessPartner = (BusinessPartners)DiManager.Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                businessPartner.GetByKey(deliveryObj.CardCode);
                if (string.IsNullOrWhiteSpace(businessPartner.LinkedBusinessPartner))
                {
                    Application.SBO_Application.SetStatusBarMessage("მიუთითეთ \"Connected Vendor\" ",
                        BoMessageTime.bmt_Short);
                    return;
                }
                int clicked = Application.SBO_Application.MessageBox("გსურთ დაპოსტვამდე დააკორექტიროთ დოკუმენტი ?", 1, "დიახ", "არა");

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
                string whs = string.Empty;
                try
                {
                    whs = deliveryObj.UserFields.Fields.Item("U_WareHouse").Value.ToString();
                }
                catch (Exception e)
                {
                    Application.SBO_Application.SetStatusBarMessage("WareHouse ველი დასამატებელია",
                        BoMessageTime.bmt_Short);
                    return;
                }

                if (string.IsNullOrWhiteSpace(whs))
                {
                    Application.SBO_Application.SetStatusBarMessage("შეავსეთ მიმღები საწყობი",
                        BoMessageTime.bmt_Short);
                    return;
                }
                DiManager.Recordset.DoQuery(DiManager.QueryHanaTransalte($"SELECT OBPL.BPLName, OBPL.BPLId FROM OWHS JOIN OBPL ON OWHS.BPLid = OBPL.BPLId WHERE OWHS.WhsCode = N'{whs}'"));
                int branchId = int.Parse(DiManager.Recordset.Fields.Item("BPLid").Value.ToString());
                string branchName = DiManager.Recordset.Fields.Item("BPLName").Value.ToString();

                if (clicked == 2)
                {
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

                    var res = goodsReceiptPo.Add();
                    if (res == 0)
                    {
                        string docEntry = DiManager.Company.GetNewObjectKey();
                        string docNum;
                        DiManager.Company.GetNewObjectCode(out docNum);
                        PostedGRPO postedGrpo = new PostedGRPO(docEntry, docNum);
                        try
                        {
                            deliveryObj.UserFields.Fields.Item("U_GrpoDocEntry").Value = docEntry;
                            var result = deliveryObj.Update();
                            if (result == 0)
                            {
                                deliveryObj.Close();
                            }

                        }
                        catch (Exception e)
                        {
                            var err1 = DiManager.Company.GetLastErrorDescription();
                            Application.SBO_Application.MessageBox("UDF - GrpoDocEntry დასამატებელია");
                        }

                        //postedGrpo.Show();
                    }
                    var err = DiManager.Company.GetLastErrorDescription();
                    if (string.IsNullOrWhiteSpace(err))
                    {
                        Application.SBO_Application.StatusBar.SetSystemMessage("წარმატება", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    }
                    else
                    {
                        Application.SBO_Application.StatusBar.SetSystemMessage(err, BoMessageTime.bmt_Short);
                    }
                }
                else
                {
                    GrpoSap1.FromDelivery = true;
                    GrpoSap1.DeliveryDocEntry = deliveryObj.DocEntry;
                    Application.SBO_Application.ActivateMenuItem("2306");
                    Form grpoForm = Application.SBO_Application.Forms.ActiveForm;
                    Matrix grpoFormMatrix = (Matrix)Application.SBO_Application.Forms.ActiveForm.Items.Item("38").Specific;

                    ((EditText)(grpoForm.Items.Item("4").Specific)).Value = businessPartner.LinkedBusinessPartner;
                    ((EditText)grpoForm.Items.Item("10").Specific).Value = deliveryObj.DocDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                    ((EditText)grpoForm.Items.Item("12").Specific).Value = deliveryObj.DocDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                    ComboBox comboItemType = (ComboBox)grpoForm.Items.Item("3").Specific;
                    comboItemType.Select(0, BoSearchKey.psk_Index);
                    var branchCombo = ((ComboBox)grpoForm.Items.Item("2001").Specific);
                    try
                    {
                        ((ComboBox)grpoForm.Items.Item("2001").Specific).Select(branchName);
                    }
                    catch (Exception e)
                    {
                        //Database Without Branch
                    }
                    //((ComboBox)(grpoForm.Items.Item("234000016").Specific)).Select("G", BoSearchKey.psk_ByValue); //Price Mode

                    for (int i = 0; i < deliveryObj.Lines.Count; i++)
                    {
                        int p = i + 1;
                        deliveryObj.Lines.SetCurrentLine(i);
                        EditText formItemCode = (EditText)grpoFormMatrix.Columns.Item("1").Cells.Item(p).Specific;
                        EditText formQuantity = (EditText)grpoFormMatrix.Columns.Item("11").Cells.Item(p).Specific;
                        EditText formWarehouseCode = (EditText)grpoFormMatrix.Columns.Item("24").Cells.Item(p).Specific;
                        EditText grossPrice;
                        grossPrice =
                            (EditText)grpoFormMatrix.Columns.Item("14").Cells.Item(p).Specific; //gross

                        grpoForm.Freeze(true);
                        grpoFormMatrix.AddRow();
                        formItemCode.Value = deliveryObj.Lines.ItemCode;
                        formQuantity.Value = deliveryObj.Lines.Quantity.ToString(CultureInfo.InvariantCulture);
                        formWarehouseCode.Value = whs;

                        grossPrice.Value = double.Parse(costs[i], CultureInfo.InvariantCulture)
                            .ToString(CultureInfo.InvariantCulture);

                        grpoForm.Freeze(false);
                    }
                }
            }

        }
    }
}
