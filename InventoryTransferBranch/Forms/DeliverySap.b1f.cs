using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
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
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_3333").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.OnCustomInitialize();

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

        private SAPbouiCOM.Button Button0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Form deliveryForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;


            if (deliveryForm.Type == 140)
            {
                Documents deliveryObj = (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                string deliveryDocEntry = deliveryForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0);
                if (string.IsNullOrWhiteSpace(deliveryDocEntry))
                {
                    Application.SBO_Application.SetStatusBarMessage("დოკუმენტი არაა დამატებული (Delivery)",
                        BoMessageTime.bmt_Short, true);
                    return;
                }

                deliveryObj.GetByKey(int.Parse(deliveryDocEntry));

                if (deliveryObj.DocumentStatus == BoStatus.bost_Close)
                {
                    Application.SBO_Application.SetStatusBarMessage("უკვე გამოწერილია",
                        BoMessageTime.bmt_Short, true);
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

                BusinessPartners businessPartner = (BusinessPartners)DiManager.Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                businessPartner.GetByKey(deliveryObj.CardCode);

                string whs = string.Empty;
                try
                {
                    whs = deliveryObj.UserFields.Fields.Item("U_WareHouse").Value.ToString();
                }
                catch (Exception e)
                {
                    Application.SBO_Application.SetStatusBarMessage("WareHouse ველი დასამატებელია",
                        BoMessageTime.bmt_Short, true);
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

                    deliveryObj.Close();
                    var res = goodsReceiptPo.Add();
                    if (res == 0)
                    {
                        string docEntry = DiManager.Company.GetNewObjectKey();
                        string docNum;
                        DiManager.Company.GetNewObjectCode(out docNum);
                        PostedGRPO postedGrpo = new PostedGRPO(docEntry, docNum);
                        postedGrpo.Show();
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
                }
                else
                {
                    Application.SBO_Application.ActivateMenuItem("2306");
                    Form grpoForm = Application.SBO_Application.Forms.ActiveForm;
                    Matrix grpoFormMatrix = (Matrix)Application.SBO_Application.Forms.ActiveForm.Items.Item("38").Specific;

                    ((EditText)(grpoForm.Items.Item("4").Specific)).Value = businessPartner.LinkedBusinessPartner;
                    ((EditText)grpoForm.Items.Item("10").Specific).Value = deliveryObj.DocDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                    ((EditText)grpoForm.Items.Item("12").Specific).Value = deliveryObj.DocDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                    ((ComboBox)(grpoForm.Items.Item("2001").Specific)).Select(branchName);
                    ((ComboBox)(grpoForm.Items.Item("234000016").Specific)).Select("G", BoSearchKey.psk_ByValue); //Price Mode

                    for (int i = 0; i < deliveryObj.Lines.Count; i++)
                    {
                        int p = i + 1;
                        deliveryObj.Lines.SetCurrentLine(i);
                        SAPbouiCOM.EditText formItemCode = (SAPbouiCOM.EditText)grpoFormMatrix.Columns.Item("1").Cells.Item(p).Specific;
                        SAPbouiCOM.EditText formQuantity = (SAPbouiCOM.EditText)grpoFormMatrix.Columns.Item("11").Cells.Item(p).Specific;
                        SAPbouiCOM.EditText formWarehouseCode = (SAPbouiCOM.EditText)grpoFormMatrix.Columns.Item("24").Cells.Item(p).Specific;
                        SAPbouiCOM.EditText grossPrice = (SAPbouiCOM.EditText)grpoFormMatrix.Columns.Item("288").Cells.Item(p).Specific;//gross



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
