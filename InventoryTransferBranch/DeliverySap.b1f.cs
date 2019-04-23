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
    [FormAttribute("140", "DeliverySap.b1f")]
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


                string whs = deliveryObj.UserFields.Fields.Item("U_WareHouse").Value.ToString();

                DiManager.Recordset.DoQuery(DiManager.QueryHanaTransalte($"SELECT * FROM OWHS WHERE WhsCode = N'{whs}'"));
                int branchId = int.Parse(DiManager.Recordset.Fields.Item("BPLid").Value.ToString());

                Documents goodsReceiptPo = (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes);
                goodsReceiptPo.CardCode = businessPartner.LinkedBusinessPartner;
                goodsReceiptPo.BPL_IDAssignedToInvoice = branchId;
                goodsReceiptPo.DocDate = deliveryObj.DocDate;

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
        }
    }
}
