using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using SAPbouiCOM.Framework;

namespace InventoryTransferBranch.Forms
{
    [FormAttribute("143", "Forms/GrpoSap1.b1f")]
    class GrpoSap1 : SystemFormBase
    {
        public static bool FromDelivery { get; set; }
        public static int DeliveryDocEntry { get; set; }
        public GrpoSap1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataAddAfter += new DataAddAfterHandler(this.Form_DataAddAfter);

        }

        private void Form_DataAddAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {

            string xmlObjectKey = pVal.ObjectKey;
            XElement xmlnew = XElement.Parse(xmlObjectKey);

            try
            {
                int docEntry = int.Parse(xmlnew.Element("DocEntry").Value);
                if (FromDelivery)
                {
                    DeliverySap.CloseDliveryUpdayeGrpo(DeliveryDocEntry, docEntry);
                }
            }
            catch (Exception e)
            {

            }
            FromDelivery = false;
            DeliveryDocEntry = 0;
        }

        private void OnCustomInitialize()
        {

        }
    }
}
