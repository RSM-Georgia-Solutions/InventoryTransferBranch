using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace InventoryTransferBranch
{
    [FormAttribute("InventoryTransferBranch.PostedGRPO", "PostedGRPO.b1f")]
    class PostedGRPO : UserFormBase
    {
        private readonly string _docEntry;
        private readonly string _docNum;

        public PostedGRPO(string docEntry, string docNum)
        {
            _docEntry = docEntry;
            _docNum = docNum;
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_1").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.VisibleAfter += new VisibleAfterHandler(this.Form_VisibleAfter);

        }

        private SAPbouiCOM.EditText EditText0;

        private void OnCustomInitialize()
        {
            
        }

        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText1;

        private void Form_VisibleAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
           var title =  Application.SBO_Application.Forms.ActiveForm.Title;
            if (title == "Posted Grpo")
            {
                EditText0.Item.Left = 10000;
                EditText0.Value = _docEntry;
                EditText1.Value = _docNum;
            }
        }
    }
}
