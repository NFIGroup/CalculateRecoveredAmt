using System.AddIn;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Text;
using System.Data;
using System.ComponentModel;
using System.Linq;
using System;
using CalculateRecoveredAmt.RightNowService;


////////////////////////////////////////////////////////////////////////////////
//
// File: WorkspaceAddIn.cs
//
// Comments:
//
// Notes: 
//
// Pre-Conditions: 
//
////////////////////////////////////////////////////////////////////////////////
namespace CalculateRecoveredAmt
{
    public class WorkspaceAddIn : Panel, IWorkspaceComponent2
    {
        /// <summary>
        /// The current workspace record context.
        /// </summary>
        private IRecordContext _recordContext;
        public static IGlobalContext _globalContext { get; private set; }
        private static IGenericObject _sClaimRecord;
        RightNowConnectService _rnConnectService;

        /// <summary>
        /// Default constructor.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        public WorkspaceAddIn(bool inDesignMode, IRecordContext RecordContext, IGlobalContext GlobalContext)
        {
            _recordContext = RecordContext;
            _globalContext = GlobalContext;
            _rnConnectService = RightNowConnectService.GetService(_globalContext);
        }

        #region IAddInControl Members

        /// <summary>
        /// Method called by the Add-In framework to retrieve the control.
        /// </summary>
        /// <returns>The control, typically 'this'.</returns>
        public Control GetControl()
        {
            return this;
        }

        #endregion

        #region IWorkspaceComponent2 Members

        /// <summary>
        /// Sets the ReadOnly property of this control.
        /// </summary>
        public bool ReadOnly { get; set; }

        /// <summary>
        /// Method which is called when any Workspace Rule Action is invoked.
        /// </summary>
        /// <param name="ActionName">The name of the Workspace Rule Action that was invoked.</param>
        public void RuleActionInvoked(string ActionName)
        {
            string[] pResults;
            Decimal calc = 0;
            Decimal dAmt=0; 
             string amt="";

            try
            {
                _sClaimRecord = _recordContext.GetWorkspaceRecord(_recordContext.WorkspaceTypeName) as ICustomObject;
                string sClaim = _sClaimRecord.Id.ToString();
              
                pResults = _rnConnectService.getCredits(sClaim);


                List<string> creds = new List<string>();

                             
                switch (ActionName)
                {
                    case "Recovered":

                        
                        if (pResults != null && pResults.Length > 0)
                        {
                            creds = pResults.ToList();
                         
                            foreach (string c in creds)
                            {
                                
                                amt = c.Split('~')[0];
                                dAmt = Convert.ToDecimal(amt);
                         
                                calc = calc + dAmt;
                                                           
                            }
                           calc= Math.Round(calc, 2);
                        }
                              break;
                    default:
                        break;
                }
                

                SetFieldValue( "recovered_amount", calc.ToString());
                _recordContext.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Save);
                
            }
            catch (Exception e)
            { MessageBox.Show("error in recovered calculation"); }
        }

        /// <summary>
        /// Method which is called when any Workspace Rule Condition is invoked.
        /// </summary>
        /// <param name="ConditionName">The name of the Workspace Rule Condition that was invoked.</param>
        /// <returns>The result of the condition.</returns>
        public string RuleConditionInvoked(string ConditionName)
        {
            return string.Empty;
        }


        public string GetFieldValue(string fieldName)
        {
            
            IList<IGenericField> fields = _sClaimRecord.GenericFields;
            
            if (null != fields)
            {
                foreach (IGenericField field in fields)
                {
                   
                    if (field.Name.Equals(fieldName))
                    {
                        if (field.DataValue.Value != null)
                            return field.DataValue.Value.ToString();
                    }
                }
            }
            return "";
        }
        /// <summary>
        /// Method which is use to set value to a field using record Context 
        /// </summary>
        /// <param name="fieldName">field name</param>
        /// <param name="value">value of field</param>
        public void SetFieldValue(string fieldName, string value)
        {
            IList<IGenericField> fields = _sClaimRecord.GenericFields;

            if (null != fields)
            {
                foreach (IGenericField field in fields)
                {
                    if (field.Name.Equals(fieldName))
                    {


                        switch (field.DataType)
                        {
                            case RightNow.AddIns.Common.DataTypeEnum.STRING:
                                field.DataValue.Value = value;
                                break;
                            case RightNow.AddIns.Common.DataTypeEnum.BOOLEAN:

                                if (value == "1" || value.ToLower() == "true")
                                {
                                    field.DataValue.Value = true;
                                }
                                else if (value == "0" || value.ToLower() == "false")
                                {
                                    field.DataValue.Value = false;
                                }
                                break;
                            case RightNow.AddIns.Common.DataTypeEnum.INTEGER:

                                if (value.Trim() == "" || value.Trim() == null)
                                {
                                    field.DataValue.Value = null;
                                }
                                else
                                {
                                    field.DataValue.Value = Convert.ToInt32(value);
                                }
                                break;
                        }
                    }
                }
            }
            return;
        }

        #endregion
    }

    [AddIn("Workspace Factory AddIn", Version = "1.0.0.0")]
    public class WorkspaceAddInFactory : IWorkspaceComponentFactory2
    {
        #region IWorkspaceComponentFactory2 Members
        static public IGlobalContext _globalContext;
        /// <summary>
        /// Method which is invoked by the AddIn framework when the control is created.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        /// <returns>The control which implements the IWorkspaceComponent2 interface.</returns>
        public IWorkspaceComponent2 CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            return new WorkspaceAddIn(inDesignMode, RecordContext, _globalContext);
        }

        #endregion

        #region IFactoryBase Members

        /// <summary>
        /// The 16x16 pixel icon to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public Image Image16
        {
            get { return Properties.Resources.AddIn16; }
        }

        /// <summary>
        /// The text to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Text
        {
            get { return "Calculate Recovered"; }
        }

        /// <summary>
        /// The tooltip displayed when hovering over the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Tooltip
        {
            get { return "Calculate Recovered"; }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            _globalContext = GlobalContext;
            return true;
        }

        #endregion
    }
}