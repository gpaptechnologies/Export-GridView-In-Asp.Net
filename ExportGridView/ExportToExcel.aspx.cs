using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;

namespace ExportGridView
{
    public partial class ExportToExcel : System.Web.UI.Page
    {
        BLL_EmployeeData objEmployeeData = new BLL_EmployeeData();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                BindEmployeeData();
            }
        }

        protected void btnExport_Click(object sender, EventArgs e)
        {
            Export();
        }

        private void BindEmployeeData()
        {
            try
            {
                DataTable dtEmployees = null;
                dtEmployees = objEmployeeData.GetEmployeeData();

                if (dtEmployees.Rows.Count > 0)
                {
                    gvEmployeeData.DataSource = dtEmployees;
                    gvEmployeeData.DataBind();
                }
                else
                {
                    lblMessage.Text = "No Data Found...!";
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;
                lblMessage.ForeColor = System.Drawing.Color.Red;
            }
        }

        private void Export()
        {
            try
            {
                StringWriter objSW = null;
                HtmlTextWriter objHTW = null;

                if (gvEmployeeData.Rows.Count > 0)
                {
                    gvEmployeeData.ShowHeader = true;
                    gvEmployeeData.GridLines = GridLines.Both;
                    objSW = new StringWriter();
                    objHTW = new HtmlTextWriter(objSW);

                    Response.ClearContent();
                    Response.Buffer = true;
                    Response.AddHeader("content-disposition", string.Format("attachment; filename=EmployeeMasterData_{0:dd-MMM-yy}", DateTime.Now.Date) + ".xls");
                    Response.ContentType = "application/ms-excel";

                    gvEmployeeData.RenderControl(objHTW);
                    Response.Write(objSW.ToString());
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.Message;
                lblMessage.ForeColor = System.Drawing.Color.Red;
            }
        }

        public override void VerifyRenderingInServerForm(Control control)
        {
            /* Verifies that the control is rendered */
        }
    }
}