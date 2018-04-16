using System;
using System.Collections.Generic;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
namespace DataLaundry
{
  
    public class ExcelHandler
    {


        public ExcelHandler()
        {
        
        }
        
        public List<Company> ExcelReader()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "CompaniesBB.xlsx";

            System.Diagnostics.Debug.WriteLine(path);

            //Instance reference for Excel Application
            Excel.Application objXL = null;
            //Workbook refrence

            
            List<Company> companies = new List<Company>();

            objXL = new Excel.ApplicationClass();
            //Adding WorkBook

            System.Globalization.CultureInfo oldCI;
            oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("sv-SE");

            Excel.Workbook objWB = objXL.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, Excel.XlCorruptLoad.xlExtractData);

            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;

            try
            {
                //Instancing Excel using COM services
            

                foreach (Excel.Worksheet objSHT in objWB.Worksheets)
                {
                    int rows = objSHT.UsedRange.Rows.Count;
                    int cols = objSHT.UsedRange.Columns.Count;

                  

                    for (int r = 2; r <= rows; r++)
                    {

                        Company company = new Company();

                        company.ID_bb = (objSHT.Cells[r, 1] as Excel.Range).Value?.ToString();
                        company.CompanyName = (objSHT.Cells[r, 2] as Excel.Range).Value?.ToString();
                        company.Verksamhet_BB = (objSHT.Cells[r, 3] as Excel.Range).Value?.ToString();
                        company.ParVAT = (objSHT.Cells[r, 5] as Excel.Range).Value?.ToString();
                        company.Phone = (objSHT.Cells[r, 6] as Excel.Range).Value?.ToString();
                        company.StreetAddress2 = (objSHT.Cells[r, 8] as Excel.Range).Value?.ToString();
                        company.City2_BB = (objSHT.Cells[r, 9] as Excel.Range).Value?.ToString();
                        company.PostalCode2_BB = (objSHT.Cells[r, 10] as Excel.Range).Value?.ToString();
                        company.Fax_BB = (objSHT.Cells[r, 11] as Excel.Range).Value?.ToString();
                        company.Ansvarig_BB = (objSHT.Cells[r, 14] as Excel.Range).Value?.ToString();
                        company.HS_Owner = (objSHT.Cells[r, 14] as Excel.Range).Value?.ToString();
                        company.Employees = (objSHT.Cells[r, 16] as Excel.Range).Value?.ToString();
                        company.StreetAddress = (objSHT.Cells[r, 17] as Excel.Range).Value?.ToString();
                        company.Postal = (objSHT.Cells[r, 18] as Excel.Range).Value?.ToString();
                        company.City = (objSHT.Cells[r, 19] as Excel.Range).Value?.ToString();
                        company.Kundnr_BB = (objSHT.Cells[r, 24] as Excel.Range).Value?.ToString();
                        company.Par_Kalla_BB = (objSHT.Cells[r, 25] as Excel.Range).Value?.ToString();
                        company.StateRegion = (objSHT.Cells[r, 27] as Excel.Range).Value?.ToString();
                        company.Turnover = (objSHT.Cells[r, 30] as Excel.Range).Value?.ToString();
                        company.Orgnr_BB = (objSHT.Cells[r, 31] as Excel.Range).Value?.ToString();
                        company.ParID_BB = (objSHT.Cells[r, 49] as Excel.Range).Value?.ToString();
                        company.SNI_kod_BB = (objSHT.Cells[r, 56] as Excel.Range).Value?.ToString();
                        company.Website_URL = (objSHT.Cells[r, 61] as Excel.Range).Value?.ToString();
                        company.Description = (objSHT.Cells[r, 66] as Excel.Range).Value?.ToString();
                        company.ArbetsStalleNr_BB = (objSHT.Cells[r, 67] as Excel.Range).Value?.ToString();
                        company.Type = "Customer";

                        companies.Add(company);

                    }
                    

                }

                //Closing workbook
                objWB.Close();
                //Closing excel application
                objXL.Quit();

            }
            catch (Exception ex)
            {
                string mess = ex.Message;
                objWB.Saved = true;
                //Closing work book
                objWB.Close();
                //Closing excel application
                objXL.Quit();
                //Response.Write("Illegal permission");
            }

            return companies;
        }
    }
}