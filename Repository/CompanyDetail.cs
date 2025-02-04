using Asp_mvc_import_excel_file.Repository;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Data.OleDb;

namespace Asp_mvc_import_excel_file.Repository
{
    public class CompanyDetail: ICompanyInterface

    {
        private IConfiguration configuration;
       
        private IWebHostEnvironment webHostingEnvironment;

        public CompanyDetail(IConfiguration configuration, IWebHostEnvironment webHostingEnvironment)
        {
            this.configuration = configuration;
            this.webHostingEnvironment = webHostingEnvironment;
            //_connectionString = configuration.GetConnectionString("DefaultConnection");
        }


        public DataTable CompanyDataTable(string filePath)
        {
           var conString= configuration.GetConnectionString("excelConnection");
            DataTable dt = new DataTable();
            conString=string.Format(conString, filePath);
            using(OleDbConnection excelConn = new OleDbConnection(conString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter adapterExcel=new OleDbDataAdapter()) 
                    { 
                      excelConn.Open();
                        cmd.Connection = excelConn;
                        DataTable excelSchema;
                        excelSchema = excelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        var sheetName= excelSchema.Rows[0]["TABLE_NAME"].ToString(); 
                        excelConn.Close();

                        excelConn.Open();
                        cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                        adapterExcel.SelectCommand = cmd;
                        adapterExcel.Fill(dt);
                        excelConn.Close();

                    }
                
                }
            }
            return dt;
        }

        public string DocumentUpload(IFormFile formFile)
        {
            string uploadsPath = webHostingEnvironment.WebRootPath;
            string destPath = Path.Combine(uploadsPath,"uploaded_doc");
            if (!Directory.Exists(destPath ))
            {
                Directory.CreateDirectory(destPath);
            }
            string sourceFielpath = Path.GetFileName( formFile.FileName);
            string path = Path.Combine(destPath, sourceFielpath);
            using (FileStream fileStream=new FileStream(path,FileMode.Create))
            {
                formFile.CopyTo(fileStream);    

            }
            return path;

        }

        public void ImportData(DataTable dataTable)
        {
            var defaultConnection = configuration.GetConnectionString("DefaultConnection");

            using (SqlConnection sqlconnection=new SqlConnection(defaultConnection)) 
            {
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlconnection)) 
                
                {
                    sqlBulkCopy.DestinationTableName = "CompanyModels";
                    sqlBulkCopy.ColumnMappings.Add("CompanyId", "CompanyId"); 
                    sqlBulkCopy.ColumnMappings.Add("CompanyName", "CompanyName"); 
                    sqlBulkCopy.ColumnMappings.Add("CompanyDescription", "CompanyDescription"); 
                    sqlBulkCopy.ColumnMappings.Add("CompanyAddress", "CompanyAddress"); 
                    sqlBulkCopy.ColumnMappings.Add("DateOfList", "DateOfList"); 
                    sqlBulkCopy.ColumnMappings.Add("OldYear", "OldYear"); 
                    sqlBulkCopy.ColumnMappings.Add("currentYear", "currentYear"); 
                    sqlconnection.Open();
                    sqlBulkCopy.WriteToServer(dataTable);
                    sqlconnection.Close();


                }
            }
        }
    }
}
