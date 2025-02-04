using System.Data;

namespace Asp_mvc_import_excel_file.Repository
{
    public interface ICompanyInterface
    {
        string DocumentUpload(IFormFile formFile);
        DataTable CompanyDataTable(string filePath);
        void ImportData(DataTable dataTable);
    }
}
