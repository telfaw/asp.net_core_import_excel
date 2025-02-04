using Microsoft.EntityFrameworkCore;

namespace Asp_mvc_import_excel_file.Models
{
    public class CompanyDbContext:DbContext
    {
        public CompanyDbContext(DbContextOptions<CompanyDbContext> options) : base(options)
        {

        }
        public DbSet<CompanyModel> CompanyModels { get; set; }
        //public DbSet<Company> Companys { get; set; }

    }
}
