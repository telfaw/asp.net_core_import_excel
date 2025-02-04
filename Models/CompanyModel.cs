using System.ComponentModel.DataAnnotations;

namespace Asp_mvc_import_excel_file.Models
{
    public class CompanyModel
    {
        [Key]
        public int CompanyId { get; set; }

        public string CompanyName { get; set; }
        public string CompanyDescription { get; set; }
        public string CompanyAddress { get; set; }
        public DateTime DateOfList { get; set; }

        public int OldYear { get; set; }
        public int currentYear { get; set; }


    }
}
