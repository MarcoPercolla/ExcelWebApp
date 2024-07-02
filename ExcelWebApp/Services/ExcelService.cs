
using ExcelWebApp.Data;

namespace ExcelWebApp.Services
{
    public class ExcelService
    {
        private readonly DatabaseContext _databaseContext;

        public ExcelService(DatabaseContext databaseContext)
        {
            _databaseContext = databaseContext;
        }

        public void ProcessExcel(string filePath)
        {
            _databaseContext.ProcessExcelFile(filePath);
        }
    }
}
