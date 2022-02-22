using Microsoft.Extensions.DependencyInjection;
using YunMa.Excel.Exporter;
using YunMa.Excel.Exporter.Base.Filter;
using YunMa.Excel.Importer;
using YunMa.Excel.Importer.Base.Filter;
using YunMa.Excel.Tests.ExporterTests.Filters;
using YunMa.Excel.Tests.ImporterTests.Filters;

namespace YunMa.Excel.Tests
{
    public class Startup
    {
        public void ConfigureServices(IServiceCollection services)
        {
            //注入
            services.AddTransient<IExcelExporter, ExcelExporter>();
            services.AddSingleton<IExporterHeaderFilter, TestExporterHeaderFilter1>();
            services.AddSingleton<IExporterHeaderFilter, TestExporterHeaderFilter2>();


            services.AddTransient<IExcelImporter, ExcelImporter>();
            services.AddSingleton<IImportHeaderFilter, ImportStudentDtoHeaderFilterTest>();
        }
    }
}