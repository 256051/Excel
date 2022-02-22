using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using Shouldly;
using Xunit;
using YunMa.Excel.Core.Extensions;
using YunMa.Excel.Exporter;
using YunMa.Excel.Tests.ExporterTests.Models;
using YunMa.Excel.Tests.Models;

namespace YunMa.Excel.Tests.ExporterTests
{
    public class ExcelExporterTests : TestBase
    {
        private readonly IExcelExporter _excelExporter;

        public ExcelExporterTests(IExcelExporter excelExporter)
        {
            _excelExporter = excelExporter;
        }

        [Fact(DisplayName = "�������Excel")]
        public async Task ByBytes()
        {
            var list = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);

            var bytes = await _excelExporter.ExportAsByteArray(list);
            var stream = new MemoryStream(bytes);
            stream.Seek(0, SeekOrigin.Begin);
            await using var fs = File.Create("d:/1.xlsx");
            await stream.CopyToAsync(fs);
        }


        [Fact(DisplayName = "DTO���Ե��������Ը�ʽ���Լ���ͷ������")]
        public async Task AttrsExport_Test()
        {
            var filePath = GetTestFilePath($"{nameof(AttrsExport_Test)}.xlsx");

            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            foreach (var item in data)
            {
                item.LongNo = 458752665;
                item.Text = "���Գ��ȳ�����Ԫ����ַ���";
            }

            var result = await _excelExporter.Export(filePath, data);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using var pck = new ExcelPackage(new FileInfo(filePath));
            pck.Workbook.Worksheets.Count.ShouldBe(1);
            var sheet = pck.Workbook.Worksheets.First();
            sheet.Cells[sheet.Dimension.Address].Rows.ShouldBe(101);
            sheet.Cells["A2"].Text.ShouldBe(data[0].Text2);

            //[ExporterHeader(DisplayName = "����1", Format = "yyyy-MM-dd")]
            sheet.Cells["E2"].Text.Equals(DateTime.Parse(sheet.Cells["E2"].Text).ToString("yyyy-MM-dd"));

            //[ExporterHeader(DisplayName = "����2", Format = "yyyy-MM-dd HH:mm:ss")]
            sheet.Cells["F2"].Text.Equals(DateTime.Parse(sheet.Cells["F2"].Text).ToString("yyyy-MM-dd HH:mm:ss"));

            //Ĭ��DateTime
            sheet.Cells["G2"].Text.Equals(DateTime.Parse(sheet.Cells["G2"].Text).ToString("yyyy-MM-dd"));

            //��Ԫ���Ȳ���
            sheet.Column(7).Width.ShouldBe(100);

            sheet.Tables.Count.ShouldBe(1);

            var tb = sheet.Tables.First();
            tb.Columns.Count.ShouldBe(9);
            tb.Columns.First().Name.ShouldBe("��ͨ�ı�");

            sheet.Tables.First();
            tb.Columns.Count.ShouldBe(9);
            tb.Columns[2].Name.ShouldBe("�Ӵ��ı�");
        }

        [Fact(DisplayName = "DataTable���Type���͵���ByteArray Excel")]
        public async Task DynamicExportByType_ByteArray_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                nameof(DynamicExportByType_ByteArray_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var exportDatas = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(1000);
            var dt = exportDatas.ToDataTable();
            var result = await _excelExporter.ExportAsByteArray(dt, typeof(ExportTestDataWithAttrs));
            result.ShouldNotBeNull();
            await using (var file = File.OpenWrite(filePath))
            {
                file.Write(result, 0, result.Length);
            }

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //���ת�����
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Dimension.Columns.ShouldBe(9);
            }
        }




        [Fact(DisplayName = "ͷ��ɸѡ������")]
        public async Task ExporterHeaderFilter_Test()
        {
            var filePath1 = Path.Combine(Directory.GetCurrentDirectory(), $"{nameof(ExporterHeaderFilter_Test)}1.xlsx");

            #region ͨ��ɸѡ���޸�����

            DeleteFile(filePath1);

            var data1 = GenFu.GenFu.ListOf<ExporterHeaderFilterTestData1>();
            var result = await _excelExporter.Export(filePath1, data1);
            result.ShouldNotBeNull();
            File.Exists(filePath1).ShouldBeTrue();

            using var pck1 = new ExcelPackage(new FileInfo(filePath1));
            //���ת�����
            var sheet1 = pck1.Workbook.Worksheets.First();
            sheet1.Cells["D1"].Value.ShouldBe("Name");
            sheet1.Dimension.Columns.ShouldBe(4);

            #endregion ͨ��ɸѡ���޸�����

            #region ͨ��ɸѡ���޸ĺ�����

            var filePath2 = Path.Combine(Directory.GetCurrentDirectory(), $"{nameof(ExporterHeaderFilter_Test)}2.xlsx");
            DeleteFile(filePath2);
            var data2 = GenFu.GenFu.ListOf<ExporterHeaderFilterTestData2>();
            result = await _excelExporter.Export(filePath2, data2);
            result.ShouldNotBeNull();
            File.Exists(filePath2).ShouldBeTrue();

            using var pck = new ExcelPackage(new FileInfo(filePath2));
            //���ת�����
            var sheet = pck.Workbook.Worksheets.First();
            sheet.Dimension.Columns.ShouldBe(5);

            #endregion ͨ��ɸѡ���޸ĺ�����
        }


    }
}