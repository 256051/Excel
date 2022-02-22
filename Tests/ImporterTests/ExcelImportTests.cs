using OfficeOpenXml;
using Shouldly;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using YunMa.Excel.Core.Extensions;
using YunMa.Excel.Importer;
using YunMa.Excel.Importer.Base.Attributes;
using YunMa.Excel.Tests.ImporterTests.Models;

namespace YunMa.Excel.Tests.ImporterTests
{
    public class ExcelImportTests : TestBase
    {
        private readonly IExcelImporter _excelImporter;

        public ExcelImportTests(IExcelImporter excelImporter)
        {
            _excelImporter = excelImporter;
        }

        /// <summary>
        /// 生成学生数据导入模板
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "生成学生数据导入模板")]
        public async Task GenerateStudentImportTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                nameof(GenerateStudentImportTemplate_Test) + ".xlsx");
            DeleteFile(filePath);

            var result = await _excelImporter.GenerateTemplate<GenerateStudentImportTemplateDto>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            // using var pck = new ExcelPackage(new FileInfo(filePath));
            // pck.Workbook.Worksheets.Count.ShouldBe(3);
            // var sheet = pck.Workbook.Worksheets.First();
            // var dataValidataions = sheet.DataValidations.FirstOrDefault()
            //     as OfficeOpenXml.DataValidation.ExcelDataValidationList;
            // dataValidataions.Formula.ExcelFormula.ShouldBe("hidden_Gender!$A$1:$A$2");
            // //TODO:读取Excel检查表头和格式
        }

        /// <summary>
        /// 生成学生数据导入模板带有数据验证
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "生成学生数据导入模板带有数据验证")]
        public async Task GenerateStudentImportSheetDataValidationTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                nameof(GenerateStudentImportSheetDataValidationTemplate_Test) + ".xlsx");
            DeleteFile(filePath);

            var result = await _excelImporter.GenerateTemplate<GenerateStudentImportSheetDataValidationDto>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Cells["A2"].Style.Numberformat.Format.ShouldBe("@");
            }
            //TODO:读取Excel检查表头和格式
        }


        /// <summary>
        /// 测试生成导入描述头
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "生成学生数据导入模板加描述")]
        public async Task GenerateStudentImportSheetDescriptionTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                nameof(GenerateStudentImportSheetDescriptionTemplate_Test) + ".xlsx");
            DeleteFile(filePath);

            var result = await _excelImporter.GenerateTemplate<ImportStudentDtoWithSheetDesc>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using var pck = new ExcelPackage(new FileInfo(filePath));
            pck.Workbook.Worksheets.Count.ShouldBe(3);
            var sheet = pck.Workbook.Worksheets.First();
            var attr = typeof(ImportStudentDtoWithSheetDesc).GetAttribute<ExcelImporterAttribute>();
            var text = sheet.Cells["A1"].Text.Replace("\n", string.Empty).Replace("\r", string.Empty);
            text.ShouldBe(attr.ImportDescription.Replace("\n", string.Empty).Replace("\r", string.Empty));
        }

        [Fact(DisplayName = "生成模板")]
        public async Task GenerateTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(GenerateTemplate_Test) + ".xlsx");
            DeleteFile(filePath);

            var result = await _excelImporter.GenerateTemplate<ImportProductDto>(filePath);
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Column(15).Style.Numberformat.Format.ShouldBe("yyyy-MM-dd");
            }
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            //TODO:读取Excel检查表头和格式
        }




        /// <summary>
        /// 测试：
        /// 表头行位置设置
        /// 导入逻辑测试
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "产品信息导入读取SheetName")]
        public async Task ImporterReadStream_Test()
        {
            //第一列乱序
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "产品导入模板.xlsx");
            await using var stream = new FileStream(filePath, FileMode.Open);
            var result = await _excelImporter.Import<ImportProductDto>(stream);
            result.ShouldNotBeNull();
        }

        /// <summary>
                /// 测试：
                /// 表头行位置设置
                /// 导入逻辑测试
                /// </summary>
                /// <returns></returns>
                [Fact(DisplayName = "产品信息导入")]
        public async Task ImporterWithStream_Test()
        {
            //第一列乱序
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "产品导入模板.xlsx");
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                var result = await _excelImporter.Import<ImportProductDto>(stream);
                result.ShouldNotBeNull();

                result.HasError.ShouldBeTrue();
                result.RowErrors.Count.ShouldBe(1);
                result.Data.ShouldNotBeNull();
                result.Data.Count.ShouldBeGreaterThanOrEqualTo(2);
                foreach (var item in result.Data)
                {
                    if (item.Name != null && item.Name.Contains("空格测试")) item.Name.ShouldBe(item.Name.Trim());

                    if (item.Code.Contains("不去除空格测试")) item.Code.ShouldContain(" ");
                    //去除中间空格测试
                    item.BarCode.ShouldBe("123123");
                }

                //可为空类型测试
                result.Data.ElementAt(4).Weight.HasValue.ShouldBe(true);
                result.Data.ElementAt(5).Weight.HasValue.ShouldBe(false);
                //提取性别公式测试
                result.Data.ElementAt(0).Sex.ShouldBe("女");
                //获取当前日期以及日期类型测试  如果时间不对，请打开对应的Excel即可更新为当前时间，然后再运行此单元测试
                //import.Data[0].FormulaTest.Date.ShouldBe(DateTime.Now.Date);
                //数值测试
                result.Data.ElementAt(0).DeclareValue.ShouldBe(123123);
                result.Data.ElementAt(0).Name.ShouldBe("1212");
                result.Data.ElementAt(0).BarCode.ShouldBe("123123");
                result.Data.ElementAt(0).ProductIdTest1.ShouldBe(Guid.Parse("C2EE3694-959A-4A87-BC8C-4003F6576352"));
                result.Data.ElementAt(0).ProductIdTest2.ShouldBe(Guid.Parse("C2EE3694-959A-4A87-BC8C-4003F6576357"));
                result.Data.ElementAt(1).Name.ShouldBe(null);
                result.Data.ElementAt(2).Name.ShouldBe("左侧空格测试");

                result.ImporterHeaderInfos.ShouldNotBeNull();
                result.ImporterHeaderInfos.Count.ShouldBe(17);
            }
        }
    }
}