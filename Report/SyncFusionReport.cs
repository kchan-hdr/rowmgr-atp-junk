using ROWM.Dal;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Reports_Ex
{
    public class SyncFusionReport : IRowmReports
    {
        readonly ROWM_Context _context;
        public SyncFusionReport(ROWM_Context c) => _context = c;

        public async Task<ReportPayload> GenerateReport(string reportCode)
        {
            var d = await GetReportListFromDb(reportCode);
            return await GenerateReport(d);
        }

        public Task<ReportPayload> GenerateReport(ReportDef d) => GenerateReport(d.ReportCode);

        #region implementation
        async Task<ReportPayload> GenerateReport(ReportList myReport)
        {
            if (myReport is null)
            {
                return new ReportPayload();
            }

            try
            {
                var myConnection = _context.Database.Connection;
                var myCommand = myConnection.CreateCommand();
                myCommand.CommandText = myReport.ReportQuery ?? $"SELECT * FROM {myReport.ReportViewName}";
                myCommand.CommandType = System.Data.CommandType.Text;
                if (myConnection.State != System.Data.ConnectionState.Open)
                    await myCommand.Connection.OpenAsync();
                var myReader = await myCommand.ExecuteReaderAsync(System.Data.CommandBehavior.CloseConnection);

                IWorkbook workbook = CreateWorkingCopy(myReport);
                IWorksheet sheet = workbook.Worksheets[0];

                sheet.ImportDataReader(myReader, true, 1, 1, true);

                var headerStyle = workbook.Styles.Add($"ColumnHeaderStyle_{Guid.NewGuid()}");        // deconflict style name
                headerStyle.Font.Bold = true;

                sheet.SetDefaultRowStyle(1, headerStyle);
                sheet.UsedRange.AutofitColumns();

                // meta data
                var dt = System.DateTimeOffset.Now.ToLocalTime().DateTime;
                IWorksheet meta = workbook.Worksheets.Last();
                meta.Name = "Report Properties";
                meta.Range["A1"].Text = "Printed";
                meta.Range["B1"].Text = dt.ToLongDateString();
                meta.Range["B2"].Text = dt.ToLongTimeString();
                meta.Range["B3"].Text = System.TimeZoneInfo.Local.DisplayName;
                meta.Range["A5"].Text = myReport.Description ?? myReport.Name;

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0;
                    return new ReportPayload
                    {
                        Content = stream.ToArray(),
                        Mime = "Application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        Filename = myReport.Name + ".xlsx"
                    };
                }
            }
            catch (System.Data.Common.DbException dbe)
            {
                System.Diagnostics.Trace.TraceError(dbe.Message);
                return new ReportPayload();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.Message);
                return new ReportPayload();
            }
        }

        // handle templated reports
        IWorkbook CreateWorkingCopy(ReportList d)
        {
            var excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook book;

            if (!string.IsNullOrWhiteSpace(d.TemplateFileName))
            {
                var myPath = Path.Combine("ReportTemplates", d.TemplateFileName);
                using (var ms = new FileStream(myPath, FileMode.Open))
                {
                    if (ms != null)
                    {
                        book = excelEngine.Excel.Workbooks.Open(ms);
                        book.Worksheets.Create("Report Properties");
                        return book;
                    }
                }
            }
            
            // always return something
            book = excelEngine.Excel.Workbooks.Create(2);
            book.Worksheets[0].Name = "Data";
            return book;
        }
        #endregion

        public Task<IEnumerable<ReportDef>> GetReports() =>  GetReportListFromDb();

        #region private db helper
        async Task<IEnumerable<ReportDef>> GetReportListFromDb()
        {
            var reports = await _context.Database.SqlQuery<ReportList>("SELECT ReportId, Name, Description FROM App.Report WHERE IsActive = 1 ORDER BY DisplayOrder").ToListAsync();
            return reports.Select((r,idx) => new ReportDef 
            { 
                 ReportCode = r.ReportId.ToString(),
                 Caption = r.Description,
                 DisplayOrder = idx + 1     
            });
        }
        async Task<ReportList> GetReportListFromDb(string code)
        {
            return await _context.Database.SqlQuery<ReportList>(
                    "SELECT ReportId, Name, Description, ReportViewName, ReportQuery, TemplateFileName FROM App.Report WHERE IsActive = 1 AND ReportId = @code"
                    , new SqlParameter("code", code) 
                ).FirstOrDefaultAsync();
        }
        #endregion
    }

    public class ReportList
    {
        public int ReportId { get; private set; }
        public string Name { get; private set; }
        public string Description { get; private set; }
        public string ReportViewName { get; private set; }
        public string ReportQuery { get; private set; }
        public string TemplateFileName { get; private set; }
    }
}
