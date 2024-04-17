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
                myCommand.CommandType = System.Data.CommandType.Text;
                if (myConnection.State != System.Data.ConnectionState.Open)
                    await myCommand.Connection.OpenAsync();

                IWorkbook workbook = CreateWorkingCopy(myReport);
                var headerStyle = workbook.Styles.Add($"ColumnHeaderStyle_{Guid.NewGuid()}");        // deconflict style name
                headerStyle.Font.Bold = true;

                foreach (var view in myReport.ExtraViews)
                {
                    IWorksheet currSheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == view.TabName) ?? workbook.Worksheets.Create(view.TabName ?? "Data");

                    myCommand.CommandText = view.ReportQuery ?? $"SELECT * FROM {view.ReportViewName}";
                    var myReader = await myCommand.ExecuteReaderAsync();

                    string rangeName = $"{view.TabName}_DataImportRange".Replace(" ", "");

                    IName namedRange = workbook.Names[rangeName];
                    if (namedRange != null)
                    {
                        currSheet.ImportDataReader(myReader, namedRange, true);
                    }
                    else { currSheet.ImportDataReader(myReader, true, 1, 1, true); }

                    if (!workbook.CustomDocumentProperties.Contains("templated"))
                    {
                        currSheet.SetDefaultRowStyle(1, headerStyle);
                        currSheet.UsedRange.AutofitColumns();
                    }
                }

                myConnection.Close();

                // meta data
                var dt = System.DateTimeOffset.Now.ToLocalTime().DateTime;
                IWorksheet meta = workbook.Worksheets.FirstOrDefault(s => s.Name == "Report Properties") ?? workbook.Worksheets.Create("Report Properties");

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
                        book.CustomDocumentProperties["templated"].Text = d.TemplateFileName;
                        book.Worksheets.Create("Report Properties");
                        return book;
                    }
                }
            }
            
            // always return something
            book = excelEngine.Excel.Workbooks.Create(0);
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
            var r = await _context.Database.SqlQuery<ReportList>(
                    "SELECT ReportId, Name, Description, ReportViewName, ReportQuery, TemplateFileName, TabName FROM App.Report WHERE IsActive = 1 AND ReportId = @code"
                    , new SqlParameter("code", code) 
                ).FirstOrDefaultAsync();

            var extraViews = await _context.Database.SqlQuery<ExtraView>(
                "SELECT TabName, ReportViewName, ReportQuery, DisplayOrder FROM App.Report_Extra_View WHERE IsActive = 1 AND ReportId = @code ORDER BY DisplayOrder"
                , new SqlParameter("code", code)
                ).ToListAsync();

            r.ExtraViews = CollectViews(extraViews, r);

            return r;
        }

        private ICollection<ExtraView> CollectViews(ICollection<ExtraView> extraViews, ReportList r)
        {
            var originalView = new ExtraView
            {
                TabName = r.TabName ?? r.Name,
                ReportViewName = r.ReportViewName,
                ReportQuery = r.ReportQuery,
                DisplayOrder = 0
            };

            extraViews.Add(originalView);

            return extraViews.OrderBy(view => view.DisplayOrder).ToList();
        }

        #endregion
    }

    public class ReportList
    {
        public int ReportId { get; private set; }
        public string Name { get; private set; }
        public string Description { get; private set; }
        public string TabName { get; private set; }
        public string ReportViewName { get; private set; }
        public string ReportQuery { get; private set; }
        public string TemplateFileName { get; private set; }
        public virtual ICollection<ExtraView> ExtraViews { get; set; } = new HashSet<ExtraView>();

    }

    public class ExtraView
    {
        public int ExtraViewId { get; set; }
        public string TabName { get; set; }
        public string Description { get; set; }
        public string ReportViewName { get; set; }
        public string ReportQuery { get; set; }
        public int ReportId { get; set; }
        public int DisplayOrder { get; set; }
        public bool IsActive { get; set; }
    }
}
