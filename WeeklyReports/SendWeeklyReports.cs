using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using ROWM.Reports;
using SendGrid;
using SendGrid.Helpers.Mail;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WeeklyReports
{
    public class SendWeeklyReports
    {
        readonly IRowmReports report;
        readonly Recipients recipientHandler;
        readonly string[] PUBLISHING = new string[] { "en1", "en2"};
        public SendWeeklyReports(IRowmReports r, Recipients rec) => (report, recipientHandler) = (r, rec);

        readonly string _content = @"Good Morning,<br /><br />Please see attached Community Engagement and Action Item report for this week.<br /><br />
                    Thanks,<br /><br />
                    <span style='color:#4298b5'><b>Lisa Cooper,</b></span><small> SR/WA, R/W-RAC<br />
                    <i>Senior Real Estate Services Project Manager</i></small><br /><br />
                    <b>HDR</b><br />
                    710 Hesters Crossing Rd #150<br />
                    Round Rock, TX 78681<br />
                    D 512-685-2968 M 512-801-5902<br />
                    <a href=""mailto: lisa.cooper@hdrinc.com"" >lisa.cooper@hdrinc.com</a><br />
                    <a href=""https://www.hdrinc.com"" >hdrinc.com/follow-us</a>";

        [FunctionName("SendWeeklyReports")]
        public async Task RunWeekly([TimerTrigger("%Weekly%")]TimerInfo myTimer, [SendGrid()] IAsyncCollector<SendGridMessage> msg, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            var dt = DateTime.Today;
            var allReports = report.GetReports();


            foreach (var rd in allReports.Where(rp => PUBLISHING.Contains(rp.ReportCode)))
            {
                log.LogInformation($"{rd.Caption} {rd.ReportCode}");

                if (rd.ReportCode.StartsWith("en") && int.TryParse(rd.ReportCode.Substring(2), out var project))
                {
                    var recipients = await recipientHandler.GetRecipients(project);

                    var message = new SendGridMessage { Subject = $"ATP {rd.Caption} - {dt.ToLongDateString()}" };
                    message.SetFrom(new EmailAddress("NO-REPLY@hdrinc.com", "OneView / Real Estate"));
                    message.AddTos(MakeList(recipients, isCopy: false));
                    message.AddCcs(MakeList(recipients, isCopy: true));
                    message.AddBccs(new List<EmailAddress> { new EmailAddress { Email = "Kelly.Chan@hdrinc.com" } });

                    message.AddContent(MimeType.Html, _content);

                    var payload = await report.GenerateReport(rd);
                    message.AddAttachment(payload.Filename, Convert.ToBase64String(payload.Content), payload.Mime, disposition: "attachment");

                    await msg.AddAsync(message);
                    await recipientHandler.Sent(project, DateTimeOffset.UtcNow);
                }
            }

            await msg.FlushAsync();
        }

        [FunctionName("OnDemandReports")]
        public async Task<IActionResult> RunOnDemand(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            [SendGrid()] IAsyncCollector<SendGridMessage> msg,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            var dt = DateTime.Today;
            var allReports = report.GetReports();

            foreach (var rd in allReports.Where(rp=> PUBLISHING.Contains(rp.ReportCode)))
            {
                log.LogInformation($"{rd.Caption} {rd.ReportCode}");

                if (rd.ReportCode.StartsWith("en") && int.TryParse(rd.ReportCode.Substring(2), out var project))
                {
                    var recipients = await recipientHandler.GetRecipients(project);

                    var message = new SendGridMessage { Subject = $"ATP {rd.Caption} - {dt.ToLongDateString()}" };
                    message.SetFrom(new EmailAddress("NO-REPLY@hdrinc.com", "OneView / Real Estate"));
                    message.AddTos(MakeList(recipients,isCopy: false));
                    message.AddCcs(MakeList(recipients, isCopy: true));
                    message.AddBccs(new List<EmailAddress> { new EmailAddress { Email = "Kelly.Chan@hdrinc.com" } });
                    message.AddContent(MimeType.Html, _content);

                    var payload = await report.GenerateReport(rd);
                    message.AddAttachment(payload.Filename, Convert.ToBase64String(payload.Content), payload.Mime);

                    await msg.AddAsync(message);
                    await recipientHandler.Sent(project, DateTimeOffset.UtcNow);
                }
            }

            await msg.FlushAsync();

            return new OkResult();
        }

        List<EmailAddress> MakeList(IEnumerable<ROWM.Dal.DistributionList> recipients, bool isCopy)
        {
            return recipients.Where(rx => rx.CcMode == isCopy).OrderBy(rx => rx.Mail).Select(rx => new EmailAddress { Email = rx.Mail }).ToList();
        }
    }
}
