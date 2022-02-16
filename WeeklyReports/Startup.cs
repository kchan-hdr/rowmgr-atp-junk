using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using ROWM.Dal;
using ROWM.Dal.Repository;
using ROWM.Reports;
using System;

[assembly: FunctionsStartup(typeof(WeeklyReports.Startup))]

namespace WeeklyReports
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            var c = Environment.GetEnvironmentVariable("ROWM_Context");

            builder.Services.AddScoped<ROWM_Context>(fac => new ROWM_Context(c));
            builder.Services.AddScoped<OwnerRepository>();
            builder.Services.AddScoped<IStatisticsRepository, FilteredStatisticsRepository>();
            builder.Services.AddScoped<IRowmReports, AustinReport>();
            builder.Services.AddScoped<Recipients>();
        }
    }
}
