using Austin_Costs;
using ExpenseTracking.Dal;
using geographia.ags;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using ROWM.Dal;
using ROWM.Dal.Repository;
using ROWM.Models;
using ROWM.Reports;
using SharePointInterface;
using System.Linq;

namespace ROWM
{
    /// <summary>
    /// this configures Status & Statistics for the new UI evaluation
    /// </summary>
    public class StartupAtpV2UI
    {
        public StartupAtpV2UI(IHostingEnvironment env)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(env.ContentRootPath)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddJsonFile($"appsettings.{env.EnvironmentName}.json", optional: true)
                .AddEnvironmentVariables();
            Configuration = builder.Build();
        }

        public IConfigurationRoot Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddCors();

            // Add framework services.
            services.AddApplicationInsightsTelemetry();
            services.AddMvc();

            services.Configure<FormOptions>(o =>
            {
                o.ValueLengthLimit = int.MaxValue;
                o.MultipartBodyLengthLimit = int.MaxValue;
            });

            var cs = Configuration.GetConnectionString("ROWM_Context");
            services.AddScoped<ROWM.Dal.ROWM_Context>(fac =>
            {
                var c = new Dal.ROWM_Context(cs);
                c.Database.CommandTimeout = 300;
                return c;
            });
            services.AddScoped<CostEstimateContext>(fac => new CostEstimateContext(cs));
            services.AddScoped<RelocationContext>(fac => new RelocationContext(cs));
            services.AddScoped(fac => new ExpenseContext(cs));


            services.AddScoped<ROWM.Dal.OwnerRepository>();
            services.AddScoped<ParcelStatusRepository>();
            services.AddScoped<ROWM.Dal.ContactInfoRepository>();
            services.AddScoped<RelocationRepository>();
            services.AddScoped<IRelocationCaseOps, RelocationCaseOps>();
            services.AddScoped<IExpenseTracking, ExpenseTracking_Op>();
            services.AddScoped<IStatisticsRepository, AustinFilteredStatisticsRepository>(); // MOD: 
            services.AddScoped<IActionItemRepository, ActionItemRepository>();
            services.AddScoped<ICostEstimateRepository, CostEstimateRepository>();
            services.AddScoped<ROWM.Dal.AppRepository>();
            services.AddScoped<ROWM.Reports_Ex.IRowmReports, ROWM.Reports_Ex.SyncFusionReport>();
            services.AddScoped<DeleteHelper>();
            services.AddSingleton<ROWM.Dal.DocTypes>(fac => new DocTypes(fac.GetRequiredService<ROWM_Context>()));
            services.AddScoped<Controllers.IParcelStatusHelper, Controllers.ParcelStatusHelperV2>();
            services.AddScoped<IUpdateParcelStatus, UpdateParcelStatus_austin2>();    // MOD:
            services.AddScoped<UpdateParcelStatus2>();
            services.AddScoped<IOwnershipHelper, ParcelOwnershipHelper>();  // ADD
            services.AddScoped<ActionItemNotification.Notification>();

            services.AddSingleton<IFeatureUpdate>(fac =>
            {
                var ctx = fac.GetRequiredService<ROWM_Context>();
                var lay = ctx.MapConfiguration.FirstOrDefault(mx => mx.IsActive && mx.LayerType == LayerType.Parcel);
                return (lay == null)
                    ? new AtpParcel("https://maps-stg.hdrgateway.com/arcgis/rest/services/Texas/ATP_Parcel_FS/FeatureServer")
                    : new AtpParcel(lay.AgsUrl);
            });
            services.AddSingleton<IRenderer>(fac =>
            {
                var ctx = fac.GetRequiredService<ROWM_Context>();
                var lay = ctx.MapConfiguration.FirstOrDefault(mx => mx.IsActive && mx.LayerType == LayerType.Reference && mx.ProjectPartId == null);
                return (lay == null)
                    ? new AtpParcel("https://maps-stg.hdrgateway.com/arcgis/rest/services/Texas/ATP_ROW_MapService/MapServer")
                    : new AtpParcel(lay.AgsUrl);
            });
            services.AddSingleton<IMapSymbology, AtpSymbology>();

            services.AddSingleton<TxDotNeogitations.ITxDotNegotiation, TxDotNeogitations.Sh72>();

            services.AddScoped<ISharePointCRUD, DenverNoOp>();
            //var msi = new AzureServiceTokenProvider();
            //var vaultClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(msi.KeyVaultTokenCallback));
            //var appid = vaultClient.GetSecretAsync("https://drippingsprings-keys.vault.azure.net/", "dev-id").GetAwaiter().GetResult();
            //var apps = vaultClient.GetSecretAsync("https://drippingsprings-keys.vault.azure.net/", "springs-secret").GetAwaiter().GetResult();
            //services.AddScoped<ISharePointCRUD, SharePointCRUD>(fac => new SharePointCRUD(
            //    __appId: appid.Value,
            //    __appSecret: apps.Value,
            //    _url: "https://hdroneview.sharepoint.com/sites/CoDS",
            //    subfolder: "Parcels",
            //    template: "Shared Documents/Parcels/_PARCEL_PARCEL No",
            //    d: fac.GetRequiredService<DocTypes>()));

            services.AddScoped<IRowmReports, AustinReport>();

            services.AddSingleton<SiteDecoration, Atp>();

            services.AddSwaggerDocument();

            services.AddLogging(b => {
                b.AddConsole();
                b.AddDebug();
            });            
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostingEnvironment env, ILoggerFactory loggerFactory)
        {
            var syncfusionKey = Configuration["syncfusion"];
            if (!string.IsNullOrWhiteSpace(syncfusionKey))
                Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(syncfusionKey);

            app.UseExceptionHandler("/Home/Error");
 
            app.UseStaticFiles();

            app.UseCors(builder => builder.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod().WithExposedHeaders("Content-Disposition"));

            app.UseMvc(routes =>
            {
                routes.MapRoute(
                    name: "default",
                    template: "{controller=Home}/{action=Index}/{id?}");
            });

            app.UseOpenApi();
            app.UseSwaggerUi3();
        }
    }
}
