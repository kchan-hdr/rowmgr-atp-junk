using geographia.ags;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var c = new ROWM.Dal.ROWM_Context("data source=tcp:atp-rowm-prod.database.windows.net;initial catalog=rowm;persist security info=True;user id=rowm_app;password=SbhrDX6Cq5VPcR9z;MultipleActiveResultSets=True;App=ROWM");

            var status = c.Parcel_Status.Where(sx => sx.Category == "engagement").ToArray();
            var parcels = c.Parcel.Where(px => px.IsActive && !px.IsDeleted).Select(px => new { p = px.Tracking_Number, c = px.OutreachStatusCode });

            IFeatureUpdate feat = new AtpParcel("https://maps.hdrgateway.com/arcgis/rest/services/Texas/ATP_Parcel_FS/FeatureServer");


            foreach ( var parcel in parcels.OrderBy(px => px.p))
            {
                var dv = status.First(sx => sx.Code == parcel.c);
                Console.WriteLine($"{parcel.p} {parcel.c} {dv.DomainValue}");

                try
                {
                    var good = await feat.UpdateFeatureOutreach(parcel.p, parcel.p, dv.DomainValue.Value, "", null);
                    Console.WriteLine($"{parcel.p} {parcel.c} {dv.DomainValue} ==  {good}");
                }
                catch( Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }


            Console.ReadKey();
        }
    }
}
