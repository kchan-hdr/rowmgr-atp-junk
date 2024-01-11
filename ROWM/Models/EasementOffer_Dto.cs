using ROWM.Dal;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ROWM.Models
{
    public class EasementOffer_Dto
    {
        public static Dictionary<string, string> Cast(Parcel p)
        {
            var o = p.Ownership.FirstOrDefault()?.Owner ?? new Owner();

            List<Owner> ownersList = p.Ownership.Select(os => os.Owner).ToList();

            string[] addressParts = o.OwnerAddress.Split(',');

            string addressLine1 = "";
            string addressLine2 = "";

            if (addressParts!=null && addressParts.Length >= 2)
            {
                addressLine1 = addressParts[0].Trim();
                addressLine2 = string.Join(", ", addressParts.Skip(1)).Trim();
            }

            Debug.WriteLine(p.ParcelId);
            Debug.WriteLine(ownersList);
            Debug.WriteLine(addressLine1);

            //var TotalAcres = p.ProjectInfo?.EasementAcreage;

            //var CropAcres = p.ProjectInfo?.EasementAcreage_Crop;

            //var PastureAcres = p.ProjectInfo?.EasementAcreage_Pasture;

            //var Crop_Percentage = (TotalAcres == null || TotalAcres == 0 || CropAcres == null) ? 0 : 100 * CropAcres / TotalAcres;

            //var Pasture_Percentage = (TotalAcres == null || TotalAcres == 0 || PastureAcres == null) ? 0 : 100 * PastureAcres / TotalAcres;

            //double CropValue = p.ProjectInfo == null || p.ProjectInfo.EasementAcreage_Crop == null ? 0 : (p.ProjectInfo.EasementAcreage_Crop.GetValueOrDefault() * 3350);

            //double PastureValue = p.ProjectInfo == null || p.ProjectInfo.EasementAcreage_Pasture == null ? 0 : (p.ProjectInfo.EasementAcreage_Pasture.GetValueOrDefault() * 2350);

            //double TotalComp = CropValue + PastureValue;

            //double OptionComp = 0.1 * TotalComp;

            //double EasementComp = 0.9 * TotalComp;

            return new Dictionary<string, string>
            {                
                { "County_Name", p.County_Name.ToUpper() },
                { "Landowner_name_1", ownersList.Count >= 1 ? ownersList[0].PartyName : " "},
                { "Landowner_name_2", ownersList.Count >= 2 ? ownersList[1].PartyName : " "},
                { "Owner_mailing_address", o.OwnerAddress },
                { "Owner_mailing_address_line_1", addressLine1 },
                { "Owner_mailing_address_line_2", addressLine2 },
                //{ "Legal_Description_of_optioned_areas", p.ProjectInfo?.LegalDescription },
                //{ "Project_Number", p.ProjectInfo?.ProjectNumber },
                //{ "Basin_ID", p.ProjectInfo?.Acquisition },
                { "HDR_MAP_ID", p.Tracking_Number },
                //{ "STR", p.ProjectInfo.STR  },
                { "Description", "See EXHIBIT B"  },
                //{ "Total_Acres", p.ProjectInfo?.EasementAcreage?.ToString("0.00") },
                //{ "Crop_Acres", p.ProjectInfo?.EasementAcreage_Crop?.ToString("0.00") },
                //{ "Crop_Percentage", Crop_Percentage.ToString() + "%" },
                //{ "Crop_Value", CropValue.ToString("#,##") },
                //{ "Pasture_Acres", p.ProjectInfo?.EasementAcreage_Pasture?.ToString("0.00") },
                //{ "Pasture_Percentage", Pasture_Percentage.ToString() + "%" },
                //{ "Pasture_Value", PastureValue.ToString("#,##") },
                //{ "Option_Compensation", OptionComp.ToString("#,##") },
                //{ "Easement_Compensation", EasementComp.ToString("#,##") },
                //{ "Total_Compensation", TotalComp.ToString("#,##") }
            };
        }       
    }
}
