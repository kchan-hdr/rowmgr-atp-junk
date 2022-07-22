using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Austin_Costs
{
    [Table("Cost_Estimate2", Schema ="Austin")]
    public class ProjDetailsCostEstimate
    {
        [Key]
        public Guid EstimateId { get; private set; }

        [Column("Acquisition_Parcel_No")]
        public string AcqNo { get; private set; }

        [Column("Property_Owner_Name")]
        public string PropertyOwnerName { get; private set; }

        [Column("Situs_Address_Street_Address")]
        public string SitusAddress { get; private set; }

        [Column("Tax_Card_Mailing_Address_Street_Address")]
        public string MailAddress { get; private set; }
        
        [Column("Tax_Card_Mailing_Address_City_State_Zip")]
        public string MailingCityStateZip { get; private set;  }

        // Texas Secretary of State
        [Column("SOS_Mailing_Address_Street_Address")]
        public string SecretaryOfStateMailAddress { get; private set; }

        [Column("SOS_Mailing_Address_City_State_Zip")]
        public string SecretaryOfStateMailingCityStateZip { get; private set; }

        //[Column("Advanced_Acquisition_Parcel_Y_N")]
        //public string IsAdvancedAcquisition { get; set; }

        public string Project { get; private set; }

        [Column("Option")]
        public string OptionForCosting { get; private set; }

        public string Segment { get; private set; }

        [Column("Sub_Segment")]
        public string SubSegment { get; private set; }

        // Acquisition_Interest
        // NEPA_Classification

        [Column("ROW_Area_SqFt")]   // Fee Area SqFt
        public double? FeeArea { get; private set; }

        [Column("Permanent_Easement_Area_SqFt")]
        public double? PermanentEasementArea { get; private set; }

        [Column("TCE_Area_SqFt")]
        public double? TemporaryEasementArea { get; private set; }

        [Column("Utility_Easement_Area_Sq_Ft")]
        public double? UtilityEasemenetArea { get; private set; }

        [Column("Sum_of_ACQ_Land_Costs")]   // ACQ Land Cost
        public decimal? AcqLandCost { get; private set; }

        [Column("Total_ACQ_Cost_Gross")]
        public decimal? AcqTotalCost { get; private set; }

        [Column("Total_All_ROW_Costs")] // Total ROW Cost
        public decimal? TotalRowCost { get; set; }

        [Column("Needs_Landplan_Y_N")]
        public string LandplanYN { get; private set; }
        public bool NeedLandplan => "Y" == this.LandplanYN;

        [Column("Needs_Relocation_Y_N")]
        public string RelocationYN { get; private set; }
        public bool NeedRelocation => "Y" == this.RelocationYN;

        // ROE Status
    }
}
