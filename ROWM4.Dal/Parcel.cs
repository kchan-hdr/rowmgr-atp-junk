﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ROWM.Dal
{
    [Table("Parcel", Schema = "ROWM")]
    public partial class Parcel
    {
        public Parcel()
        {
            Appraisal = new HashSet<Appraisal>();
            InverseParcelParcel = new HashSet<Parcel>();
            Ownership = new HashSet<Ownership>();
            ParcelContactInfo = new HashSet<ParcelContactInfo>();
            ParcelContactLogs = new HashSet<ParcelContactLogs>();
            ParcelDocuments = new HashSet<ParcelDocuments>();
            RoeConditions = new HashSet<RoeConditions>();
            StatusActivity = new HashSet<StatusActivity>();
        }

        [Key]
        public Guid ParcelId { get; set; }
        [Required]
        [Column("Assessor_Parcel_Number")]
        [StringLength(128)]
        public string AssessorParcelNumber { get; set; }
        [Required]
        [Column("County_FIPS")]
        [StringLength(5)]
        public string CountyFips { get; set; }
        [Column("County_Name")]
        [StringLength(50)]
        public string CountyName { get; set; }
        [StringLength(800)]
        public string SitusAddress { get; set; }
        public double? Acreage { get; set; }
        public DateTimeOffset Created { get; set; }
        public DateTimeOffset LastModified { get; set; }
        [StringLength(50)]
        public string ModifiedBy { get; set; }
        [Column("InitialROEOffer_OfferDate")]
        public DateTimeOffset? InitialRoeofferOfferDate { get; set; }
        [Column("InitialROEOffer_OfferAmount")]
        public double? InitialRoeofferOfferAmount { get; set; }
        [Column("InitialROEOffer_OfferNotes")]
        public string InitialRoeofferOfferNotes { get; set; }
        [Column("FinalROEOffer_OfferDate")]
        public DateTimeOffset? FinalRoeofferOfferDate { get; set; }
        [Column("FinalROEOffer_OfferAmount")]
        public double? FinalRoeofferOfferAmount { get; set; }
        [Column("FinalROEOffer_OfferNotes")]
        public string FinalRoeofferOfferNotes { get; set; }
        [Column("InitialOptionOffer_OfferDate")]
        public DateTimeOffset? InitialOptionOfferOfferDate { get; set; }
        [Column("InitialOptionOffer_OfferAmount")]
        public double? InitialOptionOfferOfferAmount { get; set; }
        [Column("InitialOptionOffer_OfferNotes")]
        public string InitialOptionOfferOfferNotes { get; set; }
        [Column("FinalOptionOffer_OfferDate")]
        public DateTimeOffset? FinalOptionOfferOfferDate { get; set; }
        [Column("FinalOptionOffer_OfferAmount")]
        public double? FinalOptionOfferOfferAmount { get; set; }
        [Column("FinalOptionOffer_OfferNotes")]
        public string FinalOptionOfferOfferNotes { get; set; }
        [Column("InitialEasementOffer_OfferDate")]
        public DateTimeOffset? InitialEasementOfferOfferDate { get; set; }
        [Column("InitialEasementOffer_OfferAmount")]
        public double? InitialEasementOfferOfferAmount { get; set; }
        [Column("InitialEasementOffer_OfferNotes")]
        public string InitialEasementOfferOfferNotes { get; set; }
        [Column("FinalEasementOffer_OfferDate")]
        public DateTimeOffset? FinalEasementOfferOfferDate { get; set; }
        [Column("FinalEasementOffer_OfferAmount")]
        public double? FinalEasementOfferOfferAmount { get; set; }
        [Column("FinalEasementOffer_OfferNotes")]
        public string FinalEasementOfferOfferNotes { get; set; }
        [StringLength(40)]
        public string ParcelStatusCode { get; set; }
        [StringLength(40)]
        public string RoeStatusCode { get; set; }
        public bool IsActive { get; set; }
        [Column("Landowner_Score")]
        public int? LandownerScore { get; set; }
        [Column("Tracking_Number")]
        [StringLength(100)]
        public string TrackingNumber { get; set; }
        [Column("Parcel_ParcelId")]
        public Guid? ParcelParcelId { get; set; }
        public bool IsDeleted { get; set; }

        [ForeignKey(nameof(ParcelParcelId))]
        [InverseProperty(nameof(Parcel.InverseParcelParcel))]
        public virtual Parcel ParcelParcel { get; set; }
        [ForeignKey(nameof(ParcelStatusCode))]
        [InverseProperty(nameof(ParcelStatus.Parcel))]
        public virtual ParcelStatus ParcelStatusCodeNavigation { get; set; }
        [ForeignKey(nameof(RoeStatusCode))]
        [InverseProperty(nameof(RoeStatus.Parcel))]
        public virtual RoeStatus RoeStatusCodeNavigation { get; set; }
        [InverseProperty("Parcel")]
        public virtual ICollection<Appraisal> Appraisal { get; set; }
        [InverseProperty(nameof(Parcel.ParcelParcel))]
        public virtual ICollection<Parcel> InverseParcelParcel { get; set; }
        [InverseProperty("Parcel")]
        public virtual ICollection<Ownership> Ownership { get; set; }
        [InverseProperty("ParcelParcel")]
        public virtual ICollection<ParcelContactInfo> ParcelContactInfo { get; set; }
        [InverseProperty("ParcelParcel")]
        public virtual ICollection<ParcelContactLogs> ParcelContactLogs { get; set; }
        [InverseProperty("ParcelParcel")]
        public virtual ICollection<ParcelDocuments> ParcelDocuments { get; set; }
        [InverseProperty("ParcelParcel")]
        public virtual ICollection<RoeConditions> RoeConditions { get; set; }
        [InverseProperty("ParentParcel")]
        public virtual ICollection<StatusActivity> StatusActivity { get; set; }
    }
}
