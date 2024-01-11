using Relocation.DAL;
using System.Data.Entity;

namespace ROWM.Dal
{
    public class RelocationContext : DbContext
    {
        #region ctor
        public RelocationContext() : base("name=ROWM_Context") { }
        public RelocationContext(string c = "name=ROWM_Context") : base(c) { }
        #endregion

        public DbSet<ParcelRelocation> Relocations { get; set; }
        public DbSet<RelocationCase> RelocationCases { get; set; }
        public DbSet<RelocationEligibilityActivity> RelocationEligibilities { get; set; }
        public DbSet<RelocationDisplaceeActivity> RelocationActivities { get; set; }

        // types
        public DbSet<RelocationActivityType> RelocationActivity_Type { get; set; }


        // imported core entity
        public DbSet<Parcel> Parcels { get; set; }
        public DbSet<Document> Documents { get; set; }
        public DbSet<ContactLog> ContactLogs { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ParcelRelocation>()
                .HasRequired(r => r.Parcel);        
                
            modelBuilder.Entity<RelocationCase>()
                .HasMany(rx => rx.Logs)
                .WithMany()
                .Map(cx => {
                    cx.MapLeftKey("RelocationCaseId");
                    cx.MapRightKey("ContactLogId");
                    cx.ToTable("Relocation_Case_ContactLogs", "Austin");
                });

            modelBuilder.Entity<RelocationCase>()
                .HasMany(rx => rx.Documents)
                .WithMany()
                .Map(cx => {
                    cx.MapLeftKey("RelocationCaseId");
                    cx.MapRightKey("DocumentId");
                    cx.ToTable("Relocation_Case_Documents", "Austin");
                });


            #region override mapping conventions
            modelBuilder.Entity<Agent>()
                .HasMany(e => e.ContactLog)
                .WithRequired(e => e.Agent)
                .HasForeignKey(e => e.ContactAgentId);

            modelBuilder.Entity<ContactLog>()
                .HasMany(e => e.Parcel)
                .WithMany(e => e.ContactLog)
                .Map(m => m.ToTable("ParcelContactLogs", "ROWM"));

            modelBuilder.Entity<Owner>()
                .HasMany(e => e.ContactLog)
                .WithOptional(e => e.Owner)
                .HasForeignKey(e => e.Owner_OwnerId);

            modelBuilder.Entity<Document>()
                .HasMany(e => e.DocumentActivity)
                .WithOptional(e => e.Document)
                .HasForeignKey(e => e.ChildDocumentId);

            modelBuilder.Entity<Document>()
                .HasMany(e => e.DocumentActivity1)
                .WithRequired(e => e.Document1)
                .HasForeignKey(e => e.ParentDocumentId);

            modelBuilder.Entity<Document>()
                .HasMany(e => e.Owner)
                .WithMany(e => e.Document)
                .Map(m => m.ToTable("OwnerDocuments", "ROWM"));

            modelBuilder.Entity<Document>()
                .HasMany(e => e.Parcel)
                .WithMany(e => e.Document)
                .Map(m => m.ToTable("ParcelDocuments", "ROWM"));

            modelBuilder.Entity<DocumentPackage>()
                .HasMany(e => e.Document)
                .WithOptional(e => e.DocumentPackage)
                .HasForeignKey(e => e.DocumentPackage_PackageId);
            #endregion
        }
    }
}
