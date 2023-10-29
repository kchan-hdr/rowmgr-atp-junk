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
                    cx.ToTable("Relocation_Case_Document", "Austin");
                });
        }
    }
}
