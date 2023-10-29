using ExpenseTracking.Dal.Entities;
using System.Data.Entity;

namespace ExpenseTracking.Dal
{
    public class ExpenseContext : DbContext
    {
        #region ctor
        public ExpenseContext() : base("name=ROWM_Context") { }
        public ExpenseContext(string c = "name=ROWM_Context") : base(c) { }
        #endregion

        public virtual DbSet<Expense> Expense { get; set; }
        public virtual DbSet<ExpenseType> ExpenseType { get; set; }
        public virtual DbSet<ExpenseCategory> ExpenseCategories { get; set; }

        //// TODO: imported core entity

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            //modelBuilder.Entity<Expense>()
            //    .HasMany(e => e.ExpenseChildActivities)
            //    .WithOne(ea => ea.ChildExpense)
            //    .HasForeignKey(e => e.ChildExpenseId);

            //modelBuilder.Entity<Expense>()
            //    .HasMany(e => e.ExpenseParentActivities)
            //    .WithOne(ea => ea.ParentExpense)
            //    .HasForeignKey(e => e.ParentExpenseId)
            //    .IsRequired();

            modelBuilder.Entity<ExpenseType>()
                .HasMany(e => e.Expenses)
                .WithRequired(e => e.ExpenseType);

            modelBuilder.Entity<ExpenseCategory>()
                .HasMany(e => e.ExpenseTypes)
                .WithRequired(e => e.Category);
        }
    }
}