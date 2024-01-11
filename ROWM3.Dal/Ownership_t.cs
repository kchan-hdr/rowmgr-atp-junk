namespace ROWM.Dal
{
    public partial class Ownership
    {
        public enum OwnershipType { Primary = 1, Related, Relinquished = 100, Correction = 200 };

        public bool IsPrimary() => this.Ownership_t == (int) OwnershipType.Primary;
        public bool IsCurrentOwner => this.Ownership_t < (int)OwnershipType.Relinquished;
    }
}
