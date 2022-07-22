using geographia.ags;
using ROWM.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Dal
{
    public class UpdateParcelStatus_austin : IUpdateParcelStatus
    {
        public Agent myAgent { get; set; }
        public IEnumerable<Parcel> myParcels { get; set; }
        public DateTimeOffset StatusChangeDate { get; set; } = DateTimeOffset.Now;
        public string AcquisitionStatus { get; set; }
        public string RoeStatus { get; set; }
        public string RoeCondition { get; set; }
        public string Notes { get; set; }
        public string ModifiedBy { get; set; } = "UP";

        readonly ROWM_Context _ctx;
        readonly IParcelStatusHelper _statusHelper;
        readonly IFeatureUpdate _featureUpdate;
        public UpdateParcelStatus_austin(ROWM_Context c, IParcelStatusHelper h, IFeatureUpdate f) => (_ctx, _statusHelper, _featureUpdate) = (c, h, f);

        public async Task<int> Apply()
        {
            _ = myParcels ?? throw new ArgumentNullException("parcels");

            var touched = false;

            foreach(var p in myParcels)
            {
                if (await Apply(p))
                    touched = true;
            }

            if (touched)
            {
                return await _ctx.SaveChangesAsync();
            }

            return 0;
        }

        // apply to each parcel
        async Task<bool> Apply(Parcel p) =>
            await HandleEngagement(p) || await HandleRoe (p);

        async Task<bool> HandleEngagement(Parcel p)
        {
            _ = p ?? throw new ArgumentNullException("parcel");

            // Not Contacted - default. neither Letter nor Contact log
            var hasLetter = p.Document.Any(px => px.DocumentType == "Engagement-Community-Engagement-Letter");
            var hasContact = p.ContactLog.Any(px => px.ProjectPhase == "Community Engagement");

            if (hasLetter == false && hasContact == false)
            {
                p.OutreachStatusCode = "Not_Contacted";
                return false;
            }

            string newCode = "Not_Contacted";
            if (hasLetter)
                newCode = "Owner_Letter_Sent";

            if (hasContact)
                newCode = "Owner_Meeting";
            //// Otherwise
            //// Action Pending - has a pending action item
            //// No current action - no open action
            //var hasOpenAction = p.ActionItems.Any(ax => ax.Status == ActionStatus.Pending);

            //var newCode = hasOpenAction ? "Action_Required" : "No_Action";

            if (p.OutreachStatusCode != newCode)
            {
                p.Activities.Add(new StatusActivity
                {
                    ActivityDate = StatusChangeDate,
                    OriginalStatusCode = p.OutreachStatusCode,
                    StatusCode = newCode,
                    ParentParcel = p,
                    Agent = myAgent
                });
                p.OutreachStatusCode = newCode;
                p.LastModified = DateTimeOffset.Now;

                var dv = _statusHelper.GetDomainValue(newCode);
                await _featureUpdate.UpdateFeatureOutreach(p.Assessor_Parcel_Number, p.Tracking_Number, dv, string.Empty, default);

                return true;
            }

            return false;
        }

        async Task<bool> HandleRoe(Parcel p)
        {
            _ = p ?? throw new ArgumentNullException("parcel");

            if (p.RoeStatusCode == "ROE_on_hold" || p.RoeStatusCode == "ROE_Denied")
                return false;

            // Not requested - neither Request doc nor Contact Log
            var hasRequest = p.Document.Any(px => px.DocumentType == "ROE-Right-of-Entry-Request-Package");
            var hasContact = p.ContactLog.Any(px => px.ProjectPhase == "ROE");

            if (hasRequest == false && hasContact == false)
            {
                p.RoeStatusCode = "No_Activity_roe";
                return false;
            }

            // Otherwise
            // ROE Requested
            var newCode = "ROE_Mailed";

            // Full Executed document
            var hasFull = p.Document.Any(px => px.DocumentType == "ROE-Full-Right-of-Entry-Executed");
            // Restricted document
            var hasRestrict = p.Document.Any(px => px.DocumentType == "ROE-Restricted-Right-of-Entry-Executed");

            // both, get the latest
            if (hasFull && hasRestrict)
            {
                var lastFull = p.Document.Where(px => px.DocumentType == "ROE-Full-Right-of-Entry-Executed").Max(px => px.DateRecorded);
                var lastRestrict = p.Document.Where(px => px.DocumentType == "ROE-Restricted-Right-of-Entry-Executed").Max(px => px.DateRecorded);

                p.RoeStatusCode = lastFull > lastRestrict ? "ROE_Full_Access" : "ROE_Partial_Access";
            }
            else
            {
                if (hasFull)
                    p.RoeStatusCode = "ROE_Full_Access";
                if (hasRestrict)
                    p.RoeStatusCode = "ROE_Partial_Access";
            }


            if (p.RoeStatusCode != newCode)
            {
                p.Activities.Add(new StatusActivity
                {
                    ActivityDate = StatusChangeDate,
                    OriginalStatusCode = p.RoeStatusCode,
                    StatusCode = newCode,
                    ParentParcel = p,
                    Agent = myAgent
                });
                p.RoeStatusCode = newCode;
                p.LastModified = DateTimeOffset.Now;

                var dv = _statusHelper.GetDomainValue(newCode);
                await _featureUpdate.UpdateFeatureRoe(p.Assessor_Parcel_Number, p.Tracking_Number, dv);

                return true;
            }

            return false;
        }
    }
}
