﻿using Microsoft.AspNetCore.Mvc;
using ROWM.ActionItemNotification;
using ROWM.Dal;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Controllers
{
    [Produces("application/json")]
    [Route("api/v2")]
    public class ActionItemController : ControllerBase
    {
        readonly IActionItemRepository _repo;
        readonly Notification _notify;
        readonly IUpdateParcelStatus _statusUpdater;

        public ActionItemController(IActionItemRepository ownerRepository, Notification n, IUpdateParcelStatus u) => (_repo, _notify, _statusUpdater) = (ownerRepository, n, u);


        [HttpGet("parcels/{pid}/actionitems")]
        [ProducesDefaultResponseType(typeof(IEnumerable<ActionItem_dto>))]
        public async Task<IActionResult> GetActionItems(string pid)
        { 
            var items = await _repo.GetActionItems(pid) ?? throw new IndexOutOfRangeException($"bad parcel {pid}");

            if (items.Any())
                return new JsonResult(items.Select(ix => new ActionItem_dto(ix)));

            return NoContent();
        }

        [HttpPost("parcels/{pid}/actionitems")]
        public async Task<IEnumerable<ActionItem_dto>> AddActionItem(string pid, [FromBody]ActionItem_Request req)
        {
            var a = req.ToActionItem();
            var agent = await req.FindAgent(_repo);
            var items = await _repo.AddActionItem(pid, a, req.ActivityDate, agent?.AgentId ?? Guid.Empty);

            try
            {
                var notif = _notify.SendNotification(a.ActionItemId, Notification.NotificationType.New);
                var upd = _statusUpdater.Apply();

                await Task.WhenAll(new Task[] { notif, upd });
            }
            catch (Exception e)
            {
                Trace.TraceError(e.Message);
            }

            var r = items.Select(ix => new ActionItem_dto(ix));
            return r;
        }

        [HttpGet("actionitems/{itemId:Guid}")]
        public async Task<ActionItem_dto> GetActionItem(Guid itemId) => new ActionItem_dto( await _repo.GetActionItem(itemId));

        [HttpPut("actionitems/{itemId:Guid}")]
        public async Task<IActionResult> UpdateActionItem(Guid itemId, [FromBody]ActionItem_Request req)
        {
            var item = await _repo.GetActionItem(req.ActionItemId ?? Guid.Empty);
            if (item == null)
                return BadRequest();

            var a = req.ToActionItem(item);
            var agent = await req.FindAgent(_repo);
            item = await _repo.UpdateActionItem(a, req.ActivityDate, agent?.AgentId ?? Guid.Empty);

            try
            {
                if (item.Status != ActionStatus.Completed)
                {
                    var t = item.Status == ActionStatus.Canceled ? Notification.NotificationType.Cancel : Notification.NotificationType.Update;
                    var notif = _notify.SendNotification(a.ActionItemId, t);
                    var upd = _statusUpdater.Apply();

                    await Task.WhenAll(new Task[] { notif, upd });
                }
            }
            catch (Exception e)
            {
                Trace.TraceError(e.Message);
            }

            return new JsonResult(new ActionItem_dto(item));
        }
    }

    #region dto
    public class ActionItem_dto
    {
        public Guid? ActionItemId { get; set; }
        public Guid? ParcelId { get; set; }
        public Guid AssignedGroupId { get; set; }
        public string Assigned { get; set; }
        public string Action { get; set; }
        public DateTimeOffset? DueDate { get; set; }
        public DateTimeOffset Created { get; set; }
        public DateTimeOffset LastModified { get; set; }
        public string StatusCode { get; set; }

        public ActionItem_dto() { }
        public ActionItem_dto(ActionItem a)
        {
            ActionItemId = a.ActionItemId;
            ParcelId = a.ParcelId.Value;
            AssignedGroupId = a.AssignedGroupId.Value;
            Assigned = a.AssignedGroup?.GroupNameCaption ?? string.Empty;
            Action = a.Action;
            DueDate = a.DueDate;
            StatusCode = Enum.GetName(typeof(ActionStatus), a.Status);
            Created = a.Created;
            LastModified = a.LastModified;
        }
    }

    public class ActionItem_Request : ActionItem_dto
    {
        public DateTimeOffset ActivityDate { get; set; }
        public string AgentName { get; set; }
        public string CreatorEmail { get; set; }

        public ActionItem ToActionItem()
        {
            return new ActionItem
            {
                Action = this.Action,
                AssignedGroupId = this.AssignedGroupId,
                DueDate = this.DueDate,
                ParcelId = this.ParcelId,
                Status = (ActionStatus) Enum.Parse(typeof(ActionStatus), this.StatusCode)
            };
        }

        public ActionItem ToActionItem(ActionItem original)
        {
            original.Action = this.Action;
            original.AssignedGroupId = this.AssignedGroupId;
            original.DueDate = this.DueDate;
            original.ParcelId = this.ParcelId;
            original.Status = (ActionStatus)Enum.Parse(typeof(ActionStatus), this.StatusCode);

            return original;
        }

        #region match agent
        /// <summary>
        /// this is a temporary workaround until we upgrade to Core
        /// </summary>
        /// <param name="email"></param>
        /// <returns></returns>
        internal async Task<Agent> FindAgent(IActionItemRepository _repo)
        {
            if (string.IsNullOrEmpty(this.CreatorEmail))
                return default;

            var allAgents = await _repo.GetAgents();
            var a = allAgents.FirstOrDefault(ax => ax.AADObjectId.ToString().Equals(this.AgentName));

            if (a == null)
                return default;

            if (string.IsNullOrEmpty(a.AgentEmail))
            {
                a.AgentEmail = this.CreatorEmail;
                a = await _repo.UpdateAgent(a);
            }

            if ( !a.AgentEmail.Equals(this.CreatorEmail))
            {
                Trace.TraceInformation($"agent has different emails {a.AgentEmail} from AAD {this.CreatorEmail}");
                a.AgentEmail = this.CreatorEmail;
                a = await _repo.UpdateAgent(a);
            }

            return a;
        }
        #endregion

    }
    #endregion


}
