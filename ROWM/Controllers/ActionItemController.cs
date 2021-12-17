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

        public ActionItemController(IActionItemRepository ownerRepository, Notification n) => (_repo, _notify) = (ownerRepository, n);


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
            var items = await _repo.AddActionItem(pid, a, req.ActivityDate);

            try
            {
                await _notify.SendNotification(a.ActionItemId, Notification.NotificationType.New);
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
            item = await _repo.UpdateActionItem(a, req.ActivityDate);

            try
            {
                if (item.Status != ActionStatus.Completed)
                {
                    var t = item.Status == ActionStatus.Canceled ? Notification.NotificationType.Cancel : Notification.NotificationType.Update;
                    await _notify.SendNotification(a.ActionItemId, t);
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
                DueDate = this.DueDate ?? DateTimeOffset.MaxValue,
                ParcelId = this.ParcelId,
                Status = (ActionStatus) Enum.Parse(typeof(ActionStatus), this.StatusCode)
            };
        }

        public ActionItem ToActionItem(ActionItem original)
        {
            original.Action = this.Action;
            original.AssignedGroupId = this.AssignedGroupId;
            original.DueDate = this.DueDate.HasValue ? this.DueDate.Value : DateTimeOffset.MaxValue;
            original.ParcelId = this.ParcelId;
            original.Status = (ActionStatus)Enum.Parse(typeof(ActionStatus), this.StatusCode);

            return original;
        }
    }
    #endregion
}
