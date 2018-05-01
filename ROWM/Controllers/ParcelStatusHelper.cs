﻿using ROWM.Dal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Controllers
{
    public class ParcelStatusHelper
    {
        readonly List<Parcel_Status> _Status;
        readonly List<Roe_Status> _roeStatus;

        public ParcelStatusHelper(ROWM_Context c)
        {
            _Status = c.Parcel_Status.AsNoTracking().ToList();
            _roeStatus = c.Roe_Status.AsNoTracking().ToList();
        }

        public string ParseDomainValue(int d)
        {
            var s = _Status.Single(sx => sx.DomainValue == d);
            return s.Code;
        }

        public int GetDomainValue(string status)
        {
            var s = _Status.Single(sx => sx.Code.Equals(status));
            return s.DomainValue;
        }

        public int GetRoeDomainValue(string status)
        {
            var s = _roeStatus.Single(sx => sx.Code.Equals(status));
            return s.DomainValue;
        }

        public static bool HasNoContact(Parcel parcel) => parcel.Parcel_Status.DomainValue <= 0;
    }
}
