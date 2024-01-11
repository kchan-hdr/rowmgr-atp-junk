using DocumentFormat.OpenXml.Packaging;
using ROWM.Dal;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace ROWM.Models
{
    public class EasementOfferHelper
    {
        internal async Task<byte[]> GeneratePackage(Parcel p)
        {
            var dot = EasementOffer_Dto.Cast(p);            

            var s = Assembly.GetExecutingAssembly().GetManifestResourceStream("ROWM.ReportTemplates.Basin_easement_offer_package_template.docx");
            
            if (s == null)
                throw new ApplicationException("Cannot find report template");

            MemoryStream working = new MemoryStream();
            await s.CopyToAsync(working);

            var dox = WordprocessingDocument.Open(working, true);

            EasementOfferMerge.DoMerge(dox, dot);

            dox.MainDocumentPart.Document.Save();
            dox.Close();

            return working.ToArray();            

        }
    }
}
