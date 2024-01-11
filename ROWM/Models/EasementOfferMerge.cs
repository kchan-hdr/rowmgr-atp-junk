using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ROWM.Dal;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.IO;

namespace ROWM.Models
{
    public class EasementOfferMerge
    {
        //static readonly Regex mergefield = new Regex(@"MERGEFIELD\s+(?<key>.+?)\s+\\* MERGEFORMAT");
        static readonly Regex mergefield = new Regex("MERGEFIELD (?<key>.+)");

        internal static WordprocessingDocument DoMerge(WordprocessingDocument dox, Dictionary<string, string> dto)
        {
            // merge simple fields
            var fields = dox.MainDocumentPart.Document.Descendants<SimpleField>();

            Debug.WriteLine("matched fields" + fields.Count());

            foreach (var field in fields)
            {
                var key = field.Instruction.Value.Trim();

                var m = mergefield.Match(key).Groups["key"];
                if (!m.Success)
                {
                    Trace.TraceWarning(key);
                    continue;
                }

                Debug.WriteLine("value to replace" + dto[m.Value]);

                var r = field.GetFirstChild<Run>();

                var t = r.GetFirstChild<Text>();
                r.ReplaceChild<Text>(new Text(dto[m.Value]), t);
            }

            return dox;
        }        
    }
 }
