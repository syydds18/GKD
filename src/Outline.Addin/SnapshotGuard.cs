using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Outline.Core;

namespace Outline.Addin
{
    internal static class SnapshotGuard
    {
        // Include coordinates, types and formulas. Refuse restore after any data edit;
        // a conservative refusal is safer than restoring group ranges at stale row indices.
        internal static string Fingerprint(SheetData data)
        {
            using(var memory=new MemoryStream())
            using(var writer=new BinaryWriter(memory,Encoding.UTF8,true))
            {
                writer.Write(data.MaxRow);writer.Write(data.MaxCol);
                foreach(var row in data.Values.OrderBy(x=>x.Key))
                foreach(var cell in row.Value.OrderBy(x=>x.Key))
                {
                    writer.Write(row.Key);writer.Write(cell.Key);
                    writer.Write(cell.Value==null?"null":cell.Value.GetType().FullName);
                    writer.Write(Convert.ToString(cell.Value,CultureInfo.InvariantCulture)??"");
                }
                writer.Write(-1);
                foreach(var row in data.Formulas.OrderBy(x=>x.Key))
                foreach(var cell in row.Value.OrderBy(x=>x.Key))
                {writer.Write(row.Key);writer.Write(cell.Key);writer.Write(cell.Value??"");}
                writer.Flush();
                using(var hash=SHA256.Create())return Convert.ToBase64String(hash.ComputeHash(memory.ToArray()));
            }
        }
    }
}
