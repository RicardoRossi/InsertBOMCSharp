using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsertBOMCSharp
{
    internal static class Helper
    {
        internal static string[] getCADFilesFromDirectory(string directory)
        {
            // check if diretory exixsts
            if (! Directory.Exists(directory))
            {
                return null;
            }

            string[] parts = Directory.GetFiles(directory,"*.sldprt");
            string[] assemblies = Directory.GetFiles(directory, "*.sldasm");

            List<string> files = new List<string>();
            if (parts != null)
            {
                files.AddRange(parts);
            }
            if (assemblies!=null)
            {
                files.AddRange(assemblies);
            }
            return files.ToArray();
        }

    }
}
