using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace XLSConverter
{
    public static class PathOperations
    {
        /// <summary>
        /// If the path has a root (e.g. C:\) does nothing. Otherwise
        /// returns a path with the current execuatable's location as
        /// root.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string VerifyRootedPath(string path)
        {
            // Get relative path. Remove "file:\" prefix
            string baseDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).Remove(0, 6);

            if (!Path.IsPathRooted(path))
            {
                return Path.Combine(baseDir, path);
            }
            else
            {
                return path;
            }
        }

    }
}
