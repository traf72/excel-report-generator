using ExcelReportGenerator.Attributes;
using ExcelReportGenerator.Rendering.Panels;
using System;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelReportGenerator
{
    internal class Licensing
    {
        public void CheckLicence()
        {
            Type[] keyPartTypes = typeof(IPanel).Assembly.GetExportedTypes().Where(t => t.IsDefined(typeof(LicenceKeyPartAttribute), false)).ToArray();
            string key = GetKey(keyPartTypes);
        }

        private string GetKey(Type[] keyPartTypes)
        {
            var key = new StringBuilder();
            foreach (Type keyPartType in keyPartTypes)
            {
                var licencePartAttr = keyPartType.GetCustomAttribute<LicenceKeyPartAttribute>();
                int startIndex = (int)(keyPartType.Name.Length * .25);
                int endIndex = (int)(keyPartType.Name.Length * .75);
                string part = keyPartType.Name.Substring(startIndex, endIndex - startIndex);
                if (licencePartAttr.L)
                {
                    part = Reverse(part);
                }
                if (licencePartAttr.U)
                {
                    part = part.ToLower();
                }
                if (licencePartAttr.R)
                {
                    part = part.ToUpper();
                }

                key.Insert((int)(key.Length * .47), part);
            }

            return key.ToString();
        }

        private static string Reverse(string str)
        {
            char[] charArray = str.ToCharArray();
            char[] result = new char[charArray.Length];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = charArray[charArray.Length - (i + 1)];
            }

            return string.Join(string.Empty, result);
        }
    }
}