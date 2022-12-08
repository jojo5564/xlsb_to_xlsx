using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsb_to_xlsx
{
    public sealed class AsposeLicenseHelper
    {
        public static void SetCellLicense()
        {
            var txtLicense = File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + "License.lic");
            var memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(txtLicense));
            new Aspose.Cells.License().SetLicense(memoryStream);
        }
    }
}
