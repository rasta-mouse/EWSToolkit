using Microsoft.Exchange.WebServices.Data;

using System.Globalization;
using System.IO;
using System.Linq;

namespace Toolkit
{
    public class HomeFolder
    {
        ExchangeService Service;
        ExtendedPropertyDefinition FolderPropInfo;
        Folder TargetFolder;

        public HomeFolder(ExchangeConnection connection)
        {
            Service = connection.ExchangeService;
            FolderPropInfo = new ExtendedPropertyDefinition(14047, MapiPropertyType.Binary);
        }

        public void SetTargetFolder(FolderName folder)
        {
            TargetFolder = Folder.Bind(Service, (WellKnownFolderName)folder);
        }

        public void SetFolderUrl(string url)
        {
            var encodedUrl = EncodeUrl(url);

            TargetFolder.SetExtendedProperty(FolderPropInfo, encodedUrl);
        }

        public void Update()
        {
            TargetFolder.Update();
        }

        private byte[] EncodeUrl(string url)
        {
            var writer = new StringWriter();
            var dataSize = ((ConvertToHex(url).Length / 2) + 2).ToString("X2");

            writer.Write("02");
            writer.Write("00000001");
            writer.Write("00000001");
            writer.Write("00000000000000000000000000000000000000000000000000000000");
            writer.Write("000000");
            writer.Write(dataSize);
            writer.Write("000000");
            writer.Write(ConvertToHex(url));
            writer.Write("0000");

            var buffer = HexStringToByteArray(writer.ToString());
            return buffer;
        }

        private string ConvertToHex(string input)
        {
            return string.Join(string.Empty, input.Select(c => ((int)c).ToString("x2") + "00").ToArray());
        }

        private byte[] HexStringToByteArray(string input)
        {
            return Enumerable
                .Range(0, input.Length / 2)
                .Select(index => byte.Parse(input.Substring(index * 2, 2), NumberStyles.AllowHexSpecifier)).ToArray();
        }
    }
}