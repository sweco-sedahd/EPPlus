using System.Text.RegularExpressions;

namespace OfficeOpenXml.Utils
{
    public static class AddressUtility
    {
        public static string ParseEntireColumnSelections(string address)
        {
            string parsedAddress = address;
            MatchCollection matches = Regex.Matches(address, "[A-Z]+:[A-Z]+");
            foreach (Match match in matches)
            {
                AddRowNumbersToEntireColumnRange(ref parsedAddress, match.Value);
            }

            return parsedAddress;
        }

        private static void AddRowNumbersToEntireColumnRange(ref string address, string range)
        {
            string parsedRange = string.Format("{0}{1}", range, ExcelPackage.MaxRows);
            string[] splitArr = parsedRange.Split(new[] { ':' });
            address = address.Replace(range, string.Format("{0}1:{1}", splitArr[0], splitArr[1]));
        }
    }
}