namespace NpoiFluentExcel.ConsoleDemo.Domain
{
    public class AccountOwner
    {
        public string CompanyName { get; set; }
        public string AddressLine { get; set; }
        public string Identifier { get; set; }

        public string FormatHeader()
        {
            return $"{CompanyName}\n" +
                   $"{AddressLine}\n" +
                   $"ID: {Identifier}";
        }
    }
}
