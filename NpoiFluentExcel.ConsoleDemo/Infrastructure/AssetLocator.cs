namespace NpoiFluentExcel.ConsoleDemo.Infrastructure
{
    public static class AssetLocator
    {
        public static string ResolveLogo()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(baseDir, "Resources", "company-logo.png");
        }
    }
}
