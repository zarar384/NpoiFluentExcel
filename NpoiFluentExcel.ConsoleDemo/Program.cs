using NpoiFluentExcel.ConsoleDemo.Scenarios;

Console.WriteLine("Select demo:");
Console.WriteLine("1 - Basic");
Console.WriteLine("2 - Advanced");
Console.WriteLine("3 - Mapping");

var key = Console.ReadLine();

switch (key)
{
    case "1":
        BasicDemo.Run();
        break;

    case "2":
        AdvancedDemo.Run();
        break;

    case "3":
        MappingDemo.Run();
        break;

    default:
        Console.WriteLine("Unknown option");
        break;
}