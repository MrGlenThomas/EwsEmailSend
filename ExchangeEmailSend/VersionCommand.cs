
namespace ExchangeEmailSend
{
    using System;
    using System.Diagnostics;
    using System.Reflection;

    using ManyConsole;

    internal class VersionCommand : ConsoleCommand
    {
        public VersionCommand()
        {
            IsCommand("Version", "Get the current version");
        }

        public override int Run(string[] remainingArguments)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            var version = fvi.FileVersion;

            Console.WriteLine(version);

            return 0;
        }
    }
}
