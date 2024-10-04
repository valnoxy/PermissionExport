using Spectre.Console;
using System.IO;
using System.Security.Principal;
using System.Security.AccessControl;

namespace PermissionExport
{
    internal class Program
    {
        public class Collection
        {
            public string? Ordner { get; set; }
            public string? BenutzerUndGruppen { get; set; }
            public string? Berechtigung { get; set; }
            public string? Zugriffstyp { get; set; }
            public string? Vererbbar { get; set; }
        }

        private static void Main(string[] args)
        {
            var collectionList = new List<Collection>();

            if (args.Length == 0)
            {
                AnsiConsole.MarkupLine("[bold red]FAIL[/]: No directory provided!");
                AnsiConsole.MarkupLine("[grey]Press any key to exit.[/]");
                Console.ReadLine();
                Environment.Exit(1);
            }

            var directory = args[0];
            if (!Directory.Exists(directory))
            {
                AnsiConsole.MarkupLine($"[bold red]FAIL[/]: Directory '{directory}' does not exist!");
                AnsiConsole.MarkupLine("[grey]Press any key to exit.[/]");
                Console.ReadLine();
                Environment.Exit(1);
            }

            AnsiConsole.Status()
               .Start("Working ...", ctx =>
               {
                   ctx.Spinner(Spectre.Console.Spinner.Known.Dots);
                   ctx.SpinnerStyle(Spectre.Console.Style.Parse("green"));

                   ctx.Status("Gathering data ...");
                   AnsiConsole.MarkupLine("[grey bold]INFO:[/] Collecting all information ...");
                   try
                   {
                       var directoryInfo = new DirectoryInfo(directory);
                       var directories = new List<DirectoryInfo>();

                       try
                       {
                           directories.AddRange(directoryInfo.GetDirectories("*", SearchOption.TopDirectoryOnly));
                       }
                       catch (UnauthorizedAccessException)
                       {
                           AnsiConsole.MarkupLine($"[bold yellow]WARN[/]: Access denied on directory '{directoryInfo.FullName}'.");
                       }

                       foreach (var subDirectory in directoryInfo.GetDirectories())
                       {
                           try
                           {
                               directories.AddRange(subDirectory.GetDirectories("*", SearchOption.AllDirectories));
                           }
                           catch (UnauthorizedAccessException)
                           {
                               AnsiConsole.MarkupLine($"[bold yellow]WARN[/]: Access denied on directory '{subDirectory.FullName}'.");
                           }
                       }

                       ctx.Status("Processing data ...");
                       foreach (var info in directories)
                       {
                           try
                           {
                               var security = info.GetAccessControl();
                               var rules = security.GetAccessRules(true, true, typeof(NTAccount));

                               foreach (var rule in rules)
                               {
                                   var fileSystemRule = rule as FileSystemAccessRule;
                                   if (fileSystemRule == null)
                                   {
                                       continue;
                                   }

                                   var collection = new Collection
                                   {
                                       Ordner = info.FullName,
                                       BenutzerUndGruppen = fileSystemRule.IdentityReference.Value,
                                       Berechtigung = fileSystemRule.FileSystemRights.ToString(),
                                       Zugriffstyp = fileSystemRule.AccessControlType.ToString(),
                                       Vererbbar = fileSystemRule.IsInherited ? "Ja" : "Nein"
                                   };

                                   collectionList.Add(collection);
                                   AnsiConsole.MarkupLine("[grey]Processed[/]: {0}", Markup.Escape($"{info}"));
                               }
                           }
                           catch (Exception ex)
                           {
                               AnsiConsole.MarkupLine("[bold red]Error[/]: An error has occurred while processing {0}.", Markup.Escape($"{info}"));

                               AnsiConsole.WriteException(ex,
                                   ExceptionFormats.ShortenPaths | ExceptionFormats.ShortenTypes |
                                   ExceptionFormats.ShortenMethods | ExceptionFormats.ShowLinks);
                           }
                       }
                   }
                   catch (Exception ex)
                   {
                       AnsiConsole.MarkupLine("[bold red]Error[/]: An error has occurred while processing data.");
                       AnsiConsole.WriteException(ex,
                           ExceptionFormats.ShortenPaths | ExceptionFormats.ShortenTypes |
                           ExceptionFormats.ShortenMethods | ExceptionFormats.ShowLinks);
                       AnsiConsole.MarkupLine("[grey]Press any key to exit.[/]");
                       Console.ReadLine();
                       Environment.Exit(1);
                   }

                   ctx.Status("Saving data ...");
                   AnsiConsole.MarkupLine("[grey bold]INFO:[/] Creating file in base directory ...");
                   var targetFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "collection.csv");
                   AnsiConsole.MarkupLine($"[grey bold]INFO:[/] Creating '{targetFile}'");
                   ExportToCsv(collectionList, targetFile);
                   AnsiConsole.MarkupLine("[green bold]INFO:[/] Data exported to file.");
               });

            // Completed
            AnsiConsole.MarkupLine("[bold green]Done[/]: Data successfully parsed.");
            AnsiConsole.MarkupLine("[grey]Press any key to exit.[/]");
            Console.ReadLine();
        }

        public static void ExportToCsv(IEnumerable<Collection> items, string filePath)
        {
            using var writer = new StreamWriter(filePath);
            writer.WriteLine("Ordner;Benutzer order Gruppe;Berechtigung;Zugriffstyp;Vererbbar");
            foreach (var item in items)
            {
                writer.WriteLine($"\"{item.Ordner}\";\"{item.BenutzerUndGruppen}\";\"{item.Berechtigung}\";\"{item.Zugriffstyp}\";\"{item.Vererbbar}\"");
            }
        }
    }
}
