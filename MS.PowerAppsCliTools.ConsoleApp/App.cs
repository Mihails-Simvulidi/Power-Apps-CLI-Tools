using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
using PC.PowerApps.Common.Entities.Dataverse;

namespace MS.PowerAppsCliTools.ConsoleApp
{
    internal class App : IDisposable
    {
        private const string DataverseConnectionString = "DataverseConnectionString";
        private const string ExportUnmanagedSolution = "exportunmanagedsolution";
        private const string UpdatePluginAssembly = "updatepluginassembly";
        private const string UpdateWebResource = "updatewebresource";
        private readonly string[] args;
        private bool disposedValue;
        private readonly Lazy<ServiceClient> serviceClient = new Lazy<ServiceClient>(() =>
        {
            Console.Write("Connecting to Dataverse... ");
            string? dataverseConnectionString = Environment.GetEnvironmentVariable(DataverseConnectionString);

            if (dataverseConnectionString == null)
            {
                throw new InvalidOperationException($"Environment variable {DataverseConnectionString} is not set.");
            }

            ServiceClient serviceClient = new ServiceClient(dataverseConnectionString);
            Console.WriteLine("Connected.");
            return serviceClient;
        });
        private readonly Lazy<ServiceContext> serviceContext;

        private ServiceClient ServiceClient => serviceClient.Value;
        private ServiceContext ServiceContext => serviceContext.Value;

        public App(string[] args)
        {
            this.serviceContext = new Lazy<ServiceContext>(() => new ServiceContext(ServiceClient));
            this.args = args;
        }

        public async Task RunAsync()
        {
            if (args.Length == 0)
            {
                ShowUsage();
                return;
            }

            var command = args[0].ToLower();
            switch (command)
            {
                case ExportUnmanagedSolution:
                    if (args.Length != 3)
                    {
                        ShowUsage();
                        break;
                    }

                    await ExportUnmanagedSolutionAsync();
                    break;

                case UpdatePluginAssembly:
                    if (args.Length != 2)
                    {
                        ShowUsage();
                        break;
                    }

                    await UpdatePluginAssemblyAsync();
                    break;

                case UpdateWebResource:
                    if (args.Length != 3)
                    {
                        ShowUsage();
                        break;
                    }

                    await UpdateWebResourceAsync();
                    break;

                default:
                    ShowUsage();
                    break;
            }
        }

        private void ShowUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine($"{AppDomain.CurrentDomain.FriendlyName} {ExportUnmanagedSolution} solution path");
            Console.WriteLine($"{AppDomain.CurrentDomain.FriendlyName} {UpdatePluginAssembly} path");
            Console.WriteLine($"{AppDomain.CurrentDomain.FriendlyName} {UpdateWebResource} prefix path");
        }

        async Task ExportUnmanagedSolutionAsync()
        {
            Console.WriteLine("Exporting solution...");
            var solutionName = args[1];
            var path = args[2];
            var exportSolutionRequest = new ExportSolutionRequest
            {
                Managed = false,
                SolutionName = solutionName,
            };
            var exportSolutionResponse = (ExportSolutionResponse)(await ServiceClient.ExecuteAsync(exportSolutionRequest));
            await File.WriteAllBytesAsync(path, exportSolutionResponse.ExportSolutionFile);
        }

        async Task UpdatePluginAssemblyAsync()
        {
            Console.WriteLine("Retrieving plug-in assembly...");
            var path = args[1];
            var fileInfo = new FileInfo(path);
            var fileNameWithoutExtension = fileInfo.Name.Substring(0, fileInfo.Name.Length - fileInfo.Extension.Length);
            var pluginAssembly = ServiceContext.PluginAssemblySet
                .Where(pa => pa.Name == fileNameWithoutExtension)
                .Single();

            Console.WriteLine("Updating plug-in assembly...");
            var updatePluginAssembly = new PluginAssembly
            {
                Id = pluginAssembly.Id,
                Content = await GetFileBase64(path),
            };
            await ServiceClient.UpdateAsync(updatePluginAssembly);
        }

        async Task UpdateWebResourceAsync()
        {
            Console.WriteLine("Retrieving web resource...");
            var prefix = args[1];
            var path = args[2];
            var fileInfo = new FileInfo(path);
            var webResource = ServiceContext.WebResourceSet
                .Where(wr => wr.Name == $"{prefix}{fileInfo.Name}")
                .Single();

            Console.WriteLine("Updating web resource...");
            var updateWebResource = new WebResource
            {
                Id = webResource.Id,
                Content = await GetFileBase64(path),
            };
            await ServiceClient.UpdateAsync(updateWebResource);

            Console.WriteLine("Publishing web resource...");
            var publishXmlRequest = new PublishXmlRequest
            {
                ParameterXml = $"<importexportxml><webresources><webresource>{webResource.Id}</webresource></webresources></importexportxml>",
            };
            await ServiceClient.ExecuteAsync(publishXmlRequest);
        }

        static async Task<string> GetFileBase64(string path)
        {
            var fileBytes = await File.ReadAllBytesAsync(path);
            var fileBase64 = Convert.ToBase64String(fileBytes);
            return fileBase64;
        }


        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (serviceContext.IsValueCreated)
                    {
                        ServiceContext.Dispose();
                    }

                    if (serviceClient.IsValueCreated)
                    {
                        ServiceClient.Dispose();
                    }
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
