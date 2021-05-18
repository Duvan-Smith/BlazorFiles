using Blazor.FileReader;
using MatBlazor;
using Microsoft.AspNetCore.Blazor.Hosting;
using System.Threading.Tasks;

namespace BlazorFiles.Client
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var builder = WebAssemblyHostBuilder.CreateDefault(args);
            builder.Services.AddMatBlazor();
            builder.Services.AddFileReaderService(options =>
            {
                options.UseWasmSharedBuffer = true;
            });

            builder.RootComponents.Add<App>("app");

            await builder.Build().RunAsync();
        }
    }
}
