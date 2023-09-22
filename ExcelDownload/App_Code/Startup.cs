using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExcelDownload.Startup))]
namespace ExcelDownload
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
