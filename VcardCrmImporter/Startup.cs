using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(VcardCrmImporter.Startup))]
namespace VcardCrmImporter
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
