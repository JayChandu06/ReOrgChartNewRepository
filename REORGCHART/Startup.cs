using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(REORGCHART.Startup))]
namespace REORGCHART
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
