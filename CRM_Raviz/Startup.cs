using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(CRM_Raviz.Startup))]
namespace CRM_Raviz
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
