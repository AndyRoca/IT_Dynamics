using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using SP.Plugins;
using SP.Plugins.Articulos_de_conocimiento;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;

namespace IT.SP.Plugins
{
    [TestClass]
    public class IT_CrearArticulosDeConocimiento
    {
        private IOrganizationService Crm;

        XrmFakedContext FakedContext { get; set; }

        [TestInitialize]
        public void Setup()
        {
            FakedContext = new XrmFakedContext();
            var s = new List<string>();
            FakedContext.ProxyTypesAssembly = Assembly.Load(Assembly.GetExecutingAssembly().GetReferencedAssemblies().Where(a => a.Name == "SP.Plugins").FirstOrDefault());
            var mocks = new List<Entity>()
            {
            };
            FakedContext.Initialize(mocks);
            Crm = FakedContext.GetOrganizationService();
        }

        [TestMethod]
        public void CrearArticulosDeConocimiento()
        {
            Crm = new Crm(ConfigurationManager.ConnectionStrings["Dyn365OnPremise"].ConnectionString).Service;
            FakedContext = new XrmRealContext(Crm);
            var pluginContext = new XrmFakedPluginExecutionContext
            {
                InputParameters = new ParameterCollection { ["Target"] = new Entity("knowledgearticle") },
                MessageName = "Create",
            };
            FakedContext.ExecutePluginWith(pluginContext, new CreateKnowledgeArticle());
        }


        [TestCleanup]
        public void CleanUp()
        {
            FakedContext = null;
        }
    }
}
