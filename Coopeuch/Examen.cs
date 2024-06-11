using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Coopeuch
{
    public class Examen
    {

        /*
         * Agregar unsecure configuration 
         * <?xml version="1.0" encoding="utf-8"?>
            <Settings>
              <setting name="expirationDays">
                <value>10</value>
              </setting>
            </Settings>
         * */

        private XmlDocument _pluginConfiguration;
        public Examen(string unsecureConfig, string secureConfig)
        {
            if (string.IsNullOrEmpty(unsecureConfig))
            {
                throw new InvalidPluginExecutionException("Unsecure configuration missing.");
            }
            _pluginConfiguration = new XmlDocument();
            _pluginConfiguration.LoadXml(unsecureConfig);
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracingService.Trace("n Proceso");
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    Entity entity = (Entity)context.InputParameters["Target"];

                    tracingService.Trace("Obtención de contexto");

                    //Obtenemos la cantidad de dias para la expiración
                    tracingService.Trace("Obtención de datos");
                    var sExpirationDays = GetValueNode(_pluginConfiguration, "expirationDays");

                    int iExpirationDays = 0;

                    if (!string.IsNullOrEmpty(sExpirationDays))
                    {
                        iExpirationDays = Convert.ToInt32(sExpirationDays);
                    }
                    else
                    {
                        iExpirationDays = 10;
                    }


                    tracingService.Trace("Obtención de ejecutivo resolutor");
                    Guid IdOwner = Guid.Empty;

                    if (entity.Contains("New_ejecutivoresolutor"))
                    {
                        IdOwner = ((EntityReference)entity.Attributes["New_ejecutivoresolutor"]).Id;
                    }

                    tracingService.Trace("Asignación datos nueva entidad");
                    Entity newPeticion = new Entity("new_peticion");

                    newPeticion.Attributes["Regarding"] = new EntityReference("incident", entity.Id);
                    if (entity.Contains("title"))
                    {
                        newPeticion.Attributes["Subject"] = entity.Attributes["title"].ToString();
                    }
                    if (entity.Contains("Description"))
                    {
                        newPeticion.Attributes["New_descripcion"] = entity.Attributes["Description"].ToString();
                    }
                    if (entity.Contains("ticketnumber"))
                    {
                        newPeticion.Attributes["New_name"] = $"PET-{entity.Attributes["ticketnumber"].ToString()}";
                    }
                    newPeticion.Attributes["New_feharesolucion"] = DateTime.Now.AddDays(iExpirationDays);

                    var IdPeticion = service.Create(newPeticion);

                    tracingService.Trace("Asignación propietario");
                    if (IdOwner != Guid.Empty)
                    {
                        AsignarRegistro(service, "new_peticion", IdPeticion, "systemuser", IdOwner);
                    }
                    tracingService.Trace("Fin proceso");
                }
            }
            catch (Exception e)
            {
                tracingService.Trace("Error: " + e.Message);
                throw;
            }
        
        }

        public static string GetValueNode(XmlDocument doc, string key)
        {
            XmlNode node = doc.SelectSingleNode(String.Format("Settings/setting[@name='{0}']", key));
            if (node != null)
            {
                XmlNode oNode = node.SelectSingleNode("value");
                if (oNode != null)
                {
                    return oNode.InnerText;
                }
            }
            return string.Empty;
        }

        public static bool AsignarRegistro(IOrganizationService service, string NameEntityTarget, Guid IdTarget, string NameEntityAssign, Guid IdAssign)
        {
            bool res = new bool();
            try
            {
                AssignRequest assign = new AssignRequest
                {
                    Assignee = new EntityReference(NameEntityAssign,
                                IdAssign),
                    Target = new EntityReference(NameEntityTarget,
                                IdTarget)
                };
                service.Execute(assign);
                res = true;
            }
            catch (Exception e)
            {
                res = false;
            }
            return res;
        }
    }
}
