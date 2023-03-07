using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace PaginatedDatasetPlugin
{
    public class PaginatedDataset : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            tracingService.Trace("plugin execution started");

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("jeschro_entityName"))
            {
                // Get input parameters.  
                Entity entity = new Entity(context.InputParameters["jeschro_entityName"].ToString());
                int pageNumber = (int)context.InputParameters["jeschro_PageNumber"];
                int pageSize = (int)context.InputParameters["jeschro_PageSize"];
                string sortField = context.InputParameters["jeschro_Sort"].ToString();

                string pagingCookie = null;

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    tracingService.Trace("executing query");
                    
                    //Create Query Expression
                    QueryExpression query = new QueryExpression(entity.LogicalName);
                    
                    //Return all columns
                    query.ColumnSet.AllColumns = true;
                    
                    //Set sort order
                    query.AddOrder(sortField, OrderType.Ascending);

                    //Set Return Total Record Count
                    query.PageInfo.ReturnTotalRecordCount = true;
                    
                    //Set paging cookie
                    if(!String.IsNullOrEmpty(pagingCookie))
                    {
                        query.PageInfo.PagingCookie = pagingCookie;
                    }

                    //Set page size
                    query.PageInfo.Count = pageSize;
                    //Set page number
                    query.PageInfo.PageNumber = pageNumber;

                    //Retrieve records
                    EntityCollection results = service.RetrieveMultiple(query);

                    tracingService.Trace($"Entity: {entity.LogicalName}");
                    tracingService.Trace($"Total record count: {results.TotalRecordCount}");
                    tracingService.Trace($"Paging cookie: {results.PagingCookie}");
                    tracingService.Trace($"More records: {results.MoreRecords}");

                    //Set ouput parameters
                    context.OutputParameters.AddOrUpdateIfNotNull("jeschro_PagingCookie", results.PagingCookie);
                    context.OutputParameters.AddOrUpdateIfNotNull("jeschro_MoreRecords", results.MoreRecords);
                    context.OutputParameters.AddOrUpdateIfNotNull("jeschro_EntityCollection", results);
                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in PaginatedDatasetPlugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("PaginatedDatasetPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }

    }
}
