using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Xrm.Tooling.Connector;
using System.Xml.Linq;
using Microsoft.Xrm.Sdk.Messages;

public class Crm : IDisposable
{
    public IOrganizationService Service { get; set; }

    private readonly CrmServiceClient _service;


    public Crm(string connectionString)
    {

        _service = new CrmServiceClient(connectionString);

        if (_service != null & _service.IsReady)
            Service = _service;
    }

    public void Trace(string log)
    {



    }

    public Guid CallingUser
    {
        get
        {
            var response = (WhoAmIResponse)_service.Execute(new WhoAmIRequest());
            return response.UserId;

        }
    }

    public string OrganizationName => _service.ConnectedOrgFriendlyName;

    public List<T> GetAll<T>(T entity, string[] fields = null, ConditionExpression conditionExpression = null) where T : Entity
    {
        QueryExpression exp = new QueryExpression(entity.LogicalName);
        if (fields != null)
            exp.ColumnSet = new ColumnSet(fields);
        else exp.ColumnSet = new ColumnSet(true);

        exp.Criteria.AddCondition(conditionExpression);
        int page = 1;

        string pagingCookie = null;
        exp.PageInfo = new PagingInfo()
        {
            PageNumber = 1,
            PagingCookie = null,
            Count = 5000
        };

        List<Entity> results = new List<Entity>();

        while (true)
        {
            exp.PageInfo = new PagingInfo()
            {
                PageNumber = page,
                PagingCookie = pagingCookie,
                Count = 5000
            };

            EntityCollection col = _service.RetrieveMultiple(exp);
            results.AddRange(col.Entities);

            if (!col.MoreRecords)
            {
                return results.Cast<T>().ToList();
            }

            page++;
            pagingCookie = col.PagingCookie;
        }
    }
    public List<T> GetAll<T>(string fetchXml) where T : Entity
    {

        FetchXmlToQueryExpressionRequest request = new FetchXmlToQueryExpressionRequest()
        {
            FetchXml = fetchXml
        };

        FetchXmlToQueryExpressionResponse response = (FetchXmlToQueryExpressionResponse)_service.Execute(request);

        return this.GetAll<T>(response.Query).Cast<T>().ToList();
    }
    public List<T> GetAll<T>(T entity, string[] fields = null, ConditionExpression conditionExpression = null, List<OrderExpression> orders = null) where T : Entity
    {
        QueryExpression exp = new QueryExpression(entity.LogicalName);
        if (fields != null)
            exp.ColumnSet = new ColumnSet(fields);
        else exp.ColumnSet = new ColumnSet(true);
        if (orders != null)
            orders.ForEach(_order => exp.Orders.Add(_order));

        exp.Criteria.AddCondition(conditionExpression);
        int page = 1;

        string pagingCookie = null;
        exp.PageInfo = new PagingInfo()
        {
            PageNumber = 1,
            PagingCookie = null,
            Count = 5000
        };

        List<Entity> results = new List<Entity>();

        while (true)
        {
            exp.PageInfo = new PagingInfo()
            {
                PageNumber = page,
                PagingCookie = pagingCookie,
                Count = 5000
            };

            EntityCollection col = Service.RetrieveMultiple(exp);
            results.AddRange(col.Entities);

            if (!col.MoreRecords)
            {
                return results.Cast<T>().ToList();
            }

            page++;
            pagingCookie = col.PagingCookie;
        }
    }
    public List<T> GetAll<T>(QueryExpression exp) where T : Entity
    {
        int page = 1;
        string pagingCookie = null;
        exp.PageInfo = new PagingInfo()
        {
            PageNumber = 1,
            PagingCookie = null,
            Count = 5000
        };

        List<T> results = new List<T>();

        while (true)
        {
            exp.PageInfo = new PagingInfo()
            {
                PageNumber = page,
                PagingCookie = pagingCookie,
                Count = 5000
            };

            EntityCollection col = _service.RetrieveMultiple(exp);
            List<T> entities = col.Entities.Cast<T>().ToList();
            results.AddRange(entities);

            if (!col.MoreRecords)
            {
                return results;
            }

            page++;
            pagingCookie = col.PagingCookie;
        }
    }
    public void AddReference(string reference)
    {
        var projectPath = @"C:\Users\xavik0f2\Downloads\GIT Windows\CrmEntityClassGenerator\EarlyBoundTypescript\EarlyBoundTypescript.njsproj";
        XDocument projDefinition = XDocument.Load(projectPath);
        XNamespace msbuild = "http://schemas.microsoft.com/developer/msbuild/2003";
        var TypescriptReferences = projDefinition.Element(msbuild + "Project").Elements(msbuild + "ItemGroup").Elements(msbuild + "TypeScriptCompile")
            .Select(refElem => refElem.Attribute("Include").Value);
        var project = projDefinition.ToString();
        var x = new System.Text.RegularExpressions.Regex(@"(?<=\</TypeScriptCompile>)(?!.*/TypeScriptCompile>)(.*?)(?=\</ItemGroup>)", System.Text.RegularExpressions.RegexOptions.Singleline);
        string repl = $"\n\t\t<TypeScriptCompile Include=\"EarlyBoundTypeScript\\{reference}\"></TypeScriptCompile>\n";
        string result = x.Replace(project, repl);
        File.WriteAllText(projectPath, result);
    }

    public Guid Create<T>(T entity) where T : Entity
    {
        return Service.Create(entity);
    }

    public List<Guid> ExecuteMultipleCreate<Y, T>(List<Y> registers, Func<Y, T> MapeaEntity, int registerPerPage = 200) where T : Entity
    {
        var collection = new OrganizationRequestCollection();
        List<Guid> guidRegisters = new List<Guid>();


        registers.ForEach(ent =>
        {
            var entity = MapeaEntity(ent);
            collection.Add(new CreateRequest { Target = entity });

            if (collection.Count == registerPerPage)
            {
                guidRegisters.AddRange(ExecuteMultipleCreate(collection));
                collection.Clear();
            }
        });

        if (collection.Count > 0)
            guidRegisters.AddRange(ExecuteMultipleCreate(collection));

        return guidRegisters;

    }

    private List<Guid> ExecuteMultipleCreate(OrganizationRequestCollection collection)
    {
        List<Guid> guidRegisters = new List<Guid>();

        ExecuteMultipleRequest requests = new ExecuteMultipleRequest
        {
            Requests = collection,
            Settings = new ExecuteMultipleSettings
            {
                ContinueOnError = true,
                ReturnResponses = true,
            }
        };
        ExecuteMultipleResponse response = (ExecuteMultipleResponse)_service.Execute(requests);
        if (response.IsFaulted == false)
        {
            for (int i = 0; i < response.Responses.Count; i++)
            {
                var createResponse = (CreateResponse)response.Responses[i].Response;
                guidRegisters.Add(createResponse.id);
            }

        }
        else
            throw new Exception("Error en ExecuteMultipleCreateCrm.Invoke()");

        return guidRegisters;
    }

    public ExecuteMultipleResponse ExecuteMultiple(OrganizationRequestCollection collectionOfEmailsToBeSent)
    {
        ExecuteMultipleRequest sendEmailRequest = new ExecuteMultipleRequest
        {
            Requests = collectionOfEmailsToBeSent,
            Settings = new ExecuteMultipleSettings
            {
                ContinueOnError = true,
                ReturnResponses = true,
            }
        };

        return (ExecuteMultipleResponse)_service.Execute(sendEmailRequest);
    }

    public T Retrieve<T>(T entity, ColumnSet columnSet = null) where T : Entity
    {
        if (columnSet == null) columnSet = new ColumnSet(true);

        return (T)Service.Retrieve(entity.LogicalName, entity.Id, columnSet);
    }

    public void Update(Entity entity)
    {
        Service.Update(entity);
    }

    public void Delete(string entityName, Guid id)
    {
        Service.Delete(entityName, id);
    }

    public void Delete<T>(T entity) where T : Entity
    {
        Service.Delete(entity.LogicalName, entity.Id);
    }

    public OrganizationResponse Execute(OrganizationRequest request)
    {
        return Service.Execute(request);
    }

    public void Associate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
    {
        Service.Associate(entityName, entityId, relationship, relatedEntities);
    }

    public void Disassociate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
    {
        Service.Disassociate(entityName, entityId, relationship, relatedEntities);
    }

    public List<T> RetrieveMultiple<T>(QueryBase query) where T : Entity
    {
        return Service.RetrieveMultiple(query).Entities.Cast<T>().ToList();
    }

    public List<T> RetriveMultipleByFetchXml<T>(string fetchXml) where T : Entity
    {

        return Service.RetrieveMultiple(new FetchExpression(fetchXml)).Entities.Cast<T>().ToList();
    }

    public T Retrieve<T>(string entityName, Guid id, ColumnSet columnSet = null) where T : Entity
    {
        if (columnSet == null)
            columnSet = new ColumnSet(true);
        return (T)Service.Retrieve(entityName, id, columnSet);

    }

    public void Debug(string message = "DEBUGGING")
    {
        throw new InvalidPluginExecutionException(message);
    }

    public void Dispose()
    {
        _service.Dispose();
    }

    public Y Execute<T, Y>(T request)
        where T : OrganizationRequest
        where Y : OrganizationResponse
    {
        return (Y)Service.Execute(request);
    }
}