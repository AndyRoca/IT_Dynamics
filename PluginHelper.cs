using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;


public class PluginHelper : IOrganizationService, ITracingService, ICrm
{
    IOrganizationService service { get; set; }
    ITracingService tracingService { get; set; }
    public IPluginExecutionContext context { get; set; }
    public PluginHelper(IServiceProvider serviceProvider)
    {

        this.tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        this.context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        this.service = serviceFactory.CreateOrganizationService(context.UserId);

    }
    public Guid CallingUser { get { return this.context.UserId; } }

    public string OrganizationName => context.OrganizationName;

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

            EntityCollection col = service.RetrieveMultiple(exp);
            results.AddRange(col.Entities);
            this.Trace($"{col.Entities.Count}");
            if (!col.MoreRecords)
            {
                return results.Cast<T>().ToList();
            }

            page++;
            pagingCookie = col.PagingCookie;
        }
    }
    public void Trace(string log)
    {
        this.tracingService.Trace(log);
    }
    public void Trace(string format, params object[] args)
    {
        this.tracingService.Trace(format);
    }

    public Guid Create<T>(T entity) where T : Entity
    {
        return this.service.Create(entity);
    }

    public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet)
    {
        return this.service.Retrieve(entityName, id, columnSet);
    }

    public void Update(Entity entity)
    {
        this.service.Update(entity);
    }

    public void Delete(string entityName, Guid id)
    {
        this.service.Delete(entityName, id);
    }

    public void Delete<T>(T entity) where T : Entity
    {
        this.service.Delete(entity.LogicalName, entity.Id);
    }

    public OrganizationResponse Execute(OrganizationRequest request)
    {
        return this.service.Execute(request);
    }

    public void Associate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
    {
        this.service.Associate(entityName, entityId, relationship, relatedEntities);
    }

    public void Disassociate(string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
    {
        this.service.Disassociate(entityName, entityId, relationship, relatedEntities);
    }
    public EntityCollection RetrieveMultiple(QueryBase query)
    {
        return this.service.RetrieveMultiple(query);
    }
    public List<T> RetrieveMultiple<T>(QueryBase query) where T : Entity
    {
        return this.service.RetrieveMultiple(query).Entities.Cast<T>().ToList();
    }
    public List<T> RetriveMultipleByFetchXml<T>(string fetchXml) where T : Entity
    {

        return this.service.RetrieveMultiple(new FetchExpression(fetchXml)).Entities.Cast<T>().ToList();
    }

    public T Retrieve<T>(string entityName, Guid id, ColumnSet columnSet = null) where T : Entity
    {
        if (columnSet == null)
            columnSet = new ColumnSet(true);
        return (T)this.Retrieve(entityName, id, columnSet);

    }

    public void Debug(string message = "DEBUGGING")
    {
        throw new InvalidPluginExecutionException(message);
    }

    public void PrintObject<T>(T _object)
    {
        this.Trace(this.GetStringObject(_object));
    }

    public string GetStringObject<T>(T _object)
    {
        StringBuilder str = new StringBuilder($"----- {_object.GetType()} -----" + Environment.NewLine);

        foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(_object))
        {
            if (descriptor.Name != null)
            {
                string name = descriptor.Name;
                object value = new object();
                if (_object.GetType() == (new Money()).GetType())
                {
                    value = ((Money)descriptor.GetValue(_object)).Value;
                }
                else if (_object.GetType() == (new OptionSetValue()).GetType())
                {
                    value = ((OptionSetValue)descriptor.GetValue(_object)).Value;
                }
                else if (_object.GetType() == (new Microsoft.Xrm.Sdk.EntityReference()).GetType())
                {
                    value = ((EntityReference)descriptor.GetValue(_object)).Id;
                }
                else
                {
                    value = descriptor.GetValue(_object);

                }
                str.Append(name + ": " + value + Environment.NewLine);
            }

        }

        str.Append(Environment.NewLine);
        return (str.ToString());
    }

    public Guid Create(Entity entity)
    {
        return this.service.Create(entity);
    }

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

            EntityCollection col = service.RetrieveMultiple(exp);
            results.AddRange(col.Entities);

            if (!col.MoreRecords)
            {
                return results.Cast<T>().ToList();
            }

            page++;
            pagingCookie = col.PagingCookie;
        }
    }

    public Y Execute<T, Y>(T request)
    where T : OrganizationRequest
    where Y : OrganizationResponse
    {
        return (Y)this.service.Execute(request);
    }
}