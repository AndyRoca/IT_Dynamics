using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

public interface ICrm
{
    Guid Create<T>(T entity) where T : Entity;
    List<T> RetriveMultipleByFetchXml<T>(string fetchXml) where T : Entity;
    List<T> RetrieveMultiple<T>(QueryBase query) where T : Entity;
    T Retrieve<T>(string entityName, Guid id, ColumnSet columnSet = null) where T : Entity;
    void Trace(string log);
    void Update(Entity entity);
    Guid CallingUser { get; }
    void Debug(string message);
    List<T> GetAll<T>(T entity, string[] fields = null, ConditionExpression conditionExpression = null) where T : Entity;
    void Delete(string entityName, Guid id);
    void Delete<T>(T entity) where T : Entity;
    List<T> GetAll<T>(T entity, string[] fields = null, ConditionExpression conditionExpression = null, List<OrderExpression> orders = null) where T : Entity;
    Y Execute<T, Y>(T request) where T : OrganizationRequest where Y : OrganizationResponse;

    string OrganizationName { get; }
}
