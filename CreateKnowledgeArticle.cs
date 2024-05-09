using Microsoft.Xrm.Sdk;
using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace SP.Plugins.Articulos_de_conocimiento
{
    public class CreateKnowledgeArticle : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            PluginHelper crm = new PluginHelper(serviceProvider);
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            try
            {
                if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity))
                    return;

                var knowledgeArticle = (Entity)context.InputParameters["Target"];

                string excelFilePath = "C:\\Proyectos\\AC\\KnowledgeBase_PropuestaTagueado.xlsx";
                string folderPath = "C:\\Proyectos\\AC\\TxtContenidos";
                string sheetName = "Artículos Carpetas KB";

                FileInfo fileInfo = new FileInfo(excelFilePath);

                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet(sheetName);

                    var rows = worksheet.RowsUsed();
                    foreach (var row in rows.Skip(1))
                    {
                        Guid id = Guid.NewGuid();
                        knowledgeArticle["knowledgearticleid"] = id;
                        knowledgeArticle["statecode"] = new OptionSetValue(0);
                        knowledgeArticle["statuscode"] = new OptionSetValue(1);

                        string title = row.Cell(1).Value.ToString();
                        knowledgeArticle["title"] = title;
                        knowledgeArticle["keywords"] = row.Cell(2).Value.ToString();

                        string txtFilePath = Path.Combine(folderPath, $"{title}.txt");
                        string rawContent = File.ReadAllText(txtFilePath);
                        string content = rawContent.Replace("\u001F", "").Replace("\x07", "").Replace("\x06", "").Replace("\x05", "").Replace("\x02", "").Replace("\x04", "").Replace("\x08", "").Replace("\x0E", "").Replace("\x0F", "").Replace("\x01", "").Replace("\x12", "").Replace("\x0C", "").Replace("\x0B", "").Replace("\x10", "").Replace("\x03", "").Replace("\x13", "").Replace("\x14", "").Replace("\u0300", "").Replace("\u0015", "").Replace("\u0011", "").Replace("\u001B", "").Replace("\u0016", "").Replace("\u001D", "").Replace("\u001E", "").Replace("\u001C", "").Replace("\u001A", "").Replace("\u0019", "").Replace("\u0018", "").Replace("\u0017", "");

                        knowledgeArticle["content"] = content;

                        crm.Create(knowledgeArticle);

                        knowledgeArticle["statecode"] = new OptionSetValue(3);
                        knowledgeArticle["statuscode"] = new OptionSetValue(7);

                        crm.Update(knowledgeArticle);
                        Console.WriteLine(title);
                    }
                }

            }
            catch (Exception ex)
            {
                crm.Trace("Exception" + ex.Message + " " + ex.StackTrace);
                throw new InvalidPluginExecutionException("Error en CreateKnowledgeArticle" + ex.Message);
            }
        }

        public class KnowledgeArticleObject
        {
            public string Title { get; set; }
            public string Content { get; set; }
            public string Keyword { get; set; }
        }

    }
}

