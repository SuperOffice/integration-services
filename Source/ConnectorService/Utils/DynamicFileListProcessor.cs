using NSwag.Generation.Processors;
using NSwag.Generation.Processors.Contexts;

public class DynamicFileListProcessor : IOperationProcessor
{
    private readonly string _directoryPath;

    public DynamicFileListProcessor(string directoryPath)
    {
        _directoryPath = directoryPath;
    }

    public bool Process(OperationProcessorContext context)
    {
        if ((context.OperationDescription.Method == "get") && (context.OperationDescription.Path.StartsWith("/file/{fileName}")))
        {
            var parameter = context.OperationDescription.Operation.Parameters.FirstOrDefault(p => p.Name == "fileName");
            if (parameter != null && parameter.Schema != null)
            {
                var availableFiles = Directory.GetFiles(_directoryPath)
                                              .Select(Path.GetFileName)
                                              .ToList<object>();

                if (availableFiles.Any())
                {
                    foreach (var file in availableFiles)
                    {
                        parameter.Schema.Enumeration.Add(file);
                    }
                    
                }
            }
        }
        return true;
    }
}
