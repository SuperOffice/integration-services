namespace ConnectorService.Utils
{
    public interface IExcelHandler
    {
        string GetAllCapabilties(string fileName);
        string GetCapabilityRowByIndex(string fileName, int index);

        string InsertCapabilities(string fileName);

        void WriteToExcelSheet(string fileName);

        string ReadSheetCell(string fileName, string sheetName, int row, int column);

        
    }
}
