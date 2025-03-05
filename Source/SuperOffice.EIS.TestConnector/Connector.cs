using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using SuperOffice.CRM;
using SuperOffice.CRM.Globalization;

namespace SuperOffice.ErpSync.TestConnector
{
    [ErpConnector(ConnectorName)]
    public class Connector : IErpConnector
    {
        private class ConnectionNotFoundException : Exception
        {
            public Guid ConnectionId { get; set; }

            public ConnectionNotFoundException(Guid connectionId)
            {
                ConnectionId = connectionId;
            }
        }

        private static class ResponseHelper<ResponseType> where ResponseType : PluginResponseInfo, new()
        {
            public static ResponseType RequestConnectionInfo(Guid connectionId)
            {
                var ri = new ResponseType
                {
                    State = ResponseState.Error,
                    ErrorCode = ConnectorWS.ResponseErrorCodes.UNKNOWN_CONNECTION_ID
                };
                return ri;
            }
        }

        public const string ConnectorName = "Test.Excel";
        private readonly string _connectionsFile = "";

        private readonly List<Connection> _connectionList = new List<Connection>();

        public Connector()
        {
            var tempPath = Path.GetTempPath();
            _connectionsFile = Path.Combine(tempPath, "EIS_Connections.txt");
        }

        private Connection GetConnection(Guid connectionId)
        {
            var conn = (
                from c in _connectionList
                where c.ConnectionId == connectionId
                select c).FirstOrDefault();

            if (conn == null)
            {
                var filename = GetConnectionFilename(connectionId);

                if (!string.IsNullOrEmpty(filename))
                {
                    conn = new Connection(connectionId, filename);
                }
                else
                {
                    throw new ConnectionNotFoundException(connectionId);
                }
            }

            if (conn == null)
                throw new ConnectionNotFoundException(connectionId);

            return conn;
        }

        public FieldMetadataInfoArrayPluginResponse GetConfigData()
        {
            var ri = new FieldMetadataInfoArrayPluginResponse
            {
                IsOk = true
            };

            var fields = new List<FieldMetadataInfo>();
            // Mandatory field
            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Mandatory,
                DefaultValue = "C:\\Temp\\EIS\\Client.xlsm",
                DisplayDescription = "NO:\"Filnavn for excel-dokumentet (hele stien)\";US:\"Filename for excel document (full path)\"",
                DisplayName = "NO:\"Filnavn\";US:\"File Name\"",
                FieldKey = "Filename",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = CultureDataFormatter.EncodeInt(0),
                DisplayDescription = "NO:\"Avkryssningstestfelt\";US:\"Checkbox field test\"",
                DisplayName = "NO:\"Avkryssningsfelt\";US:\"Checkbox field test tooltip\"",
                FieldKey = "CheckboxField",
                FieldType = FieldMetadataTypeInfo.Checkbox,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = CultureDataFormatter.EncodeDateTime(DateTime.Now),
                DisplayDescription = "NO:\"Datotestfelt\";US:\"Datepicker test field\"",
                DisplayName = "NO:\"Datotestfelt\";US:\"Date picker test field\"",
                FieldKey = "DateField",
                FieldType = FieldMetadataTypeInfo.Datetime,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = CultureDataFormatter.EncodeDouble(0.0),
                DisplayDescription = "Doubletestfelt",
                DisplayName = "Doubletestfelt",
                FieldKey = "DoubleField",
                FieldType = FieldMetadataTypeInfo.Double,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = CultureDataFormatter.EncodeInt(0),
                DisplayDescription = "Integertestfelt",
                DisplayName = "Integertestfelt",
                FieldKey = "IntegerField",
                FieldType = FieldMetadataTypeInfo.Integer,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "[password]",
                DisplayDescription = "Passordtestfelt",
                DisplayName = "Passordtestfelt",
                FieldKey = "PasswordField",
                FieldType = FieldMetadataTypeInfo.Password,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "[text]",
                DisplayDescription = "Teksttestfelt",
                DisplayName = "Teksttestfelt",
                FieldKey = "TextField",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Listetestfelt",
                DisplayName = "Listetestfelt",
                FieldKey = "ListField",
                FieldType = FieldMetadataTypeInfo.List,
                ListName = "ConnectorTestList",
                MaxLength = 500
            });

            ri.FieldMetaDataObjects = fields.ToArray();

            return ri;
        }

        public PluginResponseInfo TestConfigData(Dictionary<string, string> connectionInfo)
        {
            var ri = new FieldMetadataInfoArrayPluginResponse
            {
                IsOk = true
            };

            if (!connectionInfo.ContainsKey("Filename"))
            {
                ri.UserExplanation = "Filename field not found";
                ri.IsOk = false;
            }
            else if (string.IsNullOrEmpty(connectionInfo["Filename"]))
            {
                ri.UserExplanation = "Filename field can not be empty";
                ri.IsOk = false;
                return ri;
            }
            else if (!File.Exists(connectionInfo["Filename"]))
            {
                ri.UserExplanation = string.Format("Excel Filename '{0}' does not exist on the ERP Services server.", connectionInfo["Filename"]);
                ri.IsOk = false;
                return ri;
            }

            return ri;
        }

        private List<string> GetConnectionsFile()
        {
            List<string> lines;
            if (File.Exists(_connectionsFile))
                lines = File.ReadAllLines(_connectionsFile).ToList();
            else
                lines = new List<string>();

            return lines;
        }

        private string GetConnectionFilename(Guid connectionId)
        {
            var lines = GetConnectionsFile();

            foreach (var connInfo in lines.Select(r => r.Split(';')).Where(r => r.Count() == 2))
            {
                if (new Guid(connInfo[0]) == connectionId)
                    return connInfo[1];
            }

            return "";
        }

        private bool SaveConnectionFilename(Guid connectionId, string filename)
        {
            var lines = GetConnectionsFile();
            //var found = false;

            for (var i = lines.Count() - 1; i >= 0; i--)
            {
                var lineArr = lines[i].Split(';');
                if (lineArr.Count() != 2)
                    continue;

                if (connectionId == new Guid(lineArr[0]))
                {
                    lines.RemoveAt(i);
                    break;
                }
            }

            lines.Add(connectionId.ToString() + ";" + filename);

            File.WriteAllLines(_connectionsFile, lines.ToArray());

            // Is the connection in the active connection list? If so, remove it so that it starts up with the new file next time.
            var conn = (
                from c in _connectionList
                where c.ConnectionId == connectionId
                select c).FirstOrDefault();

            if (conn != null)
                _connectionList.Remove(conn);

            return true;
        }

        public PluginResponseInfo SaveConnection(Guid connectionID, Dictionary<string, string> connectionInfo)
        {
            var ri = new FieldMetadataInfoArrayPluginResponse();

            if (!connectionInfo.ContainsKey("Filename"))
            {
                ri.UserExplanation = "Filename field not found";
                ri.IsOk = false;
                return ri;
            }

            if (string.IsNullOrEmpty(connectionInfo["Filename"]))
            {
                ri.UserExplanation = "Filename field can not be empty";
                ri.IsOk = false;
                return ri;
            }

            var excelFilename = connectionInfo["Filename"];

            ri.IsOk = true;

            try
            {
                SaveConnectionFilename(connectionID, excelFilename);

                // Does the Excel file exist? If not, let's make an attempt to copy a "clean" one in there
                try
                {
                    var dir = new FileInfo(excelFilename).Directory.FullName;

                    if (!Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }
                    if (!File.Exists(excelFilename))
                    {
                        var assemblyLocation = new FileInfo(@Assembly.GetExecutingAssembly().Location).Directory.FullName;
                        var copySourceFile = Path.Combine(assemblyLocation, "ErpClient.xlsm");
                        if (File.Exists(copySourceFile))
                            File.Copy(copySourceFile, excelFilename);
                    }
                }
                catch (Exception)
                {
                    ri.State = ResponseState.OkWithInfo;
                    ri.UserExplanation = "Operation succeeded, but could not copy default Excel connection file to destination; needs to be copied manually.";
                }
            }
            catch (Exception ex)
            {
                ri.UserExplanation = "Error saving connection info! " + ex.Message;
                ri.TechExplanation = ex.ToString();
                ri.IsOk = false;
                return ri;
            }

            return ri;
        }

        public PluginResponseInfo TestConnection(Guid connectionID)
        {
            try
            {
                return GetConnection(connectionID).TestConnection();
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<PluginResponseInfo>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public PluginResponseInfo DeleteConnection(Guid connectionID)
        {
            var lines = GetConnectionsFile();

            for (var i = 0; i < lines.Count; i++)
            {
                var fields = lines[i].Split(';');

                if (fields.Count() != 2)
                    continue;

                if (fields[0] == connectionID.ToString())
                {
                    lines.RemoveAt(i);
                    break;
                }
            }

            File.WriteAllLines(_connectionsFile, lines.ToArray());

            return new PluginResponseInfo();
        }

        public StringArrayPluginResponse GetSupportedActorTypes(Guid connectionID)
        {
            try
            {
                return GetConnection(connectionID).GetSupportedActorTypes();
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<StringArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public FieldMetadataInfoArrayPluginResponse GetSupportedActorTypeFields(Guid connectionID, string actorType)
        {
            try
            {
                return GetConnection(connectionID).GetSupportedActorTypeFields(actorType);
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<FieldMetadataInfoArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public ActorArrayPluginResponse GetActors(Guid connectionID, string actorType, string[] erpKeys, string[] fieldKeys)
        {
            try
            {
                return GetConnection(connectionID).GetActors(actorType, erpKeys, fieldKeys);
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<ActorArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public StringArrayPluginResponse GetSearchableFields(Guid connectionId, string actorType)
        {
            try
            {
                var response = new StringArrayPluginResponse();
                response.Items = GetSupportedActorTypeFields(connectionId, actorType).FieldMetaDataObjects.Select(r => r.FieldKey).ToArray();
                response.IsOk = true;
                return response;
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<StringArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public ActorArrayPluginResponse SearchActors(Guid connectionID, string actorType, string searchText, string[] fieldKeys)
        {
            try
            {
                return GetConnection(connectionID).SearchActors(actorType, searchText, fieldKeys);
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<ActorArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public ActorArrayPluginResponse SearchActorsAdvanced(Guid connectionID, string actorType, SearchRestrictionInfo[] restrictions, string[] returnFields)
        {
            try
            {
                return GetConnection(connectionID).SearchActorsAdvanced(actorType, restrictions, returnFields);
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<ActorArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public ActorArrayPluginResponse SearchActorByParent(Guid connectionID, string actorType, string searchText, string parentActorType, string parentActorErpKey, string[] fieldKeys)
        {
            try
            {
                return GetConnection(connectionID).SearchActorByParent(actorType, searchText, parentActorType, parentActorErpKey, fieldKeys);
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<ActorArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public ActorPluginResponse CreateActor(Guid connectionID, ErpActor act)
        {
            try
            {
                return GetConnection(connectionID).CreateActor(act);
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<ActorPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public ActorArrayPluginResponse SaveActors(Guid connectionID, ErpActor[] actors)
        {
            try
            {
                return GetConnection(connectionID).SaveActors(actors);
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<ActorArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public ListItemArrayPluginResponse GetList(Guid connectionID, string listName)
        {
            if (connectionID == Guid.Empty)
            {
                var ri = new ListItemArrayPluginResponse();

                if (listName.ToLower() == "ConnectorTestList".ToLower())
                {
                    ri.ListItems = new Dictionary<string, string>
                    {
                        { "item1", "Listeverdi 1" },
                        { "item2", "Listeverdi 2" },
                        { "item3", "Listeverdi 3" },
                        { "item4", "Listeverdi 4" },
                        { "item5", "Listeverdi 5" }
                    };
                }
                return ri;
            }
            else
            {
                try
                {
                    return GetConnection(connectionID).GetList(listName);
                }
                catch (ConnectionNotFoundException ex)
                {
                    return ResponseHelper<ListItemArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
                }
            }
        }

        public ListItemArrayPluginResponse GetListItems(Guid connectionID, string listName, string[] listItemKeys)
        {
            try
            {
                if (connectionID == Guid.Empty)
                {
                    if (listName.ToLower() == "ConnectorTestList".ToLower())
                    {
                        var list = GetList(connectionID, "ConnectorTestList");

                        var items = (
                            from i in list.ListItems
                            where listItemKeys.Contains(i.Key)
                            select i).ToDictionary(r => r.Key, r => r.Value);

                        if (items.Any())
                            return new ListItemArrayPluginResponse(items);
                    }
                    return new ListItemArrayPluginResponse();
                }
                else
                {
                    return GetConnection(connectionID).GetListItems(listName, listItemKeys);
                }
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<ListItemArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }

        public ActorArrayPluginResponse GetActorsByTimestamp(Guid connectionID, string updatedOnOrAfter, string actorType, string[] fieldKeys)
        {
            try
            {
                return GetConnection(connectionID).GetActorsByTimestamp(updatedOnOrAfter, actorType, fieldKeys);
            }
            catch (ConnectionNotFoundException ex)
            {
                return ResponseHelper<ActorArrayPluginResponse>.RequestConnectionInfo(ex.ConnectionId);
            }
        }
    }
}
