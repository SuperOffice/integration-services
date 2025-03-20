using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SuperOffice.CRM;
using SuperOffice.CRM.Globalization;
using SuperOffice.ErpSync;
using SuperOffice.ErpSync.ConnectorWS;

namespace ErpConnector
{
    internal class Connection
    {
        private ExcelHandler _Excel = null;
        private readonly string ExcelDoc = "";
        private readonly string _fileBaseDir = "";

        public Connection(Guid connectionId, string excelDoc)
        {
            ConnectionId = connectionId;
            _fileBaseDir = Path.Combine(AppContext.BaseDirectory, "Resources");
            ExcelDoc = Path.Combine(_fileBaseDir, excelDoc);

            //ExcelDoc = excelDoc;
        }

        public bool Gui { get; set; }
        public Guid ConnectionId { get; set; }

        private ExcelHandler Excel
        {
            get
            {
                if (_Excel == null)
                {
                    //if (Gui)
                    //	_Excel = new ExcelHandler_Gui();
                    //else
                    _Excel = new ExcelHandler();
                }

                if (!string.IsNullOrEmpty(ExcelDoc))
                    _Excel.OpenExcelDoc(ExcelDoc);

                return _Excel;
            }
        }

        private bool TestConnectionStatus(bool attemptToOpen = true)
        {
            if (Excel.IsExcelOpen())
            {
                return true;
            }
            else
            {
                if (attemptToOpen && ExcelDoc != "")
                {
                    Excel.OpenExcelDoc(ExcelDoc);
                    return TestConnectionStatus(false);
                }
            }

            return false;
        }

        public PluginResponseInfo TestConnection()
        {
            var ri = new FieldMetadataInfoArrayPluginResponse();

            if (TestConnectionStatus())
            {
                ri.IsOk = true;
            }
            else
            {
                ri.IsOk = false;
                ri.State = ResponseState.Error;
                ri.UserExplanation = "Connection test failed";
                ri.TechExplanation = "Could not find open Excel application and/or open TestConnector Excel document";
            }

            return ri;
        }

        public StringArrayPluginResponse GetSupportedActorTypes()
        {
            var ri = new StringArrayPluginResponse();
            ri.IsOk = true;
            var types = new List<string>();
            types.Add("Customer");
            types.Add("Supplier");
            types.Add("Person");
            types.Add("Project");
            ri.Items = types.ToArray();

            return ri;
        }

        public FieldMetadataInfoArrayPluginResponse GetSupportedActorTypeFields(string actorType)
        {
            var ri = new FieldMetadataInfoArrayPluginResponse();
            ri.IsOk = true;

            switch (actorType)
            {
                case "Customer":
                    ri.FieldMetaDataObjects = GetCustomerFields();
                    break;

                case "Supplier":
                    ri.FieldMetaDataObjects = GetSupplierFields();
                    break;

                case "Person":
                    ri.FieldMetaDataObjects = GetPersonFields();
                    break;

                case "Project":
                    ri.FieldMetaDataObjects = GetProjectFields();
                    break;
            }

            return ri;
        }

        public ActorArrayPluginResponse GetActors(string actorType, string[] erpKeys, string[] fieldKeys)
        {
            var actors = new List<ErpActor>();
            var ri = new ActorArrayPluginResponse();

            if (TestConnectionStatus())
            {
                foreach (var erpKey in erpKeys)
                {
                    var row = Excel.GetRowByID(actorType, erpKey);
                    var act = GenerateActorFromExcelRow(row, actorType, fieldKeys);

                    if (act != null)
                        actors.Add(act);
                }

                ri.IsOk = true;
            }
            else
            {
                ri.IsOk = true;
                ri.State = ResponseState.Warning;
                ri.UserExplanation = "Connection closed.";
                ri.TechExplanation = "Could not find open Excel application and/or open TestConnector Excel document";
            }

            ri.Actors = actors.ToArray();
            return ri;
        }

        public ActorArrayPluginResponse SearchActorsAdvanced(string actorType, SearchRestrictionInfo[] restrictions, string[] fieldKeys)
        {
            var actors = new List<ErpActor>();
            var ri = new ActorArrayPluginResponse();
            var allFields = GetSupportedActorTypeFields(actorType).FieldMetaDataObjects;

            // We need to add the restriction fields to evalute the conditions
            var allNeededFields = new List<string>(fieldKeys == null ? 0 : fieldKeys.Length);
            allNeededFields.AddRange(fieldKeys);
            foreach (var restriction in restrictions)
                if (!allNeededFields.Contains(restriction.FieldKey))
                    allNeededFields.Add(restriction.FieldKey);

            if (TestConnectionStatus())
            {
                var rows = Excel.GetAllRows(actorType);

                foreach (var row in rows)
                {
                    var actor = GenerateActorFromExcelRow(row, actorType, allNeededFields.ToArray()); // Convert to correctly handle number/int/double/date etc.

                    var searchHit = true;

                    // Search the row according to restrictions
                    foreach (var res in restrictions)
                    {
                        if (res.FieldKey == SpecialSearchKeys.PARENT_ERPKEY)
                            res.FieldKey = "ParentID";
                        else if (res.FieldKey == SpecialSearchKeys.PARENT_ACTORTYPE)
                            res.FieldKey = "ParentType";

                        var fld = (
                            from f in allFields
                            where f.FieldKey.ToLowerInvariant() == res.FieldKey.ToLowerInvariant()
                            select f).FirstOrDefault();

                        var erpResKey = TranslateFieldFromCrm(actorType, res.FieldKey);
                        if (fld == null || !row.ContainsKey(erpResKey)) // Unsupported or missing field. Continue
                            continue;

                        if (fld.FieldType == FieldMetadataTypeInfo.List)
                        {
                            try
                            {
                                var searchValues = res.Values.Select(r => CultureDataFormatter.ParseEncodedInt(r)).ToArray();
                                var actVal = 0;

                                if (row[erpResKey] != null && !string.IsNullOrWhiteSpace(row[erpResKey].ToString()))
                                    actVal = Convert.ToInt32(row[erpResKey]);

                                if (!searchValues.Contains(actVal) && res.Operator == ListOperators.ONE_OF)
                                    searchHit = false;
                                else if (searchValues.Contains(actVal) && res.Operator == ListOperators.NOT_ONE_OF)
                                    searchHit = false;
                            }
                            catch (Exception)
                            {
                                searchHit = false;
                            }

                            if (!searchHit)
                                break;
                        }
                        else
                        {
                            // Use the convenient IsMatch method, now available from SuperOffice. Call 555-123-0293 today!
                            //if (!SearchHelper.IsMatch(row[erpResKey], fld.FieldType, res))
                            if (!SearchHelper.IsMatch(actor.FieldValues[TranslateFieldToCrm(actorType, erpResKey)], fld.FieldType, res))
                            {
                                searchHit = false;
                                break;
                            }
                        }
                    }

                    // If we get all the way through the loop without searchHit being set to false, we've got a match!
                    if (searchHit)
                        actors.Add(actor);
                    //actors.Add(GenerateActorFromExcelRow(row, actorType, fieldKeys));
                }

                ri.IsOk = true;
            }
            else
            {
                ri.IsOk = true;
                ri.State = ResponseState.Warning;
                ri.UserExplanation = "Connection closed.";
                ri.TechExplanation = "Could not find open Excel application and/or open TestConnector Excel document";
            }

            ri.Actors = actors.ToArray();
            return ri;
        }

        public ActorArrayPluginResponse SearchActors(string actorType, string searchText, string[] fieldKeys)
        {
            var actors = new List<ErpActor>();
            var ri = new ActorArrayPluginResponse();

            if (TestConnectionStatus())
            {
                var rows = Excel.GetRowBySearchString(actorType, searchText);

                foreach (var row in rows)
                {
                    actors.Add(GenerateActorFromExcelRow(row, actorType, fieldKeys));
                }

                ri.IsOk = true;
            }
            else
            {
                ri.IsOk = true;
                ri.State = ResponseState.Warning;
                ri.UserExplanation = "Connection closed.";
                ri.TechExplanation = "Could not find open Excel application and/or open TestConnector Excel document";
            }

            ri.Actors = actors.ToArray();
            return ri;
        }

        public ActorArrayPluginResponse SearchActorByParent(string actorType, string searchText, string parentActorType, string parentActorErpKey, string[] fieldKeys)
        {
            // TODO: Horribly inefficient, just thrown together as an extremely quick version
            var tmpActors = new List<ErpActor>();
            var actors = new List<ErpActor>();
            var ri = new ActorArrayPluginResponse();

            if (TestConnectionStatus())
            {
                var rows = Excel.GetRowBySearchString(actorType, searchText);

                foreach (var row in rows)
                {
                    tmpActors.Add(GenerateActorFromExcelRow(row, actorType, fieldKeys));
                }

                foreach (var act in tmpActors)
                {
                    if (act.ParentActorType == parentActorType && act.ParentErpKey == parentActorErpKey)
                        actors.Add(act);
                }

                ri.IsOk = true;
            }
            else
            {
                ri.IsOk = true;
                ri.State = ResponseState.Warning;
                ri.UserExplanation = "Connection closed.";
                ri.TechExplanation = "Could not find open Excel application and/or open TestConnector Excel document";
            }

            ri.Actors = actors.ToArray();
            return ri;
        }

        public ActorPluginResponse CreateActor(ErpActor act)
        {
            var ri = new ActorPluginResponse();

            if (TestConnectionStatus())
            {
                var decodedFieldValues = new Dictionary<string, object>();

                foreach (var fld in act.FieldValues)
                {
                    decodedFieldValues.Add(TranslateFieldFromCrm(act.ActorType, fld.Key), CultureDataFormatter.ParseEncoded(fld.Value));
                }

                if (act.ActorType == "Person")
                {
                    decodedFieldValues.Add("ParentID", act.ParentErpKey);
                    decodedFieldValues.Add("ParentType", act.ParentActorType);
                }

                if (act.ActorType == "Person")
                    decodedFieldValues["PersonNo"] = Excel.NextId(act.ActorType);
                else if (act.ActorType == "Customer")
                    decodedFieldValues["CustNo"] = Excel.NextId(act.ActorType);
                else if (act.ActorType == "Supplier")
                    decodedFieldValues["SupNo"] = Excel.NextId(act.ActorType);
                else if (act.ActorType == "Project")
                    decodedFieldValues["ProjNo"] = "PROJ" + Excel.NextId(act.ActorType);

                var ID = Excel.NewRow(act.ActorType, decodedFieldValues);

                var retAct = GenerateActorFromExcelRow(Excel.GetRowByID(act.ActorType, ID.ToString()), act.ActorType, act.FieldValues.Keys.ToArray<string>());
                ri.IsOk = true;
                ri.Actor = retAct;
            }
            else
            {
                ri.IsOk = true;
                ri.State = ResponseState.Warning;
                ri.UserExplanation = "Connection closed.";
                ri.TechExplanation = "Could not find open Excel application and/or open TestConnector Excel document";
            }

            return ri;
        }

        public ActorArrayPluginResponse SaveActors(ErpActor[] actors)
        {
            var retActors = new List<ErpActor>();
            var ri = new ActorArrayPluginResponse();

            if (TestConnectionStatus())
            {
                foreach (var act in actors)
                {
                    var decodedFieldValues = new Dictionary<string, object>();

                    foreach (var fld in act.FieldValues)
                    {
                        decodedFieldValues.Add(TranslateFieldFromCrm(act.ActorType, fld.Key), CultureDataFormatter.ParseEncoded(fld.Value));
                    }

                    if (act.ActorType == "Person")
                    {
                        decodedFieldValues.Add("ParentID", act.ParentErpKey);
                        decodedFieldValues.Add("ParentType", act.ParentActorType);
                    }

                    Excel.UpdateRowByID(act.ActorType, act.ErpKey, decodedFieldValues);
                    retActors.Add(GenerateActorFromExcelRow(Excel.GetRowByID(act.ActorType, act.ErpKey), act.ActorType, act.FieldValues.Keys.ToArray<string>()));
                }

                ri.IsOk = true;
                ri.Actors = retActors.ToArray();
            }
            else
            {
                ri.IsOk = true;
                ri.State = ResponseState.Warning;
                ri.UserExplanation = "Connection closed.";
                ri.TechExplanation = "Could not find open Excel application and/or open TestConnector Excel document";
            }

            return ri;
        }

        public ListItemArrayPluginResponse GetList(string listName)
        {
            var listItems = new Dictionary<string, string>();

            var retActors = new List<ErpActor>();
            var ri = new ListItemArrayPluginResponse();

            if (TestConnectionStatus())
            {
                var rows = Excel.GetAllRows(listName);

                foreach (var rw in rows)
                {
                    if (rw.ContainsKey("ID") && rw.ContainsKey("Text"))
                        listItems.Add(rw["ID"].ToString(), rw["Text"].ToString());
                }

                ri.IsOk = true;
            }
            else
            {
                ri.IsOk = true;
                ri.State = ResponseState.Warning;
                ri.UserExplanation = "Connection closed.";
                ri.TechExplanation = "Could not find open Excel application and/or open TestConnector Excel document";
            }

            ri.ListItems = listItems;

            return ri;
        }

        public ListItemArrayPluginResponse GetListItems(string listName, string[] listItemKeys)
        {
            // Quick and dirty, but hey... It's a test.
            var riList = GetList(listName);

            if (!riList.IsOk)
                return riList;

            var listItems = new Dictionary<string, string>();

            foreach (var key in listItemKeys)
            {
                if (riList.ListItems.ContainsKey(key))
                    listItems.Add(key, riList.ListItems[key]);
            }

            riList.ListItems = listItems;

            return riList;
        }

        public ActorArrayPluginResponse GetActorsByTimestamp(string updatedOnOrAfter, string actorType, string[] fieldKeys)
        {
            var actors = new List<ErpActor>();
            var ri = new ActorArrayPluginResponse();

            if (TestConnectionStatus())
            {
                var dt = new DateTime(1970, 1, 1);

                try
                {
                    dt = CultureDataFormatter.ParseEncodedDate(updatedOnOrAfter);
                }
                catch (Exception)
                { }

                // TODO: Getting all rows and then filtering is obviously horribly inefficient, but it'll work for now
                var rows = Excel.GetAllRows(actorType);

                foreach (var row in rows)
                {
                    var act = GenerateActorFromExcelRow(row, actorType, fieldKeys);
                    var actDt = CultureDataFormatter.ParseEncodedDate(act.LastModified);

                    if (actDt >= dt)
                        actors.Add(act);
                }

                ri.IsOk = true;
                ri.Actors = actors.ToArray();
            }
            else
            {
                ri.IsOk = true;
                ri.State = ResponseState.Warning;
                ri.UserExplanation = "Connection closed.";
                ri.TechExplanation = "Could not find open Excel application and/or open TestConnector Excel document";
            }

            return ri;
        }

        private ErpActor GenerateActorFromExcelRow(Dictionary<string, object> row, string actorType, string[] fieldKeys)
        {
            ErpActor act = null;

            if (row != null && row.Count > 0)
            {
                act = new ErpActor();

                act.ActorType = actorType;

                FieldMetadataInfo[] fields = null;

                switch (actorType)
                {
                    case "Customer":
                        fields = GetCustomerFields();
                        break;
                    case "Supplier":
                        fields = GetSupplierFields();
                        break;
                    case "Person":
                        fields = GetPersonFields();
                        break;
                    case "Project":
                        fields = GetProjectFields();
                        break;
                }

                if (row.ContainsKey("ID") && row["ID"] != null)
                    act.ErpKey = row["ID"].ToString();

                if (row.ContainsKey("LastModified") && row["LastModified"] != null)
                {
                    var dtVal = DateTime.MinValue;

                    if (row["LastModified"] is DateTime)
                        dtVal = (DateTime)row["LastModified"];

                    //act.LastModified = CultureDataFormatter.EncodeDateTime(dtVal);
                    act.LastModified = dtVal.ToString("s"); // To sortable string
                }

                // Parent actor information
                if (actorType == "Person")
                {
                    if (row.ContainsKey("ParentID") && row["ParentID"] != null)
                        act.ParentErpKey = row["ParentID"].ToString();
                    if (row.ContainsKey("ParentType") && row["ParentType"] != null)
                        act.ParentActorType = row["ParentType"].ToString();
                }

                foreach (var fieldKey in fieldKeys)
                {
                    if (row.ContainsKey(TranslateFieldFromCrm(actorType, fieldKey)))
                    {
                        var originalValue = row[TranslateFieldFromCrm(actorType, fieldKey)];

                        if (originalValue == null)
                            originalValue = "";

                        var encodedValue = "";

                        if (fields != null)
                        {
                            var fld = (
                                from f in fields
                                where f.FieldKey == fieldKey
                                select f).FirstOrDefault();

                            if (fld != null)
                            {
                                switch (fld.FromPlugin().FieldType)
                                {
                                    case FieldMetadataTypeInfoWS.Checkbox:
                                        var numVal = 0;
                                        if (originalValue is bool)
                                            numVal = (bool)originalValue == true ? 1 : 0;
                                        else if (originalValue.ToString() == "1")
                                            numVal = 1;

                                        encodedValue = CultureDataFormatter.EncodeInt(numVal);
                                        break;

                                    case FieldMetadataTypeInfoWS.Datetime:
                                        var dtVal = DateTime.MinValue;
                                        if (originalValue is DateTime)
                                            dtVal = (DateTime)originalValue;
                                        else if (originalValue is double) // Stupid Excel date format
                                            dtVal = DateTime.FromOADate((double)originalValue);
                                        else
                                            if (!DateTime.TryParse(originalValue.ToString(), out dtVal))
                                            dtVal = DateTime.MinValue;

                                        encodedValue = CultureDataFormatter.EncodeDateTime(dtVal);
                                        break;

                                    case FieldMetadataTypeInfoWS.Double:
                                        var dblVal = 0.0;
                                        if (originalValue is double)
                                            dblVal = (double)originalValue;
                                        else
                                            if (!double.TryParse(originalValue.ToString(), out dblVal))
                                            dblVal = 0.0;

                                        encodedValue = CultureDataFormatter.EncodeDouble(dblVal);
                                        break;

                                    case FieldMetadataTypeInfoWS.Integer:
                                        var intVal = 0;
                                        if (originalValue is int)
                                            intVal = (int)originalValue;
                                        else if (originalValue is double)// Excel handler returns doubles for all numeric values
                                            intVal = Convert.ToInt32((double)originalValue);
                                        else
                                            if (!int.TryParse(originalValue.ToString(), out intVal))
                                            intVal = 0;

                                        encodedValue = CultureDataFormatter.EncodeInt(intVal);
                                        break;

                                    case FieldMetadataTypeInfoWS.List: // List IDs are shipped as strings, regardless of their original type
                                    case FieldMetadataTypeInfoWS.Text:
                                    case FieldMetadataTypeInfoWS.Password:
                                    case FieldMetadataTypeInfoWS.Label:
                                    default:
                                        encodedValue = originalValue.ToString();
                                        break;
                                }
                            }
                            //encodedValue = CultureDataFormatter.FormatFromMetadata(encodedValue, fld.FromPlugin());
                        }

                        act.FieldValues.Add(fieldKey, encodedValue);
                    }
                }
            }

            return act;
        }

        private Dictionary<string, string> CustomerfieldToCrm = new Dictionary<string, string>();
        private Dictionary<string, string> SupplierfieldToCrm = new Dictionary<string, string>();
        private Dictionary<string, string> PersonfieldToCrm = new Dictionary<string, string>();
        private Dictionary<string, string> ProjectfieldToCrm = new Dictionary<string, string>();

        private Dictionary<string, string> CustomerfieldFromCrm = new Dictionary<string, string>();
        private Dictionary<string, string> SupplierfieldFromCrm = new Dictionary<string, string>();
        private Dictionary<string, string> PersonfieldFromCrm = new Dictionary<string, string>();
        private Dictionary<string, string> ProjectfieldFromCrm = new Dictionary<string, string>();

        private void initFieldMappings()
        {
            CustomerfieldToCrm.Clear();
            CustomerfieldToCrm.Add("Nm", "NAME");
            CustomerfieldToCrm.Add("Ad1", "POSTALAD1");
            CustomerfieldToCrm.Add("Ad2", "POSTALAD2");
            CustomerfieldToCrm.Add("Zip", "POSTALZIP");
            CustomerfieldToCrm.Add("City", "POSTALCITY");

            CustomerfieldFromCrm = CustomerfieldToCrm.ToDictionary(x => x.Value, x => x.Key);
            CustomerfieldToCrm = CustomerfieldToCrm.ToDictionary(x => x.Key.ToUpper(), x => x.Value);

            SupplierfieldToCrm.Clear();
            SupplierfieldToCrm.Add("Nm", "NAME");
            SupplierfieldToCrm.Add("Ad1", "POSTALAD1");
            SupplierfieldToCrm.Add("Ad2", "POSTALAD2");
            SupplierfieldToCrm.Add("Zip", "POSTALZIP");
            SupplierfieldToCrm.Add("City", "POSTALCITY");

            SupplierfieldFromCrm = SupplierfieldToCrm.ToDictionary(x => x.Value, x => x.Key);
            SupplierfieldToCrm = SupplierfieldToCrm.ToDictionary(x => x.Key.ToUpper(), x => x.Value);

            PersonfieldToCrm.Clear();
            PersonfieldToCrm.Add("FirstName", "FIRSTNAME");
            PersonfieldToCrm.Add("LastName", "LASTNAME");
            PersonfieldToCrm.Add("Address", "POSTALAD1");
            PersonfieldToCrm.Add("Phone", "PHONE_DIRECT");

            PersonfieldFromCrm = PersonfieldToCrm.ToDictionary(x => x.Value, x => x.Key);
            PersonfieldToCrm = PersonfieldToCrm.ToDictionary(x => x.Key.ToUpper(), x => x.Value);


            ProjectfieldToCrm.Clear();
            ProjectfieldToCrm.Add("Nm", "NAME");
            ProjectfieldToCrm.Add("EndDate", "ENDDATE");
            ProjectfieldToCrm.Add("Description", "TEXT");

            ProjectfieldFromCrm = ProjectfieldToCrm.ToDictionary(x => x.Value, x => x.Key);
            ProjectfieldToCrm = ProjectfieldToCrm.ToDictionary(x => x.Key.ToUpper(), x => x.Value);

        }

        private string TranslateFieldFromCrm(string actorType, string fieldName)
        {
            if (CustomerfieldToCrm.Count <= 0)
                initFieldMappings();

            var foundField = string.Empty;
            switch (actorType)
            {
                case "Customer":
                    foundField = CustomerfieldFromCrm.ContainsKey(fieldName.ToUpper()) ? CustomerfieldFromCrm[fieldName.ToUpper()] : string.Empty;
                    break;
                case "Supplier":
                    foundField = SupplierfieldFromCrm.ContainsKey(fieldName.ToUpper()) ? SupplierfieldFromCrm[fieldName.ToUpper()] : string.Empty;
                    break;
                case "Person":
                    foundField = PersonfieldFromCrm.ContainsKey(fieldName.ToUpper()) ? PersonfieldFromCrm[fieldName.ToUpper()] : string.Empty;
                    break;
                case "Project":
                    foundField = ProjectfieldFromCrm.ContainsKey(fieldName.ToUpper()) ? ProjectfieldFromCrm[fieldName.ToUpper()] : string.Empty;
                    break;
            }

            return string.IsNullOrEmpty(foundField) ? fieldName : foundField;
        }

        private string TranslateFieldToCrm(string actorType, string fieldName)
        {
            if (CustomerfieldToCrm.Count <= 0)
                initFieldMappings();

            var foundField = string.Empty;
            switch (actorType)
            {
                case "Customer":
                    foundField = CustomerfieldToCrm.ContainsKey(fieldName.ToUpper()) ? CustomerfieldToCrm[fieldName.ToUpper()] : string.Empty;
                    break;
                case "Supplier":
                    foundField = SupplierfieldToCrm.ContainsKey(fieldName.ToUpper()) ? SupplierfieldToCrm[fieldName.ToUpper()] : string.Empty;
                    break;
                case "Person":
                    foundField = PersonfieldToCrm.ContainsKey(fieldName.ToUpper()) ? PersonfieldToCrm[fieldName.ToUpper()] : string.Empty;
                    break;
                case "Project":
                    foundField = ProjectfieldToCrm.ContainsKey(fieldName.ToUpper()) ? ProjectfieldToCrm[fieldName.ToUpper()] : string.Empty;
                    break;
            }

            return string.IsNullOrEmpty(foundField) ? fieldName : foundField;
        }

        private FieldMetadataInfo[] GetCustomerFields()
        {
            var fields = new List<FieldMetadataInfo>();

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.ReadOnly,
                DefaultValue = "",
                DisplayDescription = "NO:\"Kunde nr.\";US:\"Customer no.\"",
                DisplayName = "NO:\"Kunde nummer\";US:\"Customer number\"",
                FieldKey = "CustNo",
                FieldType = FieldMetadataTypeInfo.Integer,
                ListName = "",
                MaxLength = 4
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Mandatory,
                DefaultValue = "",
                DisplayDescription = "Customer name",
                DisplayName = "US:\"Name\";NO:\"Navn\"",
                FieldKey = "Nm",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Customer address, line 1",
                DisplayName = "Address 1",
                FieldKey = "Ad1",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Customer address, line 2",
                DisplayName = "Address 2",
                FieldKey = "Ad2",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Customer zip code",
                DisplayName = "Zip code",
                FieldKey = "Zip",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 20
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Customer city",
                DisplayName = "US:\"City\";NO:\"By\"",
                FieldKey = "City",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 50
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Customer group",
                DisplayName = "Customer group",
                FieldKey = "CustGr",
                FieldType = FieldMetadataTypeInfo.List,
                ListName = "CustomerGroup",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Customer generic integer field",
                DisplayName = "Integer field",
                FieldKey = "IntField",
                FieldType = FieldMetadataTypeInfo.Integer,
                ListName = "",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Customer generic double/decimal field",
                DisplayName = "Double field",
                FieldKey = "DoubleField",
                FieldType = FieldMetadataTypeInfo.Double,
                ListName = "",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Customer generic date field",
                DisplayName = "Date field",
                FieldKey = "DateField",
                FieldType = FieldMetadataTypeInfo.Datetime,
                ListName = "",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Customer generic checkbox field",
                DisplayName = "Checkbox field",
                FieldKey = "CheckboxField",
                FieldType = FieldMetadataTypeInfo.Checkbox,
                ListName = "",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Category ERP",
                DisplayName = "Category ERP",
                FieldKey = "CatErp",
                FieldType = FieldMetadataTypeInfo.List,
                ListName = "CustomerCategory",
                MaxLength = 0
            });


            foreach (var f in fields)
            {
                f.FieldKey = TranslateFieldToCrm("Customer", f.FieldKey);
            }

            return fields.ToArray();
        }

        private FieldMetadataInfo[] GetSupplierFields()
        {
            var fields = new List<FieldMetadataInfo>();

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.ReadOnly,
                DefaultValue = "",
                DisplayDescription = "Supplier number",
                DisplayName = "US:\"Supplier no.\";NO:\"Leverandør nr\"",
                FieldKey = "SupNo",
                FieldType = FieldMetadataTypeInfo.Integer,
                ListName = "",
                MaxLength = 4
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Mandatory,
                DefaultValue = "",
                DisplayDescription = "Supplier name",
                DisplayName = "US:\"Name\";NO:\"Navn\"",
                FieldKey = "Nm",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Supplier address, line 1",
                DisplayName = "Address 1",
                FieldKey = "Ad1",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Supplier address, line 2",
                DisplayName = "Address 2",
                FieldKey = "Ad2",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Supplier zip code",
                DisplayName = "Zip code",
                FieldKey = "Zip",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 20
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Supplier city",
                DisplayName = "City",
                FieldKey = "City",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 50
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Supplier group",
                DisplayName = "Supplier group",
                FieldKey = "SupGr",
                FieldType = FieldMetadataTypeInfo.List,
                ListName = "SupplierGroup",
                MaxLength = 4
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Supplier generic integer field",
                DisplayName = "Integer field",
                FieldKey = "IntField",
                FieldType = FieldMetadataTypeInfo.Integer,
                ListName = "",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Supplier generic double/decimal field",
                DisplayName = "Double field",
                FieldKey = "DoubleField",
                FieldType = FieldMetadataTypeInfo.Double,
                ListName = "",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Supplier generic date field",
                DisplayName = "Date field",
                FieldKey = "DateField",
                FieldType = FieldMetadataTypeInfo.Datetime,
                ListName = "",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Supplier generic checkbox field",
                DisplayName = "Checkbox field",
                FieldKey = "CheckboxField",
                FieldType = FieldMetadataTypeInfo.Checkbox,
                ListName = "",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Category SupplierERP",
                DisplayName = "Category SupplierERP",
                FieldKey = "CatSupErp",
                FieldType = FieldMetadataTypeInfo.List,
                ListName = "SupplierCategory",
                MaxLength = 0
            });

            foreach (var f in fields)
            {
                f.FieldKey = TranslateFieldToCrm("Supplier", f.FieldKey);
            }

            return fields.ToArray();
        }

        private FieldMetadataInfo[] GetPersonFields()
        {
            var fields = new List<FieldMetadataInfo>();

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.ReadOnly,
                DefaultValue = "",
                DisplayDescription = "Person number",
                DisplayName = "Person no.",
                FieldKey = "PersonNo",
                FieldType = FieldMetadataTypeInfo.Integer,
                ListName = "",
                MaxLength = 4
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "First name",
                DisplayName = "First name",
                FieldKey = "FirstName",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Last name",
                DisplayName = "Last name",
                FieldKey = "LastName",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Address line",
                DisplayName = "Address",
                FieldKey = "Address",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Phone number",
                DisplayName = "Phone number",
                FieldKey = "Phone",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 50
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Position",
                DisplayName = "Position",
                FieldKey = "PersPos",
                FieldType = FieldMetadataTypeInfo.List,
                ListName = "PersonPosition",
                MaxLength = 50
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "PersonType",
                DisplayName = "PersonType",
                FieldKey = "PersType",
                FieldType = FieldMetadataTypeInfo.List,
                ListName = "PersonType",
                MaxLength = 50
            });

            foreach (var f in fields)
            {
                f.FieldKey = TranslateFieldToCrm("Person", f.FieldKey);
            }


            return fields.ToArray();
        }

        private FieldMetadataInfo[] GetProjectFields()
        {
            var fields = new List<FieldMetadataInfo>();

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.ReadOnly,
                DefaultValue = "",
                DisplayDescription = "Project number",
                DisplayName = "Project no.",
                FieldKey = "ProjNo",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 4
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Project name",
                DisplayName = "Name",
                FieldKey = "Nm",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 500
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "End date",
                DisplayName = "Project end date",
                FieldKey = "EndDate",
                FieldType = FieldMetadataTypeInfo.Datetime,
                ListName = "",
                MaxLength = 0
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Description",
                DisplayName = "Project text/description",
                FieldKey = "Description",
                FieldType = FieldMetadataTypeInfo.Text,
                ListName = "",
                MaxLength = 5000
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Project type",
                DisplayName = "Project type",
                FieldKey = "ProjType",
                FieldType = FieldMetadataTypeInfo.List,
                ListName = "ProjectType",
                MaxLength = 5000
            });

            fields.Add(new FieldMetadataInfo()
            {
                Access = FieldAccessInfo.Normal,
                DefaultValue = "",
                DisplayDescription = "Project Status",
                DisplayName = "Project status",
                FieldKey = "ProjStatus",
                FieldType = FieldMetadataTypeInfo.List,
                ListName = "ProjectStatus",
                MaxLength = 5000
            });

            foreach (var f in fields)
            {
                f.FieldKey = TranslateFieldToCrm("Project", f.FieldKey);
            }

            return fields.ToArray();
        }
    }
}
