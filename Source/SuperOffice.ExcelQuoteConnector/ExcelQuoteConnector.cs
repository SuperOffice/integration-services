using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SuperOffice.CRM;
using SuperOffice.Exceptions;
using SuperOffice.Util;
using OfficeOpenXml;

namespace SuperOffice.Connectors
{
    public static class Incapabilities
    {
        public const string SendQuoteUrl = "send_quote_url";
        public const string PlaceOrderUrl = "place_order_url";

        public const string SendQuoteSOProto = "send_quote_soproto";
        public const string PlaceOrderSOProto = "place_order_soproto";

        public const string CannotStart = "cannot_start";
        public const string FailStart = "fail_start";
        public const string WarnStart = "warn_start";

        public const string CannotCreateQuote = "cannot_create";
        public const string FailCreateQuote = "fail_create";
        public const string WarnCreateQuote = "warn_create";

        public const string CannotCreateQuoteVersion = "cannot_create_version";
        public const string FailCreateQuoteVersion = "fail_create_version";
        public const string WarnCreateQuoteVersion = "warn_create_version";

        public const string CannotCreateQuoteAlternative = "cannot_create_alternative";
        public const string FailCreateQuoteAlternative = "fail_create_alternative";
        public const string WarnCreateQuoteAlternative = "warn_create_alternative";

        public const string CannotFind = "cannot_find";

        public const string CannotProduct = "cannot_product";
        public const string FailProduct = "fail_product";
        public const string WarnProduct = "warn_product";

        public const string CannotSave = "cannot_save";
        public const string CannotDelete = "cannot_delete";

        public const string CannotRecalc = "cannot_recalc";
        public const string FailRecalc = "fail_recalc";
        public const string WarnRecalc = "warn_recalc";

        public const string CannotValidateVer = "cannot_validate_ver";
        public const string FailValidateVer = "fail_validate_ver";
        public const string WarnValidateVer = "warn_validate_ver";
        public const string NeedValidateVer = "need_validate_ver";

        public const string CannotValidateAlt = "cannot_validate_alt";
        public const string FailValidateAlt = "fail_validate_alt";
        public const string WarnValidateAlt = "warn_validate_alt";

        public const string CannotValidateLine = "cannot_validate_line";
        public const string FailValidateLine = "fail_validate_line";
        public const string WarnValidateLine = "warn_validate_line";

        public const string CannotUpdate = "cannot_update";
        public const string FailUpdate = "fail_update";
        public const string WarnUpdate = "warn_update";

        public const string CannotSendQuote = "cannot_send_quote";
        public const string FailSendQuote = "fail_send_quote";
        public const string WarnSendQuote = "warn_send_quote";

        public const string CannotPlaceOrder = "cannot_place_order";
        public const string FailPlaceOrder = "fail_place_order";

        public const string CannotOrderState = "cannot_order_state";
        public const string FailOrderState = "fail_order_state";
        public const string WarnOrderState = "warn_order_state";

        public const string RemoveFirstLineOnRecalc = "remove_first_line_on_recalc";
        public const string Replace90pctLinesOnRecalc = "replace_90pct_lines_on_recalc";    // replace any lines with more than 90% discount with a number of lines based on the quantity.

        public const string FailConfigureFieldRanks = "fail_configure_field_ranks";
        public const string FailConfigureFieldKeys = "fail_configure_field_keys";
    }

    public static class EPPlusExtensions
    {
        public static string ReadStringValue(this ExcelWorksheet sheet, int rownum, int colnum)
        {
            try
            {
                var value = sheet.Cells[rownum, colnum].Value?.ToString() ?? string.Empty;
                return value;
            }
            catch (Exception ex)
            {
                throw new Exception("Problem reading string value from row {0}, col {1} in sheet {2}".FormatWith(rownum, colnum, sheet.Name), ex);
            }
        }
        public static bool ReadBoolValue(this ExcelWorksheet sheet, int rownum, int colnum)
        {
            try
            {
                var value = false;
                var cell = sheet.Cells[rownum, colnum];
                value = Convert.ToBoolean(cell.Value ?? false);
                return value;
            }
            catch (Exception ex)
            {
                throw new Exception("Problem reading bool value from row {0}, col {1} in sheet {2}".FormatWith(rownum, colnum, sheet.Name), ex);
            }
        }
        public static int ReadIntValue(this ExcelWorksheet sheet, int rownum, int colnum)
        {
            try
            {
                var value = 0;
                var cell = sheet.Cells[rownum, colnum];
                if (cell.Value != null)
                {
                    value = Convert.ToInt32(cell.Value.ToString());
                }
                return value;
            }
            catch (Exception ex)
            {
                throw new Exception("Problem reading int value from row {0}, col {1} in sheet {2}".FormatWith(rownum, colnum, sheet.Name), ex);
            }
        }
        public static double ReadDoubleValue(this ExcelWorksheet sheet, int rownum, int colnum)
        {
            try
            {
                var value = 0.0;
                var cell = sheet.Cells[rownum, colnum];
                if (cell.Value != null)
                { 
                    Double.TryParse(cell.Value.ToString(), out value);
                }
                return value;
            }
            catch (Exception ex)
            {
                throw new Exception("Problem reading double value from row {0}, col {1} in sheet {2}".FormatWith(rownum, colnum, sheet.Name), ex);
            }
        }
        public static DateTime ReadDateTimeValue(this ExcelWorksheet sheet, int rownum, int colnum)
        {
            try
            {
                DateTime value;
                var cell = sheet.Cells[rownum, colnum];
                if (cell.Value == null)
                {
                    value = new DateTime();
                }                
                else if (cell.Value.GetType() == typeof(double))
                {
                    value = DateTime.FromOADate((double)cell.Value);
                }
                else
                {
                    value = Convert.ToDateTime(cell.Value);
                }
                return value;
            }
            catch (Exception ex)
            {
                throw new Exception("Problem reading date-time value from row {0}, col {1} in sheet {2}".FormatWith(rownum, colnum, sheet.Name), ex);
            }
        }
    }





    [QuoteConnector(Name)]
    public class ExcelQuoteConnector : QuoteConnectorBase
    {
        public const string Name = "ExcelQuoteConnector";

        public ExcelQuoteConnector()
        {
            ProductProvider = new InMemoryProductProvider();
        }

        public override Dictionary<string, FieldMetadataInfo> GetConfigurationFields()
        {
            var res = new Dictionary<string, FieldMetadataInfo>
                       {
                           {
                               "#1", new FieldMetadataInfo()
                                         {
                                             Access = FieldAccessInfo.Mandatory,
                                             DefaultValue = string.Empty,
                                             DisplayName =
                                                 "US:\"Excel file name\";NO:\"Excel filnavn\";GE:\"Excel Dateinamen\";FR:\"Nom de fichier Excel\"",
                                             DisplayDescription =
                                                 "US:\"The name of the Excel file to load pricelists and products from.\";NO:\"Navnet på Excel-fil for å laste prislister og produkter fra.\";GE:\"Der Name der Excel-Datei auf Preislisten und Produkte aus laden.\";FR:\"Le nom du fichier Excel pour charger des listes de prix et de produits à partir.\"",
                                             FieldType = FieldMetadataTypeInfo.Text,
                                             FieldKey = "#1",
                                             Rank = 1,
                                         }
                           },
                           {
                               "#2", new FieldMetadataInfo()
                                         {
                                             DisplayName = "[SR_ADMIN_IMPORT_ONLY_EXCEL]",
                                             DisplayDescription = "[SR_ADMIN_IMPORT_ERROR_XLS]",
                                             FieldType = FieldMetadataTypeInfo.Label,
                                             FieldKey = "#2",
                                             Rank = 2,
                                         }
                           },
                           {
                               "#3", new FieldMetadataInfo()
                                         {
                                             DisplayName = "[SR_LABEL_CATEGORY]",
                                             FieldType = FieldMetadataTypeInfo.List,
                                             ListName = "productcategory",
                                             FieldKey = "#3",
                                             Rank = 3,
                                         }
                           }

                       };
            if (CanProvideCapability(Incapabilities.FailConfigureFieldKeys))
            {
                res["#1"].FieldKey = "glops";
                res["#2"].FieldKey = "glips";
                res["#3"].FieldKey = "glups";
            }
            if (CanProvideCapability(Incapabilities.FailConfigureFieldRanks))
            {
                res["#1"].Rank = 0;
                res["#2"].Rank = 3;
                res["#3"].Rank = 3;
            }

            return res;
        }

        public override PluginResponseInfo TestConnection(Dictionary<string, string> connectionConfigFields)
        {
            return CheckConnectionData(connectionConfigFields);
        }


        private string Filename { get; set; }

        private PluginResponseInfo CheckConnectionData(Dictionary<string, string> connectionConfigFields)
        {
            if (connectionConfigFields != null && connectionConfigFields.Count() >= 1)
            {
                Filename = connectionConfigFields.First().Value;
                if (File.Exists(Filename))
                {
                    ReadInData();
                    return ProductProvider.ValidateData()
                                            .Merge(ValidateList(DeliveryTerms, "DeliveryTerms"))
                                            .Merge(ValidateList(DeliveryTypes, "DeliveryTypes"))
                                            .Merge(ValidateList(PaymentTerms, "PaymentTerms"))
                                            .Merge(ValidateList(PaymentTypes, "PaymentTypes"))
                                            .Merge(ValidateList(ProductCategories, "ProductCategories"));
                }
                else
                    return GetFileNotFoundResponse();
            }
            else
                return GetWrongConfigResponse();
        }

        private PluginResponseInfo ValidateList(List<ListItemInfo> list, string listname)
        {
            var uniqueItems = list.Select(item => item.ERPQuoteListItemKey).Distinct().ToArray();
            if (uniqueItems.Length != list.Count())
                return new PluginResponseInfo()
                {
                    IsOk = false
                    ,
                    State = ResponseState.Error
                    ,
                    UserExplanation = "US:\"{0}\";NO:\"{1}\";GE:\"{2}\";FR:\"{3}\"".FormatWith(
                                    "All listitems in '{0}' list must have an unique 'ERPQuoteListItemKey'.".FormatWith(listname)
                                    , "Alle listeelementer i '{0}' listen må ha en unik 'ERPQuoteListItemKey'.".FormatWith(listname)
                                    , "Alle Listenelemente in '{0}' Liste muss über eine eindeutige' ERPQuoteListItemKey '.".FormatWith(listname)
                                    , "Tous les éléments de la liste dans '{0}' liste doit avoir un unique 'ERPQuoteListItemKey'.".FormatWith(listname)
                                    )
                    ,
                    TechExplanation = "Some items in {0} had the same ERPQuoteListItemKey.".FormatWith(listname)
                };
            else
                return new PluginResponseInfo() { State = ResponseState.Ok };
        }

        private static PluginResponseInfo GetFileNotFoundResponse()
        {
            return new PluginResponseInfo()
            {
                IsOk = false,
                State = ResponseState.Error
                ,
                UserExplanation = "US:\"{0}\";NO:\"{1}\";GE:\"{2}\";FR:\"{3}\"".FormatWith(
                                  "You must enter an absolute file name, or a filename relative to where the ExcelConnector resides in the file hierarchy."
                                  , "Du må angi en absolutt filnavn, eller et filnavn i forhold til hvor ExcelConnectoren ligger i filhierarkiet."
                                  , "Sie müssen einen absoluten Dateinamen oder einen Dateinamen relativ zu denen die ExcelConnector befindet sich in der Datei-Hierarchie."
                                  , "Vous devez entrer un nom de fichier absolu, ou un nom de fichier relatif à celui du ExcelConnector réside dans la hiérarchie des fichiers."
                                  )
                ,
                TechExplanation = "File not found :-("
            };
        }

        private static PluginResponseInfo GetWrongConfigResponse()
        {
            return new PluginResponseInfo()
            {
                IsOk = false,
                State = ResponseState.Error,
                UserExplanation = "Technical error: Config parameters were null or not exactly one element long",
                TechExplanation = "ConnectionData was null or had not excactly one element"
            };
        }




        public override PluginResponseInfo InitializeConnection(QuoteConnectionInfo connectionData
                                                                    , UserInfo user
                                                                    , bool isOnTravel
                                                                    , Dictionary<string, string> connectionConfigFields
                                                                    , IProductRegisterCache productRegister)
        {
            var retv = CheckConnectionData(connectionConfigFields);

            if(!retv.IsOk)
                return retv;

            if (connectionConfigFields == null || connectionConfigFields.Count == 0)
                throw new SoException("Tried to initialize ExcelQuoteConnector with no connection config!");

            Filename = connectionConfigFields.First().Value;

            ReadInData();

            if (CanProvideCapability(Incapabilities.CannotStart))
                throw new SoIllegalOperationException("Cannot_Start: Excel Connector was OK before you touched it.");

            if (CanProvideCapability(Incapabilities.FailStart))
            {
                retv.State = ResponseState.Error;
                retv.TechExplanation = "Ionization from the air-conditioning";
                retv.UserExplanation = "Fail_Start: You did wha???...     oh _dear_....";
            }
            if (CanProvideCapability(Incapabilities.WarnStart))
            {
                retv.State = ResponseState.Warning;
                retv.TechExplanation = "Too much radiation coming from the soil.";
                retv.UserExplanation = "Warn_Start: Unfortunately we have run out of bits/bytes/whatever.";
            }

            return retv;
        }


        public override ListItemInfo[] GetQuoteList(string quoteListType)
        {
            ReadInData();

            ListItemInfo[] retv = null;

            var cap = "ilistprovider_provide_{0}list".FormatWith(quoteListType.ToLower());

            if (CanProvideCapability(cap))
            {
                switch (quoteListType.ToLower())
                {
                    case "productcategory":
                        retv = ProductCategories.ToArray();
                        break;
                    case "productfamily":
                        retv = ProductFamilies.ToArray();
                        break;
                    case "producttype":
                        retv = ProductTypes.ToArray();
                        break;

                    case "paymentterms":
                        retv = PaymentTerms.ToArray();
                        break;
                    case "paymenttype":
                        retv = PaymentTypes.ToArray();
                        break;
                    case "deliveryterms":
                        retv = DeliveryTerms.ToArray();
                        break;
                    case "deliverytype":
                        retv = DeliveryTypes.ToArray();
                        break;
                    default:
                        // just return null
                        break;
                }
            }

            return retv;
        }



        /// <summary>
        /// Check if one named capability can be provided (now)
        /// The excel connector basically can do anything, just set it up in the "Capability" excel sheet
        /// </summary>
        /// <param name="capabilityName"></param>
        /// <returns></returns>
        public override bool CanProvideCapability(string capabilityName)
        {
            var retv = false;
            if (Capabilities != null && Capabilities.Count() != 0 && Capabilities.ContainsKey(capabilityName))
            {
                retv = Capabilities[capabilityName];
            }
            else
            {
                switch (capabilityName)
                {
                    case CRMQuoteConnectorCapabilities.CanProvideCost:
                    case CRMQuoteConnectorCapabilities.CanProvideMinimumPrice:
                    case CRMQuoteConnectorCapabilities.CanProvidePicture:
                    case CRMQuoteConnectorCapabilities.CanProvideExtraData:
                    case CRMQuoteConnectorCapabilities.CanProvideStockData:
                        retv = true;
                        break;

                    case CRMQuoteConnectorCapabilities.CanPlaceOrder:
                    case CRMQuoteConnectorCapabilities.CanProvideOrderState:
                    case CRMQuoteConnectorCapabilities.CanSendOrderConfirmation:
                        retv = false;
                        break;

                    case CRMQuoteConnectorCapabilities.CanProvideProductCategoryList:
                        retv = true;
                        break;

                    case CRMQuoteConnectorCapabilities.CanProvideProductFamilyList:
                    case CRMQuoteConnectorCapabilities.CanProvideProductTypeList:
                        retv = false;
                        break;

                    case CRMQuoteConnectorCapabilities.CanProvidePaymentTermsList:
                    case CRMQuoteConnectorCapabilities.CanProvidePaymentTypeList:
                    case CRMQuoteConnectorCapabilities.CanProvideDeliveryTermsList:
                    case CRMQuoteConnectorCapabilities.CanProvideDeliveryTypeList:
                        retv = true;
                        break;

                    case CRMQuoteConnectorCapabilities.CanPerformComplexSearch:
                        retv = false;
                        break;

                    case CRMQuoteConnectorCapabilities.CanProvideAddresses:
                        retv = true;
                        break;
                }
            }

            return retv;
        }

        public override QuoteResponseInfo OnBeforeCreateQuote(QuoteAlternativeContextInfo context)
        {
            ReadInData();
            if (CanProvideCapability(Incapabilities.CannotCreateQuote))
                throw new SoClassFactoryException("Cannot_Create_Quote: Little hamster in running wheel had coronary; waiting for replacement to be Fedexed from Wyoming");
            var resp = base.OnBeforeCreateQuote(context);
            if (CanProvideCapability(Incapabilities.FailCreateQuote))
            {
                resp.IsOk = false;
                resp.UserExplanation = "Fail_Create_Quote: Hamster died";
                resp.TechExplanation = "Hamster server error 0x404";
            }
            if (CanProvideCapability(Incapabilities.WarnCreateQuote))
            {
                resp.State = ResponseState.Warning;
                resp.UserExplanation = "Warn_Create_Quote: Hamster sick";
                resp.TechExplanation = "Hamster server error 0x301";
            }
            return resp;
        }

        public override QuoteVersionResponseInfo OnBeforeCreateQuoteVersion(QuoteVersionContextInfo context)
        {
            ReadInData();
            if (CanProvideCapability(Incapabilities.CannotCreateQuoteVersion))
                throw new SoClassFactoryException("Cannot_Create_Version: Little hamster in running wheel had coronary; waiting for replacement to be Fedexed from Wyoming");
            var resp = base.OnBeforeCreateQuoteVersion(context);
            if (CanProvideCapability(Incapabilities.FailCreateQuoteVersion))
            {
                resp.IsOk = false;
                resp.UserExplanation = "Fail_Create_Version: Hamster died";
                resp.TechExplanation = "Hamster server error 0x404";
            }
            if (CanProvideCapability(Incapabilities.WarnCreateQuoteVersion))
            {
                resp.State = ResponseState.Warning;
                resp.UserExplanation = "Warn_Create_Version: Hamster sick";
                resp.TechExplanation = "Hamster server error 0x301";
            }
            // Fill in some random data to simulate ERP keys
            if (string.IsNullOrEmpty(resp.CRMQuoteVersion.ERPQuoteVersionKey))
                resp.CRMQuoteVersion.ERPQuoteVersionKey = GenKey();
            return resp;
        }

        public override QuoteAlternativeResponseInfo OnBeforeCreateQuoteAlternative(QuoteAlternativeContextInfo context)
        {
            ReadInData();
            if (CanProvideCapability(Incapabilities.CannotCreateQuoteAlternative))
                throw new SoClassFactoryException("Cannot_Createe_Alternative: Little hamster in running wheel had coronary; waiting for replacement to be Fedexed from Wyoming");
            var resp = base.OnBeforeCreateQuoteAlternative(context);
            if (CanProvideCapability(Incapabilities.FailCreateQuoteAlternative))
            {
                resp.IsOk = false;
                resp.UserExplanation = "Fail_Create_Alternative: Hamster died";
                resp.TechExplanation = "Hamster server error 0x404";
            }
            if (CanProvideCapability(Incapabilities.WarnCreateQuoteAlternative))
            {
                resp.State = ResponseState.Warning;
                resp.UserExplanation = "Warn_Create_Alternative: Hamster sick";
                resp.TechExplanation = "Hamster server error 0x301";
            }
            // Fill in some random data to simulate ERP keys
            if (resp.CRMAlternativesWithLines != null && string.IsNullOrEmpty(resp.CRMAlternativesWithLines[0].CRMAlternative.ERPQuoteAlternativeKey))
                resp.CRMAlternativesWithLines[0].CRMAlternative.ERPQuoteAlternativeKey = GenKey();
            return resp;
        }

        public override void OnAfterSaveQuote(CRM.QuoteAlternativeContextInfo context)
        {
            ReadInData();
            if (CanProvideCapability(Incapabilities.CannotSave))
                throw new SoDbAccessException("Cannot_Save: Little hamster in running wheel had coronary; waiting for replacement to be Fedexed from Wyoming");
        }

        public override void OnBeforeDeleteQuote(CRM.QuoteInfo quote, CRM.ISaleInfo sale, IContactInfo contact)
        {
            if (CanProvideCapability(Incapabilities.CannotDelete))
                throw new SoDbAccessException("Cannot_Delete: Unable to kill the little hamster in running wheel because hamster is missing.");
        }


        public override QuoteSentResponseInfo OnAfterSentQuoteVersion(CRM.QuoteVersionContextInfo context)
        {
            ReadInData();
            var res = new QuoteSentResponseInfo();
            if (CanProvideCapability(Incapabilities.SendQuoteUrl))
                res.Url = "http://www.visma.no/";
            if (CanProvideCapability(Incapabilities.SendQuoteSOProto))
                res.Url = "superoffice:contact.main?contact_id=2";

            res.VersionResponse = new QuoteVersionResponseInfo(context);

            if (CanProvideCapability(Incapabilities.CannotSendQuote))
                throw new SoDbAccessException("Cannot_Send_Quote: Could not find mailbox");
            if (CanProvideCapability(Incapabilities.FailSendQuote))
            {
                res.VersionResponse.IsOk = false;
                res.VersionResponse.UserExplanation = "Fail_Send_Quote: Mailbox on fire.";
                res.VersionResponse.TechExplanation = "Mailbox is undergoing exothermic combustion.";
            }

            if (res.VersionResponse.IsOk)
                res.VersionResponse.CRMQuoteVersion.ERPQuoteVersionKey = AddQuote(context);


            return res;
        }



        public override FieldMetadataInfo[] GetSearchableFields()
        {
            var result = new FieldMetadataInfo[6];
            result[0] = new FieldMetadataInfo() { Access = FieldAccessInfo.Mandatory, DefaultValue = "", DisplayDescription = "Mandatory field", DisplayName = "Mandatory", FieldKey = "man", FieldType = FieldMetadataTypeInfo.Text, MaxLength = 10, Rank = 5 };
            result[1] = new FieldMetadataInfo() { Access = FieldAccessInfo.Normal, DefaultValue = "123", DisplayDescription = "Number field", DisplayName = "Number", FieldKey = "num", FieldType = FieldMetadataTypeInfo.Integer, MaxLength = 3, Rank = 4 };
            result[2] = new FieldMetadataInfo() { Access = FieldAccessInfo.Normal, DefaultValue = "", DisplayDescription = "Checkbox field", DisplayName = "Checkbox", FieldKey = "chk", FieldType = FieldMetadataTypeInfo.Checkbox, Rank = 3 };
            result[3] = new FieldMetadataInfo() { Access = FieldAccessInfo.Normal, DefaultValue = "", DisplayDescription = "List field", DisplayName = "List", FieldKey = "lst", FieldType = FieldMetadataTypeInfo.List, ListName = "category", Rank = 2 };
            result[4] = new FieldMetadataInfo() { Access = FieldAccessInfo.Normal, DefaultValue = "", DisplayDescription = "Double field", DisplayName = "Decimal", FieldKey = "dec", FieldType = FieldMetadataTypeInfo.Double, Rank = 1 };
            result[5] = new FieldMetadataInfo() { Access = FieldAccessInfo.Normal, DefaultValue = "", DisplayDescription = "Date field", DisplayName = "Date", FieldKey = "dat", FieldType = FieldMetadataTypeInfo.Datetime, Rank = 0 };

            return result;
        }

        public override ProductInfo[] GetSearchResults(SearchRestrictionInfo[] restrictions)
        {
            ReadInData();

            return ProductProvider.FindProduct(null, ProductProvider.PriceLists[0].Currency, " ", "");
        }





        protected InMemoryProductProvider ProductProvider { get; set; }


        public override int GetNumberOfActivePriceLists(string isoCurrencyCode)
        {
            ReadInData();
            return ProductProvider.GetNumberOfActivePriceLists(isoCurrencyCode);
        }

        public override PriceListInfo[] GetActivePriceLists(string isoCurrencyCode)
        {
            ReadInData();
            return ProductProvider.GetActivePriceLists(isoCurrencyCode);
        }

        public override PriceListInfo[] GetAllPriceLists(string isoCurrencyCode)
        {
            ReadInData();
            return ProductProvider.GetAllPriceLists(isoCurrencyCode);
        }

        public override ProductInfo[] FindProduct(QuoteAlternativeContextInfo context, string currencyCode, string userinput, string priceListKey)
        {
            ReadInData();
            if (CanProvideCapability(Incapabilities.CannotFind))
                throw new SoException("Cannot_Find: Unable to connect to remote host: Connection refused");
            var res = ProductProvider.FindProduct(context, currencyCode, userinput, priceListKey);
            return res;
        }

        public override ProductInfo GetProduct(QuoteAlternativeContextInfo context, string erpProductKey)
        {
            ReadInData();
            var product = ProductProvider.GetProduct(context, erpProductKey);
            if (!CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvideCost))
                product.UnitCost = 0;
            if (!CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvideMinimumPrice))
                product.UnitMinimumPrice = 0;
            return product;
        }

        public override ProductInfo[] GetProducts(QuoteAlternativeContextInfo context, string[] erpProductKeys)
        {
            ReadInData();
            return ProductProvider.GetProducts(context, erpProductKeys);
        }

        public override QuoteLineInfo[] GetQuoteLinesFromProduct(QuoteAlternativeContextInfo context, string erpProductKey)
        {
            ReadInData();
            if (CanProvideCapability(Incapabilities.CannotProduct))
                throw new SoIllegalOperationException("Cannot_Product exception.");
            var products = ProductProvider.GetQuoteLinesFromProduct(context, erpProductKey);
            foreach (var product in products)
            {
                if (!CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvideCost))
                    product.UnitCost = 0;
                if (!CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvideMinimumPrice))
                    product.UnitMinimumPrice = 0;
                if (CanProvideCapability(Incapabilities.FailProduct))
                {
                    product.Status = QuoteStatusInfo.Error;
                    product.Reason = "Fail_Product error: fat electrons in the lines";
                }
                if (CanProvideCapability(Incapabilities.WarnProduct))
                {
                    product.Status = QuoteStatusInfo.Warning;
                    product.Reason = "Warn_Product warning: IRQ-problems with the Un-Interruptible-Power-Supply";
                }
                // fill in random data for ERP field
                if (string.IsNullOrEmpty(product.ERPQuoteLineKey))
                    product.ERPQuoteLineKey = GenKey();
            }
            return products;
        }

        public override int GetNumberOfProductImages(string erpProductKey)
        {
            ReadInData();
            if (!CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvidePicture))
                return 0;
            return ProductProvider.GetNumberOfProductImages(erpProductKey);
        }

        public override string GetProductImage(string erpProductKey, int rank)
        {
            ReadInData();
            if (!CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvidePicture))
                return null;
            return ProductProvider.GetProductImage(erpProductKey, rank);
        }



        public override QuoteLineInfo OnQuoteLineChanged(QuoteAlternativeContextInfo context, QuoteLineInfo ql, string[] changedFields)
        {
            ql.ClearStatus();

            return base.OnQuoteLineChanged(context, ql, changedFields);
        }

        public override QuoteAlternativeWithLinesInfo RecalculateQuoteAlternative(QuoteAlternativeContextInfo inContext)
        {
            if (CanProvideCapability(Incapabilities.CannotRecalc))
                throw new OverflowException("Cannot_recalc: 9999999999999999999999 is too big");

            if (CanProvideCapability(Incapabilities.FailRecalc))
            {
                inContext.CRMAlternativeWithLines.CRMAlternative.Status = QuoteStatusInfo.Error;
                inContext.CRMAlternativeWithLines.CRMAlternative.Reason = "Fail Recalc: divide by zero error.";
                return inContext.CRMAlternativeWithLines;
            }


            if (CanProvideCapability(Incapabilities.Replace90pctLinesOnRecalc))
            {
                var lst = inContext.CRMAlternativeWithLines.CRMQuoteLines.ToList();
                var killList = new List<QuoteLineInfo>();
                var addList = new List<QuoteLineInfo>();
                foreach (var i in lst)
                {
                    if (i.DiscountPercent >= 90)
                    {
                        killList.Add(i);
                        for (var n = 0; n < i.Quantity; n++)
                        {
                            var newItem = new QuoteLineInfo()
                            {
                                Code = i.Code,
                                DeliveredQuantity = 1.0,
                                Description = i.Description,
                                DiscountAmount = 0,
                                DiscountPercent = 0,
                                EarningAmount = 0,
                                EarningPercent = 0,
                                ERPDiscountAmount = 0,
                                ERPDiscountPercent = 0,
                                ERPProductKey = i.ERPProductKey,
                                ERPQuoteLineKey = i.ERPQuoteLineKey,
                                Name = i.Name + " #" + n.ToString(),
                                Quantity = i.Quantity,
                                Rank = n + lst.Count,
                                SubTotal = 10,
                                TotalPrice = 10,
                                UnitCost = 10,
                                UnitListPrice = 10,
                                UnitMinimumPrice = 10,
                            };

                            newItem.PriceUnit = i.PriceUnit;
                            newItem.QuantityUnit = i.QuantityUnit;
                            addList.Add(newItem);
                        }
                    }
                }
                foreach (var i in killList)
                    lst.Remove(i);
                foreach (var i in addList)
                    lst.Add(i);

                inContext.CRMAlternativeWithLines.CRMQuoteLines = lst.ToArray();
            }
            if (CanProvideCapability(Incapabilities.RemoveFirstLineOnRecalc))
            {
                var lst = inContext.CRMAlternativeWithLines.CRMQuoteLines.ToList();

                lst.Remove(inContext.CRMAlternativeWithLines.CRMQuoteLines.First());

                inContext.CRMAlternativeWithLines.CRMQuoteLines = lst.ToArray();
            }

            var result = base.RecalculateQuoteAlternative(inContext);

            if (CanProvideCapability(Incapabilities.WarnRecalc))
            {
                result.CRMAlternative.Reason = "Warn Recalc: 42 is a significant number, don't you think?";
                result.CRMAlternative.Status = QuoteStatusInfo.Warning;
            }

            return result;
        }

        private AddressInfo GetAddressInfo(QuoteAlternativeContextInfo context, AddressType type)
        {
            var address = from a in Addresses
                          where a.Key.ContactKey == context.CRMCompany.ContactId.ToString()
                                && a.Key.Type == type
                          select a.Value;

            return address.FirstOrDefault();
        }

        public override AddressInfo[] GetAddresses(QuoteAlternativeContextInfo context)
        {
            if (CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvideAddresses))
                return new AddressInfo[]
                       {
                           GetAddressInfo(context, AddressType.Invoice),
                           GetAddressInfo(context, AddressType.Delivery)
                       };
            else
                return new AddressInfo[0];
        }


        public override QuoteResponseInfo ValidateQuoteVersion(QuoteVersionContextInfo inContext, QuoteAction action)
        {
            if (CanProvideCapability(Incapabilities.CannotValidateVer))
                throw new SoIllegalOperationException("Cannot_ValidVer: Fatal Error: calculator missing. Google today announced that the next version of Android will be named \"KitKat,\" after the ubiquitous chocolate bars sold around the world. It's the first time a mainstream operating system has been given a licensed name, and the deal with trademark owner Nestle took time to complete: the BBC reports that Google director of Android global partnerships John Lagerling first called Nestle about the name in late November of 2012, and that the deal was only finalized at Mobile World Congress in Barcelona in February of this year. \"We decided within the hour to say let's do it,\" said Nestle executive vice president of marketing Patrice Bula.");

            var orgstate = inContext.CRMQuoteVersion.State;

            var result = base.ValidateQuoteVersion(inContext, action);

            if (CanProvideCapability(Incapabilities.FailValidateVer))
            {
                result.CRMQuoteVersion.State = orgstate;
                result.CRMQuoteVersion.Status = QuoteStatusInfo.Error;
                result.CRMQuoteVersion.Reason = "Fail_ValidVer: Calculator is out of battery.";
            }

            if (CanProvideCapability(Incapabilities.WarnValidateVer))
            {
                //Validate before place order can have other state than draft
                if (action != QuoteAction.PlaceOrder)
                    result.CRMQuoteVersion.State = QuoteVersionStateInfo.Draft;
                result.CRMQuoteVersion.Status = QuoteStatusInfo.Warning;
                result.CRMQuoteVersion.Reason = "Warn_ValidVer: Calculator says 8008135.";
            }
            if (CanProvideCapability(Incapabilities.NeedValidateVer))
            {
                result.CRMQuoteVersion.State = QuoteVersionStateInfo.DraftNeedsApproval;
            }
            return result;
        }


        public override QuoteVersionResponseInfo UpdateQuoteVersionPrices(QuoteVersionContextInfo inContext, HashSet<string> writeableFields)
        {
            if (CanProvideCapability(Incapabilities.CannotUpdate))
                throw new SoIllegalOperationException("Cannot_Update: Fatal Error: calculator missing.");

            var result = base.UpdateQuoteVersionPrices(inContext, writeableFields);

            if (CanProvideCapability(Incapabilities.FailUpdate))
            {
                result.CRMQuoteVersion.State = QuoteVersionStateInfo.Draft;
                result.CRMQuoteVersion.Status = QuoteStatusInfo.Error;
                result.CRMQuoteVersion.Reason = "Fail_Update: Calculator is out of battery.";
            }

            if (CanProvideCapability(Incapabilities.WarnUpdate))
            {
                result.CRMQuoteVersion.State = QuoteVersionStateInfo.Draft;
                result.CRMQuoteVersion.Status = QuoteStatusInfo.Warning;
                result.CRMQuoteVersion.Reason = "Warn_Update: Calculator says 8008135.";
            }
            foreach (var alt in result.CRMAlternativesWithLines)
                foreach (var ql in alt.CRMQuoteLines)
                {
                    ql.UnitListPrice = ql.UnitListPrice * 1.1;
                    ql.ERPDiscountPercent = 10;
                    ql.UnitCost = ql.UnitCost * 0.89;
                    QuoteCalculation.CalculateValues(ql);
                }
            result.CRMQuoteVersion.LastRecalculated = DateTime.UtcNow;
            return result;
        }



        private void ReadInData()
        {
            if (Filename == null)
                return;

            var package = LoadExcelFile();
            var book = package.Workbook;

            ReadCapabilities(book);
            ReadAddresses(book);

            ReadPricelists(book);
            ReadProducts(book);

            ReadPaymentTerms(book);
            ReadPaymentTypes(book);
            ReadDeliveryTerms(book);
            ReadDeliveryTypes(book);

            ReadProductCategories(book);
            ReadProductFamilies(book);
            ReadProductTypes(book);
        }

        private ExcelPackage LoadExcelFile()
        {
            if (!Filename.EndsWith(".xlsx") && !Filename.EndsWith(".xls"))
                throw new ArgumentException("Filename does not end with .xlsx or .xls extension", "Filename");

            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var excelPackage = new ExcelPackage();

            using (var fileStream = File.OpenRead(Filename))
            {
                excelPackage.Load(fileStream);
            }

            return excelPackage;
        }

        public Dictionary<string, bool> Capabilities { get; set; }
        private ExcelWorksheet ReadCapabilities(ExcelWorkbook book)
        {
            var sheet = book.Worksheets[0];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                Capabilities = new Dictionary<string, bool>(rows);

                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new ListItemInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    var name = sheet.ReadStringValue(rownum, colNum++);
                    var value = sheet.ReadBoolValue(rownum, colNum++);

                    Capabilities.Add(name, value);
                }
            }
            return sheet;
        }

        internal enum AddressType
        {
            Invoice = 0,
            Delivery = 1
        }
        internal class AddressAddress
        {
            internal string ContactKey { get; set; }
            internal AddressType Type { get; set; }
        }
        internal Dictionary<AddressAddress, AddressInfo> Addresses { get; set; }
        private ExcelWorksheet ReadAddresses(ExcelWorkbook book)
        {
            var sheet = book.Worksheets[1];
            if (sheet != null && CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvideAddresses))
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                Addresses = new Dictionary<AddressAddress, AddressInfo>(rows);

                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var t = new AddressAddress();
                    var a = new AddressInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    t.ContactKey = sheet.ReadStringValue(rownum, colNum++);
                    t.Type = (sheet.ReadStringValue(rownum, colNum++) == "Invoice") ? AddressType.Invoice : AddressType.Delivery;

                    a.AddressLine1 = sheet.ReadStringValue(rownum, colNum++);
                    a.AddressLine2 = sheet.ReadStringValue(rownum, colNum++);
                    a.AddressLine3 = sheet.ReadStringValue(rownum, colNum++);
                    a.City = sheet.ReadStringValue(rownum, colNum++);
                    a.ZipCode = sheet.ReadStringValue(rownum, colNum++);
                    a.CountryCode = sheet.ReadStringValue(rownum, colNum++);
                    a.Country = sheet.ReadStringValue(rownum, colNum++);

                    Addresses.Add(t, a);
                }
            }
            return sheet;
        }


        private ExcelWorksheet ReadPricelists(ExcelWorkbook book)
        {
            // Load products:
            var sheet = book.Worksheets[2];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                ProductProvider.PriceLists = new List<PriceListInfo>(rows);

                var counter = 0;
                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new PriceListInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    p.ERPPriceListKey = sheet.ReadStringValue(rownum, colNum++);
                    p.Name = sheet.ReadStringValue(rownum, colNum++);
                    p.Description = sheet.ReadStringValue(rownum, colNum++);
                    //p.CurrencyName = sheet.ReadStringValue(rownum, colNum++);
                    p.Currency = sheet.ReadStringValue(rownum, colNum++);
                    p.ValidFrom = sheet.ReadDateTimeValue(rownum, colNum++);
                    p.ValidTo = sheet.ReadDateTimeValue(rownum, colNum++);
                    p.IsActive = sheet.ReadBoolValue(rownum, colNum++);

                    ProductProvider.PriceLists.Add(p);

                    counter++;
                }
            }
            return sheet;
        }

        private ExcelWorksheet ReadProducts(ExcelWorkbook book)
        {
            // Load products:
            var sheet = book.Worksheets[3];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                ProductProvider.Products = new List<ProductInfo>(rows);

                ProductProvider.Images = new Dictionary<string, string[]>(rows);

                var counter = 0;
                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new ProductInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    p.ERPPriceListKey = sheet.ReadStringValue(rownum, colNum++);
                    p.InAssortment = sheet.ReadBoolValue(rownum, colNum++);
                    p.InStock = sheet.ReadDoubleValue(rownum, colNum++);
                    p.ERPProductKey = sheet.ReadStringValue(rownum, colNum++);
                    p.Name = sheet.ReadStringValue(rownum, colNum++);
                    p.Description = sheet.ReadStringValue(rownum, colNum++);
                    p.Code = sheet.ReadStringValue(rownum, colNum++);
                    p.QuantityUnit = sheet.ReadStringValue(rownum, colNum++);
                    p.PriceUnit = sheet.ReadStringValue(rownum, colNum++);
                    p.ItemNumber = sheet.ReadStringValue(rownum, colNum++);
                    p.Url = sheet.ReadStringValue(rownum, colNum++);
                    p.ProductCategoryKey = sheet.ReadStringValue(rownum, colNum++);
                    p.ProductFamilyKey = sheet.ReadStringValue(rownum, colNum++);
                    p.ProductTypeKey = sheet.ReadStringValue(rownum, colNum++);

                    p.Rights = sheet.ReadStringValue(rownum, colNum++);
                    p.Rule = sheet.ReadStringValue(rownum, colNum++);

                    p.SupplierCode = sheet.ReadStringValue(rownum, colNum++);
                    p.Supplier = sheet.ReadStringValue(rownum, colNum++);
                    p.VATInfo = sheet.ReadStringValue(rownum, colNum++);
                    p.VAT = sheet.ReadDoubleValue(rownum, colNum++);
                    if (CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvideCost))
                        p.UnitCost = sheet.ReadDoubleValue(rownum, colNum++);
                    else
                        colNum++;
                    if (CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvideMinimumPrice))
                        p.UnitMinimumPrice = sheet.ReadDoubleValue(rownum, colNum++);
                    else
                        colNum++;
                    p.UnitListPrice = sheet.ReadDoubleValue(rownum, colNum++);
                    p.Thumbnail = sheet.ReadStringValue(rownum, colNum++);

                    var extraInfo = sheet.ReadStringValue(rownum, colNum++);
                    if (CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvideExtraData))
                    {
                        p.ExtraInfo = SuperOffice.CRM.Sale.QuoteExtensions.GetProductExtraDataFieldInfo(extraInfo);
                        p.RawExtraInfo = extraInfo;
                    }

                    p.ExtraField1 = sheet.ReadStringValue(rownum, colNum++);
                    p.ExtraField2 = sheet.ReadStringValue(rownum, colNum++);
                    p.ExtraField3 = sheet.ReadStringValue(rownum, colNum++);
                    p.ExtraField4 = sheet.ReadStringValue(rownum, colNum++);
                    p.ExtraField5 = sheet.ReadStringValue(rownum, colNum++);

                    ProductProvider.Products.Add(p);

                    var img1 = sheet.ReadStringValue(rownum, colNum++);
                    var img2 = sheet.ReadStringValue(rownum, colNum++);
                    // Read in two images:
                    if (CanProvideCapability(CRMQuoteConnectorCapabilities.CanProvidePicture))
                        ProductProvider.Images[p.ERPProductKey] = new string[] { img1, img2 };



                    //Products[counter].ProductImage1 = sheet.readStr(rownum, colNum++);
                    //Products[counter].ProductImage2 = sheet.readStr(rownum, colNum++);

                    counter++;
                }
            }
            return sheet;
        }


        public List<ListItemInfo> PaymentTerms { get; set; }
        private ExcelWorksheet ReadPaymentTerms(ExcelWorkbook book)
        {
            var sheet = book.Worksheets[6];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                PaymentTerms = new List<ListItemInfo>(rows);

                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new ListItemInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    p.ERPQuoteListItemKey = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayDescription = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayValue = sheet.ReadStringValue(rownum, colNum++);

                    PaymentTerms.Add(p);
                }
            }
            return sheet;
        }

        public List<ListItemInfo> PaymentTypes { get; set; }
        private ExcelWorksheet ReadPaymentTypes(ExcelWorkbook book)
        {
            var sheet = book.Worksheets[7];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                PaymentTypes = new List<ListItemInfo>(rows);

                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new ListItemInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    p.ERPQuoteListItemKey = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayDescription = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayValue = sheet.ReadStringValue(rownum, colNum++);

                    PaymentTypes.Add(p);
                }
            }
            return sheet;
        }

        public List<ListItemInfo> DeliveryTerms { get; set; }
        private ExcelWorksheet ReadDeliveryTerms(ExcelWorkbook book)
        {
            var sheet = book.Worksheets[8];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                DeliveryTerms = new List<ListItemInfo>(rows);

                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new ListItemInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    p.ERPQuoteListItemKey = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayDescription = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayValue = sheet.ReadStringValue(rownum, colNum++);

                    DeliveryTerms.Add(p);
                }
            }
            return sheet;
        }

        public List<ListItemInfo> DeliveryTypes { get; set; }
        private ExcelWorksheet ReadDeliveryTypes(ExcelWorkbook book)
        {
            var sheet = book.Worksheets[9];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                DeliveryTypes = new List<ListItemInfo>(rows);

                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new ListItemInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    p.ERPQuoteListItemKey = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayDescription = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayValue = sheet.ReadStringValue(rownum, colNum++);

                    DeliveryTypes.Add(p);
                }
            }
            return sheet;
        }


        public List<ListItemInfo> ProductCategories { get; set; }
        private ExcelWorksheet ReadProductCategories(ExcelWorkbook book)
        {
            var sheet = book.Worksheets[10];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                ProductCategories = new List<ListItemInfo>(rows);

                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new ListItemInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    p.ERPQuoteListItemKey = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayDescription = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayValue = sheet.ReadStringValue(rownum, colNum++);

                    ProductCategories.Add(p);
                }
            }
            return sheet;
        }

        public List<ListItemInfo> ProductFamilies { get; set; }
        private ExcelWorksheet ReadProductFamilies(ExcelWorkbook book)
        {
            var sheet = book.Worksheets[11];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                ProductFamilies = new List<ListItemInfo>(rows);

                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new ListItemInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    p.ERPQuoteListItemKey = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayDescription = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayValue = sheet.ReadStringValue(rownum, colNum++);

                    ProductFamilies.Add(p);
                }
            }
            return sheet;
        }

        public List<ListItemInfo> ProductTypes { get; set; }
        private ExcelWorksheet ReadProductTypes(ExcelWorkbook book)
        {
            var sheet = book.Worksheets[12];
            if (sheet != null)
            {
                var rows = sheet.Dimension.End.Row - sheet.Dimension.Start.Row;
                ProductTypes = new List<ListItemInfo>(rows);

                //+ 1 = Skip the first row
                for (var rownum = sheet.Dimension.Start.Row + 1; rownum <= sheet.Dimension.End.Row; rownum++)
                {
                    var p = new ListItemInfo();

                    var colNum = sheet.Dimension.Start.Column;
                    p.ERPQuoteListItemKey = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayDescription = sheet.ReadStringValue(rownum, colNum++);
                    p.DisplayValue = sheet.ReadStringValue(rownum, colNum++);

                    ProductTypes.Add(p);
                }
            }
            return sheet;
        }

        readonly TSRandom rnd = new TSRandom();

        private string GenKey()
        {
            var x = rnd.Next();
            var fmt = x.ToString("X") + "000000";
            return fmt.Substring(0, 6);
        }

        private string AddQuote(CRM.QuoteVersionContextInfo quote)
        {
            var package = LoadExcelFile();
            var book = package.Workbook;
            var sheet = book.Worksheets[4];
            if (sheet != null)
            {
                if (sheet.Dimension == null ||  sheet.Dimension.End.Row == 0)
                {
                    sheet.Cells[1, 1].Value = "ContactId";
                    sheet.Cells[1, 2].Value = "ContactName";
                    sheet.Cells[1, 2].Value = "ContactErpKey";
                    sheet.Cells[1, 3].Value = "EISConnectionId";
                    sheet.Cells[1, 4].Value = "QuoteErpKey";
                    sheet.Cells[1, 5].Value = "OrderErpKey";
                    sheet.Cells[1, 6].Value = "InvoiceAddr1";
                    sheet.Cells[1, 7].Value = "InvoiceCity";
                    sheet.Cells[1, 8].Value = "DeliveryAddr1";
                    sheet.Cells[1, 9].Value = "DeliveryCity";
                    sheet.Cells[1, 10].Value = "VersionId";
                    sheet.Cells[1, 11].Value = "VerNumber";
                    sheet.Cells[1, 12].Value = "VerRank";
                }
                var row = sheet.Dimension.End.Row; // + 1;

                if (string.IsNullOrEmpty(quote.CRMQuote.ERPQuoteKey))
                {
                    quote.CRMQuote.ERPQuoteKey = GenKey();
                    quote.CRMQuote.ERPQuoteKey = quote.CRMQuote.ERPQuoteKey.ToLowerInvariant();
                }

                var colNum = sheet.Dimension.Start.Column;

                sheet.Cells[row, colNum++].Value = quote.CRMCompany.ContactId;
                sheet.Cells[row, colNum++].Value = quote.CRMCompany.Name;
                sheet.Cells[row, colNum++].Value = quote.ERPCompanyKey;
                sheet.Cells[row, colNum++].Value = quote.EISConnectionId;
                sheet.Cells[row, colNum++].Value = quote.CRMQuote.ERPQuoteKey;
                sheet.Cells[row, colNum++].Value = quote.CRMQuote.ERPOrderKey;
                if (quote.InvoiceAddress != null)
                {
                    sheet.Cells[row, colNum++].Value = quote.InvoiceAddress.AddressLine1;
                    sheet.Cells[row, colNum++].Value = quote.InvoiceAddress.City;
                }
                else
                {
                    sheet.Cells[row, colNum++].Value = quote.CRMCompany.PostalAddressLine1;
                    sheet.Cells[row, colNum++].Value = quote.CRMCompany.PostalAddressCity;
                }
                if (quote.DeliveryAddress != null)
                {
                    sheet.Cells[row, colNum++].Value = quote.DeliveryAddress.AddressLine1;
                    sheet.Cells[row, colNum++].Value = quote.DeliveryAddress.City;
                }
                else
                {
                    sheet.Cells[row, colNum++].Value = quote.CRMCompany.StreetAddressLine1;
                    sheet.Cells[row, colNum++].Value = quote.CRMCompany.StreetAddressCity;
                }

                sheet.Cells[row, colNum++].Value = quote.CRMQuoteVersion.QuoteVersionId;
                sheet.Cells[row, colNum++].Value = quote.CRMQuoteVersion.Number;
                sheet.Cells[row, colNum++].Value = quote.CRMQuoteVersion.Rank;

                row++;

                //+ 1 = Skip the first row
                foreach (var alt in quote.CRMAlternativesWithLines)
                {
                    sheet.Cells[row, 1].Value = "AltName=";
                    sheet.Cells[row, 2].Value = alt.CRMAlternative.Name;
                    sheet.Cells[row, 3].Value = "AltErpKey=";
                    sheet.Cells[row, 4].Value = alt.CRMAlternative.ERPQuoteAlternativeKey;
                    sheet.Cells[row, 5].Value = "AltTotal=";
                    sheet.Cells[row, 6].Value = alt.CRMAlternative.TotalPrice;

                    row++;
                    sheet.Cells[row, 3].Value = "Code";
                    sheet.Cells[row, 4].Value = "Name";
                    sheet.Cells[row, 5].Value = "LineErpKey";
                    sheet.Cells[row, 6].Value = "ProdErpKey";
                    sheet.Cells[row, 7].Value = "Quant";
                    sheet.Cells[row, 8].Value = "Listprice";
                    sheet.Cells[row, 9].Value = "DiscAmt";
                    sheet.Cells[row, 10].Value = "Total";
                    row++;

                    foreach (var line in alt.CRMQuoteLines)
                    {
                        sheet.Cells[row, 3].Value = line.Code;
                        sheet.Cells[row, 4].Value = line.Name;
                        sheet.Cells[row, 5].Value = line.ERPQuoteLineKey;
                        sheet.Cells[row, 6].Value = line.ERPProductKey;
                        sheet.Cells[row, 7].Value = line.Quantity;
                        sheet.Cells[row, 8].Value = line.UnitListPrice;
                        sheet.Cells[row, 9].Value = line.DiscountAmount;
                        sheet.Cells[row, 10].Value = line.TotalPrice;
                        row++;
                    }
                }
                sheet.Cells[row, 1].Value = "--end of quote--";
            }
            package.Save(Filename);
            return quote.CRMQuote.ERPQuoteKey;
        }


        private string AddOrder(QuoteAlternativeContextInfo order)
        {
            var package = LoadExcelFile();
            var book = package.Workbook;
            var sheet = book.Worksheets[5];
            if (sheet != null)
            {
                if (sheet.Dimension == null || sheet.Dimension.End.Row == 0)
                {
                    sheet.Cells[1, 1].Value = "ContactId";
                    sheet.Cells[1, 2].Value = "ContactName";
                    sheet.Cells[1, 3].Value = "ContactErpKey";
                    sheet.Cells[1, 4].Value = "EISConnectionId";
                    sheet.Cells[1, 5].Value = "QuoteErpKey";
                    sheet.Cells[1, 6].Value = "OrderErpKey";
                    sheet.Cells[1, 7].Value = "InvoiceAddr1";
                    sheet.Cells[1, 8].Value = "InvoiceCity";
                    sheet.Cells[1, 9].Value = "DeliveryAddr1";
                    sheet.Cells[1, 10].Value = "DeliveryCity";
                    sheet.Cells[1, 11].Value = "PO Num";
                    sheet.Cells[1, 12].Value = "Comment";
                    sheet.Cells[1, 13].Value = "VersionId";
                    sheet.Cells[1, 14].Value = "VerNumber";
                    sheet.Cells[1, 15].Value = "VerRank";
                    sheet.Cells[1, 16].Value = "AltName";
                    sheet.Cells[1, 17].Value = "AltERPKey";
                    sheet.Cells[1, 18].Value = "AltDiscountAmt";
                    sheet.Cells[1, 19].Value = "AltTotalPrice";
                }
                var row = sheet.Dimension.End.Row;

                order.CRMQuote.ERPOrderKey = GenKey();

                sheet.Cells[row, 1].Value = order.CRMCompany.ContactId;
                sheet.Cells[row, 2].Value = order.CRMCompany.Name;
                sheet.Cells[row, 3].Value = order.ERPCompanyKey;
                sheet.Cells[row, 4].Value = order.EISConnectionId;
                sheet.Cells[row, 5].Value = order.CRMQuote.ERPQuoteKey;
                sheet.Cells[row, 6].Value = order.CRMQuote.ERPOrderKey;
                if (order.InvoiceAddress != null)
                {
                    sheet.Cells[row, 7].Value = order.InvoiceAddress.AddressLine1;
                    sheet.Cells[row, 8].Value = order.InvoiceAddress.City;
                }
                else
                {
                    sheet.Cells[row, 7].Value = order.CRMCompany.PostalAddressLine1;
                    sheet.Cells[row, 8].Value = order.CRMCompany.PostalAddressCity;
                }
                if (order.DeliveryAddress != null)
                {
                    sheet.Cells[row, 9].Value = order.DeliveryAddress.AddressLine1;
                    sheet.Cells[row, 10].Value = order.DeliveryAddress.City;
                }
                else
                {
                    sheet.Cells[row, 9].Value = order.CRMCompany.StreetAddressLine1;
                    sheet.Cells[row, 10].Value = order.CRMCompany.StreetAddressCity;
                }

                sheet.Cells[row, 11].Value = order.CRMQuote.PoNumber;
                sheet.Cells[row, 12].Value = order.CRMQuote.OrderComment;

                sheet.Cells[row, 13].Value = order.CRMQuoteVersion.QuoteVersionId;
                sheet.Cells[row, 14].Value = order.CRMQuoteVersion.Number;
                sheet.Cells[row, 15].Value = order.CRMQuoteVersion.Rank;

                sheet.Cells[row, 16].Value = order.CRMAlternativeWithLines.CRMAlternative.Name;
                sheet.Cells[row, 17].Value = order.CRMAlternativeWithLines.CRMAlternative.ERPQuoteAlternativeKey;
                sheet.Cells[row, 18].Value = order.CRMAlternativeWithLines.CRMAlternative.DiscountAmount;
                sheet.Cells[row, 19].Value = order.CRMAlternativeWithLines.CRMAlternative.TotalPrice;

                row++;
                sheet.Cells[row, 1].Value = "Code";
                sheet.Cells[row, 2].Value = "Name";
                sheet.Cells[row, 3].Value = "LineErpKey";
                sheet.Cells[row, 4].Value = "ProdErpKey";
                sheet.Cells[row, 5].Value = "Quant";
                sheet.Cells[row, 6].Value = "Listpri";
                sheet.Cells[row, 7].Value = "DiscAmt";
                sheet.Cells[row, 8].Value = "Total";

                row++;
                foreach (var line in order.CRMAlternativeWithLines.CRMQuoteLines)
                {
                    sheet.Cells[row, 1].Value = line.Code;
                    sheet.Cells[row, 2].Value = line.Name;
                    sheet.Cells[row, 3].Value = line.ERPQuoteLineKey;
                    sheet.Cells[row, 4].Value = line.ERPProductKey;
                    sheet.Cells[row, 5].Value = line.Quantity;
                    sheet.Cells[row, 6].Value = line.UnitListPrice;
                    sheet.Cells[row, 7].Value = line.DiscountAmount;
                    sheet.Cells[row, 8].Value = line.TotalPrice;
                    row++;
                }
                sheet.Cells[row, 1].Value = "--end of order--";
            }
            package.Save(Filename);
            return order.CRMQuote.ERPOrderKey;
        }


        public override QuoteLineInfo ValidateQuoteLine(QuoteAlternativeContextInfo context, QuoteLineInfo ql, bool clearOldValues = false)
        {
            var result = base.ValidateQuoteLine(context, ql, clearOldValues);

            ReadInData();

            if (CanProvideCapability(Incapabilities.CannotValidateLine))
                throw new SoException("Cannot_Validate_Line: system needs to be rebooted");

            if (CanProvideCapability(Incapabilities.FailValidateLine))
            {
                result.AddStatus(QuoteStatusInfo.Error, "Fail_Validate_Line: permission denied");
            }
            if (CanProvideCapability(Incapabilities.WarnValidateLine))
            {
                result.AddStatus(QuoteStatusInfo.Warning, "Warn_Validate_Line: The Token fell out of the ring. Call us when you find it");
            }
            return result;
        }


        public override QuoteAlternativeWithLinesInfo ValidateAlternative(QuoteAlternativeContextInfo context, bool clearOldValues)
        {
            var result = base.ValidateAlternative(context, clearOldValues);

            ReadInData();

            if (CanProvideCapability(Incapabilities.CannotValidateAlt))
                throw new SoException("Cannot_Validate_Alt: Write-only-memory subsystem too slow for this machine. Contact your local dealer.");

            if (CanProvideCapability(Incapabilities.FailValidateAlt))
            {
                result.CRMAlternative.AddStatus(QuoteStatusInfo.Error, "Fail_Validate_Alt: Someone hooked the twisted pair wires into the answering machine.");
            }
            if (CanProvideCapability(Incapabilities.WarnValidateAlt))
            {
                result.CRMAlternative.AddStatus(QuoteStatusInfo.Warning, "Warn_Validate_Alt: astropneumatic oscillations in the water-cooling");
            }
            return result;
        }


        public override PlaceOrderResponseInfo PlaceOrder(QuoteAlternativeContextInfo context)
        {
            ReadInData();

            var result = new PlaceOrderResponseInfo(context);
            result.IsOk = true;
            result.State = ResponseState.Ok;

            if (CanProvideCapability(Incapabilities.PlaceOrderUrl))
                result.Url = "http://qa-build/echo.asp?contact={0}&quotever={1}&alt={2}&amount={3}".FormatWith(context.CRMCompany.Name, context.CRMQuoteVersion.Rank, context.CRMAlternativeWithLines.CRMAlternative.Name, context.CRMSale.Amount);
            if (CanProvideCapability(Incapabilities.PlaceOrderSOProto))
                result.Url = "superoffice:project.main";

            if (CanProvideCapability(Incapabilities.CannotPlaceOrder))
                throw new SoIllegalOperationException("Cannot_Place_Order: Fatal Error: Borg nanites have infested the server.");

            if (CanProvideCapability(Incapabilities.FailPlaceOrder))
            {
                result.IsOk = false;
                result.State = ResponseState.Error;
                result.UserExplanation = "Fail_Place_Order: That function is not currently supported, but Bill Gates assures us it will be featured in the next upgrade.";
                result.TechExplanation = "wrong polarity of neutron flow";
            }

            if (result.IsOk)
                result.CRMQuote.ERPOrderKey = AddOrder(context);


            return result;
        }

        public override OrderResponseInfo GetOrderState(QuoteAlternativeContextInfo context)
        {
            if (CanProvideCapability(Incapabilities.CannotOrderState))
                throw new SoException("Cannot_Order_State: Too few computrons available.");
            var result = new OrderResponseInfo(context);
            if (CanProvideCapability(Incapabilities.FailOrderState))
            {
                result.State = ResponseState.Error;
                result.UserExplanation = "Fail_Order_State: knot in cables caused data stream to become twisted and kinked";
                result.TechExplanation = "wrong polarity of neutron flow.";
            }
            if (CanProvideCapability(Incapabilities.WarnOrderState))
            {
                result.State = ResponseState.Warning;
                result.UserExplanation = "Warn_Order_State: It's not plugged in.";
                result.TechExplanation = "wrong polarity of neutron flow";
            }
            return result;
        }

    }
}
