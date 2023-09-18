using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FCC.AF.CargaSociedades.Shared.Config;
using FCC.AF.CargaSociedades.Shared.Settings;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using DataTable = System.Data.DataTable;

namespace FCC.AF.CargaSociedades.Application
{
    public class CargaSociedades
    {
        public static string k_folder_type = "Folder";
        public static string k_subfolder_type = "Subfolder";

        public static string k_field_Title = "Title";
        public static string k_field_Display = "Display";
        public static string k_field_TipoDato = "DataType";
        public static string k_field_Tipo = "ItemType";
        public static string k_field_Statutory = "Statutory";
        public static string k_field_StatutoryExcelColumn = "StatutoryExcelColumn";
        public static string k_field_Reporting = "Reporting";
        public static string k_field_ReportingExcelColumn = "ReportingExcelColumn";
        public static string k_field_Certificates = "Certificates";
        public static string k_field_CertificatesExcelColumn = "CertificatesExcelColumn";

        public static string k_list_sitecolumnsettings = "SiteColumn Settings";
        public static string k_list_statutory = "Statutory Audits";
        public static string k_list_reporting = "Reporting Audits";
        public static string k_list_certificates = "Certificates Audits";

        public static string k_sitecolumn_group = "FCC";

        public static string k_contenttype_group_statutory = "FCC Statutory";
        public static string k_contenttype_group_reporting = "FCC Reporting";
        public static string k_contenttype_group_certificates = "FCC Certificates";

        public static int k_statutory_groupindex = 7;
        public static int k_reporting_groupindex = 6;
        public static int k_certificates_groupindex = 6;

        public static async void DoProcess(X509Certificate2 certificate, Microsoft.Extensions.Logging.ILogger log)
        {
            try
            {
                log.LogWarning("Cargamos datos de configuración");
                AZConfig cfg = new AZConfig
                {
                    _TenantId = Environment.GetEnvironmentVariable("_TenantId", EnvironmentVariableTarget.Process),
                    _ClientId = Environment.GetEnvironmentVariable("_ClientId", EnvironmentVariableTarget.Process),
                    _ClientSecret = Environment.GetEnvironmentVariable("_ClientSecret", EnvironmentVariableTarget.Process),
                    _HubUrl = Environment.GetEnvironmentVariable("_HubUrl", EnvironmentVariableTarget.Process),
                    _TenantUrl = Environment.GetEnvironmentVariable("_TenantUrl", EnvironmentVariableTarget.Process),
                    _TenantAdminUrl = Environment.GetEnvironmentVariable("_TenantAdminUrl", EnvironmentVariableTarget.Process),
                    _SitePattern = Environment.GetEnvironmentVariable("_SitePattern", EnvironmentVariableTarget.Process),
                    _MembersGroup = Environment.GetEnvironmentVariable("_MembersGroup", EnvironmentVariableTarget.Process)

                };
                                
                DataTable StatutoryDT = null;
                DataTable ReportingDT = null;
                DataTable CertificatesDT = null;
                DataTable DocumentTypesDT = null;
                DataTable SecurityGroupsDT = null;

                log.LogWarning("Conectamos a la biblioteca de documentos del hub");
                using (ClientContext ctx = CreateContext(cfg, certificate, cfg._HubUrl))
                {
                    log.LogWarning("Contexto creado");

                    List list = ctx.Web.Lists.GetByTitle("Documents");
                    ctx.Load(list);
                    ctx.Load(list.Fields);
                    ctx.ExecuteQuery();

                    CamlQuery camlQuery = new CamlQuery();
                    ListItemCollection collListItem = list.GetItems(camlQuery);
                    ctx.Load(collListItem);
                    ctx.ExecuteQuery();

                    log.LogWarning("Lista cargada correctamente");
                    log.LogWarning("Nº Documentos: " + list.ItemCount);
                    
                    foreach (Microsoft.SharePoint.Client.ListItem item in collListItem)
                    {
                        try
                        {
                            if (Convert.ToBoolean(item["Processed"]) == false)
                            {
                                ctx.Load(item.File);
                                ctx.ExecuteQuery();
                                log.LogWarning("Abrimos fichero excel de sociedades.");
                                log.LogWarning("[" + item.File.ServerRelativeUrl + "]");
                                                            
                                Microsoft.SharePoint.Client.File file = ctx.Web.GetFileByUrl(string.Format("{0}{1}", cfg._TenantUrl, item.File.ServerRelativeUrl));
                                Microsoft.SharePoint.Client.ClientResult<Stream> mstream = file.OpenBinaryStream();
                                ctx.Load(file);
                                ctx.ExecuteQuery();

                                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(mstream.Value, false))
                                {
                                    WorkbookPart workbookPart = doc.WorkbookPart;
                                    IEnumerable<Sheet> sheets = doc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                                    StatutoryDT = GetDataFromExcel(ctx, doc, sheets.ElementAt(0).Id.Value, "Statutory Audits");
                                    ReportingDT = GetDataFromExcel(ctx, doc, sheets.ElementAt(1).Id.Value, "Reporting Audits");
                                    CertificatesDT = GetDataFromExcel(ctx, doc, sheets.ElementAt(2).Id.Value, "Certificates Audits");
                                    DocumentTypesDT = GetDataFromExcel(ctx, doc, sheets.ElementAt(3).Id.Value, "Document Types");
                                    SecurityGroupsDT = GetDataFromExcel(ctx, doc, sheets.ElementAt(4).Id.Value, "Security Groups");
                                }

                                string siteurl = String.Format("{0}{1}", cfg._TenantUrl, cfg._SitePattern.Replace("XXXX", item["Year"].ToString()));

                                using (ClientContext ctx2 = CreateContext(cfg, certificate, siteurl))
                                {
                                    log.LogWarning("Contexto creado para el año");

                                    CheckContentTypesStatutory(ctx2, DocumentTypesDT, "FCC Statutory", log);
                                    CheckContentTypesReporting(ctx2, DocumentTypesDT, "FCC Reporting", log);
                                    CheckContentTypesCertificates(ctx2, DocumentTypesDT, "FCC Certificates", log);
                                    CheckSiteColumns(ctx2, log);
                                    CheckSiteGroups(ctx2, SecurityGroupsDT, log);

                                    //**************************************************
                                    LoadDataStatutory(cfg, ctx2, StatutoryDT, DocumentTypesDT, SecurityGroupsDT, log);
                                    LoadDataReporting(cfg, ctx2, ReportingDT, DocumentTypesDT, SecurityGroupsDT, log);
                                    LoadDataCertificates(cfg, ctx2, CertificatesDT, DocumentTypesDT, SecurityGroupsDT, log);
                                    //**************************************************


                                } // END using SITE AÑO

                                // Actualizamos el estado del documento de sociedades
                                item["Processed"] = true;
                                item["Error"] = String.Empty;
                                item.Update();
                                ctx.ExecuteQuery();

                            } // END if NOT Processed
                        }
                        catch (Exception error)
                        {
                            log.LogError(error, error.Message);
                            item["Error"] = error.Message;
                            item.Update();
                            ctx.ExecuteQuery();
                        }

                    } // END ForEach item

                } // END using HUB  

            }
            catch (Exception error)
            {
                log.LogError(error, error.Message);
            }

        }

        private static string AddFolder(ClientContext ctx, List list, DataRow row, List<SiteColumnSettings> statutorycols, Microsoft.Extensions.Logging.ILogger log)
        {
            string result = String.Empty;

            for (int i = 0; i < statutorycols.Count; i++)
            {
                if (statutorycols[i].Tipo == k_folder_type)
                {
                    log.LogWarning(statutorycols[i].Display + " [" + row[i].ToString() + "]");
                    result = CreateFolder(ctx, row[i].ToString(), list, log);
                    break;
                }
            }

            return result;
        }

        private static string CreateFolder(ClientContext ctx, string title, List list, Microsoft.Extensions.Logging.ILogger log)
        {
            string result = String.Empty;

            try
            {
                if (!string.IsNullOrEmpty(title))
                {
                    ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                    newItemInfo.UnderlyingObjectType = Microsoft.SharePoint.Client.FileSystemObjectType.Folder;
                    newItemInfo.LeafName = title.Trim();
                    Microsoft.SharePoint.Client.ListItem newListItem = list.AddItem(newItemInfo);
                    newListItem[k_field_Title] = title.Trim();
                    newListItem.Update();
                    ctx.ExecuteQuery();
                }
            }
            catch(Exception ex) {
                log.LogError(ex.Message);
            }
            finally
            {
                result = title.Trim();
            }

            return result;
        }

        private static string AddSubfolder(ClientContext ctx, List list, string foldername, DataRow row, List<SiteColumnSettings> statutorycols, Microsoft.Extensions.Logging.ILogger log)
        {
            string result = String.Empty;

            for (int i = 0; i < statutorycols.Count; i++)
            {
                if (statutorycols[i].Tipo == k_subfolder_type)
                {
                    log.LogWarning(statutorycols[i].Display + " [" + foldername + "]" + " [" + row[i].ToString() + "]");
                    result = CreateSubfolder(ctx, statutorycols, list, foldername, row[i].ToString(), log);
                    break;
                }
            }

            return result;
        }

        private static string CreateSubfolder(ClientContext ctx, List<SiteColumnSettings> statutorycols, List list, string foldername, string title, Microsoft.Extensions.Logging.ILogger log)
        {
            string result = String.Empty;

            try
            {
                if (!string.IsNullOrEmpty(title)) {
                    ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                    newItemInfo.UnderlyingObjectType = Microsoft.SharePoint.Client.FileSystemObjectType.Folder;
                    newItemInfo.LeafName = foldername + "/" + title.Trim();
                    Microsoft.SharePoint.Client.ListItem newListItem = list.AddItem(newItemInfo);
                    newListItem[k_field_Title] = title.Trim();
                    newListItem.Update();
                    ctx.ExecuteQuery();
                }
            }
            catch(Exception ex) {
                log.LogError(ex.Message);    
            }
            finally
            {
                result = title.Trim();
            }

            return result;
        }

        private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            string value = string.Empty;
            try
            {
                if (cell != null)
                {
                    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (cell.CellValue != null)
                    {
                        value = cell.CellValue.InnerXml;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                            {
                                isError = false;
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                            }
                        }
                        else
                        {
                            isError = false;
                            return value;
                        }
                    }
                }
                isError = false;
                return string.Empty;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                }
            }
            return value;
        }

        private static DataTable GetDataFromExcel(ClientContext ctx, SpreadsheetDocument doc, string SheetId, string datatablename)
        {
            DataTable result = new DataTable(datatablename);

            WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(SheetId);
            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();

            foreach (Cell cell in rows.ElementAt(0))
            {
                string str = GetCellValue(ctx, doc, cell);
                result.Columns.Add(str);
            }
            foreach (Row row in rows)
            {
                if (row != null)
                {
                    DataRow dataRow = result.NewRow();
                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        dataRow[i] = GetCellValue(ctx, doc, row.Descendants<Cell>().ElementAt(i));
                    }
                    result.Rows.Add(dataRow);
                }
            }
            result.Rows.RemoveAt(0);

            return result;
        }

        public static ClientContext CreateContext(AZConfig cfg, X509Certificate2 certificate, string url)
        {

            ClientContext context = null;
            using (PnP.Framework.AuthenticationManager am = new PnP.Framework.AuthenticationManager(cfg._ClientId, certificate, cfg._TenantId))
            {
                context = am.GetContext(url);
            }

            context.RequestTimeout = 1000 * 60 * 15;
            return context;
        }

        private static void LoadSiteColumn(ClientContext ctx, Microsoft.SharePoint.Client.Field newfield, string groupname, Microsoft.Extensions.Logging.ILogger log)
        {
            ContentTypeCollection ctypes = ctx.Web.ContentTypes;
            ctx.Load(ctypes);
            ctx.ExecuteQuery();

            foreach (ContentType ctype in ctypes)
            {
                if (ctype.Group == groupname)
                {
                    Microsoft.SharePoint.Client.FieldLinkCreationInformation FieldLink = new Microsoft.SharePoint.Client.FieldLinkCreationInformation();
                    FieldLink.Field = newfield;
                    ctype.FieldLinks.Add(FieldLink);
                    ctype.Update(true);
                    ctx.ExecuteQuery();
                }
            }
        }

        private static void LoadSiteColumnToContentType(ClientContext ctx, Microsoft.SharePoint.Client.Field newfield, ContentType ctype, string groupname, Microsoft.Extensions.Logging.ILogger log)
        {
            Microsoft.SharePoint.Client.FieldLinkCreationInformation FieldLink = new Microsoft.SharePoint.Client.FieldLinkCreationInformation();
            FieldLink.Field = newfield;
            ctype.FieldLinks.Add(FieldLink);
            ctype.Update(true);
            ctx.ExecuteQuery();
        }

        private static List<SiteColumnSettings> LoadSiteColumnsMaster(ClientContext ctx) {

            List<SiteColumnSettings> result = new List<SiteColumnSettings>();

            // Cargamos la lista maestra de columnas
            List scmasterlist = ctx.Web.Lists.GetByTitle(k_list_sitecolumnsettings);
            ctx.Load(scmasterlist);
            ctx.Load(scmasterlist.Fields);
            ctx.ExecuteQuery();

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection scitems = scmasterlist.GetItems(camlQuery);
            ctx.Load(scitems);
            ctx.ExecuteQuery();

            foreach (Microsoft.SharePoint.Client.ListItem scitem in scitems)
            {
                SiteColumnSettings setting = new SiteColumnSettings(
                    (scitem[k_field_Title] == null) ? "" : scitem[k_field_Title].ToString(),
                    (scitem[k_field_Display] == null) ? "" : scitem[k_field_Display].ToString(),
                    (scitem[k_field_TipoDato] == null) ? "" : scitem[k_field_TipoDato].ToString(),
                    (scitem[k_field_Tipo] == null) ? "" : scitem[k_field_Tipo].ToString(),
                    (scitem[k_field_Statutory] == null) ? false : Convert.ToBoolean(scitem[k_field_Statutory]),
                    (scitem[k_field_StatutoryExcelColumn] == null) ? 0 : Convert.ToInt16(scitem[k_field_StatutoryExcelColumn].ToString()),
                    (scitem[k_field_Reporting] == null) ? false : Convert.ToBoolean(scitem[k_field_Reporting]),
                    (scitem[k_field_ReportingExcelColumn] == null) ? 0 : Convert.ToInt16(scitem[k_field_ReportingExcelColumn].ToString()),
                    (scitem[k_field_Certificates] == null) ? false : Convert.ToBoolean(scitem[k_field_Certificates]),
                    (scitem[k_field_CertificatesExcelColumn] == null) ? 0 : Convert.ToInt16(scitem[k_field_CertificatesExcelColumn].ToString())
                    );

                result.Add(setting);
            }

            return result;
        }

        private static string LoadSiteGroupFromDT(string groupname, DataTable SecurityGroupsDT)
        {
            string result = String.Empty;
            foreach (DataRow secrow in SecurityGroupsDT.Rows)
            {
                if (secrow[0].ToString() == groupname)
                {
                    result = secrow[1].ToString();
                    break;
                }
            }

            return result;
        }

        private static void SetSecurityFolders(AZConfig cfg, string groupname, List list, Folder folder, Folder subfolder, Microsoft.SharePoint.Client.Group ownersgroup, Microsoft.SharePoint.Client.Group membersgroup, Microsoft.Extensions.Logging.ILogger log)
        {
            if (!String.IsNullOrEmpty(groupname))
            {
                list.BreakRoleInheritance(true, false);
                list.AddPermissionLevelToGroup(groupname, Microsoft.SharePoint.Client.RoleType.Reader, true);
                list.AddPermissionLevelToGroup(ownersgroup.Title, Microsoft.SharePoint.Client.RoleType.Administrator, true);
                list.AddPermissionLevelToGroup(membersgroup.Title, Microsoft.SharePoint.Client.RoleType.Contributor, true);
                //list.AddPermissionLevelToUser(cfg._MembersGroup, Microsoft.SharePoint.Client.RoleType.Contributor, true);

                folder.ListItemAllFields.BreakRoleInheritance(false, true);
                folder.ListItemAllFields.AddPermissionLevelToGroup(groupname, Microsoft.SharePoint.Client.RoleType.Reader, true);
                folder.ListItemAllFields.AddPermissionLevelToGroup(ownersgroup.Title, Microsoft.SharePoint.Client.RoleType.Administrator, true);
                folder.ListItemAllFields.AddPermissionLevelToGroup(membersgroup.Title, Microsoft.SharePoint.Client.RoleType.Contributor, true);
                //folder.ListItemAllFields.AddPermissionLevelToUser(cfg._MembersGroup, Microsoft.SharePoint.Client.RoleType.Contributor, true);

                subfolder.ListItemAllFields.BreakRoleInheritance(false, true);
                subfolder.ListItemAllFields.AddPermissionLevelToGroup(groupname, Microsoft.SharePoint.Client.RoleType.Contributor, true);
                subfolder.ListItemAllFields.AddPermissionLevelToGroup(ownersgroup.Title, Microsoft.SharePoint.Client.RoleType.Administrator, true);
                subfolder.ListItemAllFields.AddPermissionLevelToGroup(membersgroup.Title, Microsoft.SharePoint.Client.RoleType.Contributor, true);
                //subfolder.ListItemAllFields.AddPermissionLevelToUser(cfg._MembersGroup, Microsoft.SharePoint.Client.RoleType.Contributor, true);
            }
            else
            {
                log.LogWarning(String.Format("No se ha encontrado el grupo de SharePoint '{0}' para aplicar la seguridad.", groupname));
            }
        }

        private static void SetSecurityItem(AZConfig cfg, string groupname, Microsoft.SharePoint.Client.ListItem item, Microsoft.SharePoint.Client.Group ownersgroup, Microsoft.SharePoint.Client.Group membersgroup, Microsoft.SharePoint.Client.RoleType roletype, Microsoft.Extensions.Logging.ILogger log)
        {
            if (!String.IsNullOrEmpty(groupname))
            {
                item.BreakRoleInheritance(false, true);
                item.AddPermissionLevelToGroup(groupname, roletype, true);
                item.AddPermissionLevelToGroup(ownersgroup.Title, Microsoft.SharePoint.Client.RoleType.Contributor, true);
                item.AddPermissionLevelToGroup(membersgroup.Title, Microsoft.SharePoint.Client.RoleType.Contributor, true);
                //item.AddPermissionLevelToUser(cfg._MembersGroup, Microsoft.SharePoint.Client.RoleType.Contributor, true);
            }
            else
            {
                log.LogWarning(String.Format("No se ha encontrado el grupo de SharePoint '{0}' para aplicar la seguridad.", groupname));
            }
        }
        
        private static void CheckSiteColumns(ClientContext ctx, Microsoft.Extensions.Logging.ILogger log)
        {
            List<SiteColumnSettings> sitecolumns = LoadSiteColumnsMaster(ctx);

            log.LogInformation("Cargando columnas de sitio.");
            FieldCollection fields = ctx.Web.Fields;
            ctx.Load(fields);
            ctx.ExecuteQuery();
            List<Microsoft.SharePoint.Client.Field> fccsitecolumns = fields.Where(x => x.Group == k_sitecolumn_group).ToList();

            foreach (SiteColumnSettings fieldsetting in sitecolumns)
            {
                bool existe = false;
                foreach (Microsoft.SharePoint.Client.Field sitecolumn in fccsitecolumns)
                {
                    if (sitecolumn.InternalName == fieldsetting.Title)
                    {
                        existe = true;
                        break;
                    }
                }

                // Si el campo no existe lo agregamos a las columnas de sitio y a cada content type donde pertenece
                if (!existe)
                {
                    log.LogError("Columna de sitio '" + fieldsetting.Title + "' no existe. La creamos.");

                    // Creamos la columna de sitio
                    ctx.Web.Fields.AddFieldAsXml("<Field Type='" + fieldsetting.TipoDato + "' Name='" + fieldsetting.Title + "' DisplayName='" + fieldsetting.Display + "' Group='" + k_sitecolumn_group + "'/>", true, Microsoft.SharePoint.Client.AddFieldOptions.AddFieldToDefaultView);
                    ctx.Web.Update();
                    ctx.ExecuteQuery();

                    // Cargamos la columna que acabamos de crear
                    fields = ctx.Web.Fields;
                    ctx.Load(fields);
                    ctx.ExecuteQuery();
                    Microsoft.SharePoint.Client.Field newfield = fields.Where(x => x.Group == k_sitecolumn_group && x.InternalName == fieldsetting.Title).FirstOrDefault();

                    if (fieldsetting.Statutory == true)
                    {
                        log.LogWarning("Añadimos columna de sitio '" + fieldsetting.Title + "' a Statutory.");
                        LoadSiteColumn(ctx, newfield, k_contenttype_group_statutory, log);
                    }
                    if (fieldsetting.Reporting == true)
                    {
                        log.LogWarning("Añadimos columna de sitio '" + fieldsetting.Title + "' a Reporting.");
                        LoadSiteColumn(ctx, newfield, k_contenttype_group_reporting, log);
                    }
                    if (fieldsetting.Certificates == true)
                    {
                        log.LogWarning("Añadimos columna de sitio '" + fieldsetting.Title + "' a Certificates.");
                        LoadSiteColumn(ctx, newfield, k_contenttype_group_certificates, log);
                    }
                }

            }
        }

        private static void CheckContentTypesStatutory(ClientContext ctx, DataTable DocumentTypesDT, string groupname, Microsoft.Extensions.Logging.ILogger log)
        {
            ContentTypeCollection ctypes = ctx.Web.ContentTypes;
            ctx.Load(ctypes);
            ctx.ExecuteQuery();

            log.LogWarning("Revisando tipos de contenido de '" + groupname + "'");

            foreach (DataRow row in DocumentTypesDT.Rows)
            {
                if (groupname == ("FCC "+row[0].ToString()))
                {
                    List<Microsoft.SharePoint.Client.ContentType> ctypeslist = ctypes.Where(x => x.Name == row[1].ToString() && x.Group == groupname).ToList();

                    if (ctypeslist.Count == 0)
                    {
                        // Creamos el grupo                                
                        log.LogError("El tipo de contenido '" + row[1].ToString() + "' no existe. Lo creamos");

                        ContentType parentctype = ctypes.Where(x => x.Name == "Item").First();

                        ContentTypeCreationInformation CTypeCreationInfo = new ContentTypeCreationInformation();
                        CTypeCreationInfo.Name = row[1].ToString();
                        CTypeCreationInfo.Description = "";
                        CTypeCreationInfo.Group = groupname;
                        CTypeCreationInfo.ParentContentType = parentctype;

                        ContentType newctype = ctypes.Add(CTypeCreationInfo);
                        ctx.ExecuteQuery();

                        // Añadimos todas las columnas del tipo
                        List<SiteColumnSettings> sitecolumns = LoadSiteColumnsMaster(ctx);

                        FieldCollection fields = ctx.Web.Fields;
                        ctx.Load(fields);
                        ctx.ExecuteQuery();

                        foreach (SiteColumnSettings fieldsetting in sitecolumns)
                        {
                            Microsoft.SharePoint.Client.Field newfield = fields.Where(x => x.Group == k_sitecolumn_group && x.InternalName == fieldsetting.Title).FirstOrDefault();
                            if (fieldsetting.Statutory == true)
                            {
                                log.LogWarning("Añadimos columna de sitio '" + fieldsetting.Title + "' a '" + groupname + "'.");
                                LoadSiteColumnToContentType(ctx, newfield, newctype, groupname, log);
                            }
                        }

                        ctypes = ctx.Web.ContentTypes;
                        ctx.Load(ctypes);
                        ctx.ExecuteQuery();

                        // Añadimos el contenttype a la lista
                        List list = ctx.Web.Lists.GetByTitle(k_list_statutory);
                        ctx.Load(list);
                        ctx.Load(list.ContentTypes);
                        ctx.ExecuteQuery();

                        list.ContentTypes.AddExistingContentType(newctype);
                        ctx.ExecuteQuery();
                    }
                    else
                    {
                        log.LogInformation("El tipo de contenido '" + row[1].ToString() + "' ya existe.");
                    }
                }
                
            }
        }

        private static void CheckContentTypesReporting(ClientContext ctx, DataTable DocumentTypesDT, string groupname, Microsoft.Extensions.Logging.ILogger log)
        {
            ContentTypeCollection ctypes = ctx.Web.ContentTypes;
            ctx.Load(ctypes);
            ctx.ExecuteQuery();

            log.LogWarning("Revisando tipos de contenido de '" + groupname + "'");

            foreach (DataRow row in DocumentTypesDT.Rows)
            {
                if (groupname == ("FCC " + row[0].ToString()))
                {
                    List<Microsoft.SharePoint.Client.ContentType> ctypeslist = ctypes.Where(x => x.Name == row[1].ToString() && x.Group == groupname).ToList();

                    if (ctypeslist.Count == 0)
                    {
                        // Creamos el grupo                                
                        log.LogError("El tipo de contenido '" + row[1].ToString() + "' no existe. Lo creamos");

                        ContentType parentctype = ctypes.Where(x => x.Name == "Item").First();

                        ContentTypeCreationInformation CTypeCreationInfo = new ContentTypeCreationInformation();
                        CTypeCreationInfo.Name = row[1].ToString();
                        CTypeCreationInfo.Description = "";
                        CTypeCreationInfo.Group = groupname;
                        CTypeCreationInfo.ParentContentType = parentctype;

                        ContentType newctype = ctypes.Add(CTypeCreationInfo);
                        ctx.ExecuteQuery();

                        // Añadimos todas las columnas del tipo
                        List<SiteColumnSettings> sitecolumns = LoadSiteColumnsMaster(ctx);

                        FieldCollection fields = ctx.Web.Fields;
                        ctx.Load(fields);
                        ctx.ExecuteQuery();

                        foreach (SiteColumnSettings fieldsetting in sitecolumns)
                        {
                            Microsoft.SharePoint.Client.Field newfield = fields.Where(x => x.Group == k_sitecolumn_group && x.InternalName == fieldsetting.Title).FirstOrDefault();
                            if (fieldsetting.Reporting == true)
                            {
                                log.LogWarning("Añadimos columna de sitio '" + fieldsetting.Title + "' a '"+groupname+"'.");
                                LoadSiteColumnToContentType(ctx, newfield, newctype, groupname, log);
                            }
                        }

                        ctypes = ctx.Web.ContentTypes;
                        ctx.Load(ctypes);
                        ctx.ExecuteQuery();

                        // Añadimos el contenttype a la lista
                        List list = ctx.Web.Lists.GetByTitle(k_list_reporting);
                        ctx.Load(list);
                        ctx.Load(list.ContentTypes);
                        ctx.ExecuteQuery();

                        list.ContentTypes.AddExistingContentType(newctype);
                        ctx.ExecuteQuery();

                    }
                    else
                    {
                        log.LogInformation("El tipo de contenido '" + row[1].ToString() + "' ya existe.");
                    }
                }

            }
        }

        private static void CheckContentTypesCertificates(ClientContext ctx, DataTable DocumentTypesDT, string groupname, Microsoft.Extensions.Logging.ILogger log)
        {
            ContentTypeCollection ctypes = ctx.Web.ContentTypes;
            ctx.Load(ctypes);
            ctx.ExecuteQuery();

            log.LogWarning("Revisando tipos de contenido de '"+groupname+"'");

            foreach (DataRow row in DocumentTypesDT.Rows)
            {
                if (groupname == ("FCC " + row[0].ToString()))
                {
                    List<Microsoft.SharePoint.Client.ContentType> ctypeslist = ctypes.Where(x => x.Name == row[1].ToString() && x.Group == groupname).ToList();

                    if (ctypeslist.Count == 0)
                    {
                        // Creamos el grupo                                
                        log.LogError("El tipo de contenido '" + row[1].ToString() + "' no existe. Lo creamos");

                        ContentType parentctype = ctypes.Where(x => x.Name == "Item").First();

                        ContentTypeCreationInformation CTypeCreationInfo = new ContentTypeCreationInformation();
                        CTypeCreationInfo.Name = row[1].ToString();
                        CTypeCreationInfo.Description = "";
                        CTypeCreationInfo.Group = groupname;
                        CTypeCreationInfo.ParentContentType = parentctype;

                        ContentType newctype = ctypes.Add(CTypeCreationInfo);
                        ctx.ExecuteQuery();

                        // Añadimos todas las columnas del tipo
                        List<SiteColumnSettings> sitecolumns = LoadSiteColumnsMaster(ctx);

                        FieldCollection fields = ctx.Web.Fields;
                        ctx.Load(fields);
                        ctx.ExecuteQuery();

                        foreach (SiteColumnSettings fieldsetting in sitecolumns)
                        {
                            Microsoft.SharePoint.Client.Field newfield = fields.Where(x => x.Group == k_sitecolumn_group && x.InternalName == fieldsetting.Title).FirstOrDefault();
                            if (fieldsetting.Certificates == true)
                            {
                                log.LogWarning("Añadimos columna de sitio '" + fieldsetting.Title + "' a '" + groupname + "'.");
                                LoadSiteColumnToContentType(ctx, newfield, newctype, groupname, log);
                            }
                        }

                        ctypes = ctx.Web.ContentTypes;
                        ctx.Load(ctypes);
                        ctx.ExecuteQuery();

                        // Añadimos el contenttype a la lista
                        List list = ctx.Web.Lists.GetByTitle(k_list_certificates);
                        ctx.Load(list);
                        ctx.Load(list.ContentTypes);
                        ctx.ExecuteQuery();

                        list.ContentTypes.AddExistingContentType(newctype);
                        ctx.ExecuteQuery();
                    }
                }
                else
                {
                    log.LogInformation("El tipo de contenido '" + row[1].ToString() + "' ya existe.");
                }

            }
        }

        private static void CheckSiteGroups(ClientContext ctx, DataTable SecurityGroupsDT, Microsoft.Extensions.Logging.ILogger log)
        {
            GroupCollection sitegroups = ctx.Web.SiteGroups;
            ctx.Load(sitegroups);
            ctx.ExecuteQuery();

            foreach (DataRow row in SecurityGroupsDT.Rows)
            {
                List<Microsoft.SharePoint.Client.Group> groups = sitegroups.Where(x => x.Title == row[1].ToString()).ToList();
                if (groups.Count == 0)
                {
                    // Creamos el grupo                                
                    log.LogError("El grupo '" + row[1].ToString() + "' no existe. Lo creamos");

                    Microsoft.SharePoint.Client.GroupCreationInformation GroupInfo = new GroupCreationInformation();
                    GroupInfo.Title = row[1].ToString();
                    ctx.Web.SiteGroups.Add(GroupInfo);
                    ctx.ExecuteQuery();

                }
            }
        }

        private static void LoadDataStatutory(AZConfig cfg, ClientContext ctx, DataTable StatutoryDT, DataTable DocumentTypesDT, DataTable SecurityGroupsDT, Microsoft.Extensions.Logging.ILogger log)
        {
            List<SiteColumnSettings> sitecolumns = LoadSiteColumnsMaster(ctx);
            List<SiteColumnSettings> statutorycols = sitecolumns.Where(x => x.Statutory == true).OrderBy(x => x.StatutoryExcelColumn).ToList();

            List list = ctx.Web.Lists.GetByTitle(k_list_statutory);
            ctx.Load(list);
            ctx.Load(list.Fields);
            ctx.Load(list.RootFolder);
            ctx.ExecuteQuery();

            // Cargamos los content type de statutory
            ContentTypeCollection ctypes = ctx.Web.ContentTypes;
            ctx.Load(ctypes);
            ctx.ExecuteQuery();

            bool consolidated = false;

            Microsoft.SharePoint.Client.Group ownersgroup = ctx.Web.AssociatedOwnerGroup;
            Microsoft.SharePoint.Client.Group membersgroup = ctx.Web.AssociatedMemberGroup;
            ctx.Load(ownersgroup);
            ctx.Load(membersgroup);
            ctx.ExecuteQuery();

            foreach (DataRow row in StatutoryDT.Rows)
            {
                log.LogInformation("Creando carpeta principal y secundaria.");

                string foldername = AddFolder(ctx, list, row, statutorycols, log);
                string subfoldername = AddSubfolder(ctx, list, foldername, row, statutorycols, log);

                if( (!string.IsNullOrEmpty(foldername)) && (!string.IsNullOrEmpty(subfoldername)) )
                {
                    string folderurl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, foldername);
                    string subfolderurl = String.Format("{0}/{1}/{2}", list.RootFolder.ServerRelativeUrl, foldername, subfoldername);

                    log.LogInformation("Carpetas creadas.");

                    log.LogWarning(folderurl);
                    log.LogWarning(subfolderurl);

                    Folder folder = ctx.Web.GetFolderByServerRelativeUrl(folderurl);
                    Folder subfolder = ctx.Web.GetFolderByServerRelativeUrl(subfolderurl);
                    ctx.Load(folder);
                    ctx.Load(subfolder);
                    ctx.ExecuteQuery();

                    string shgroupname = row[k_statutory_groupindex].ToString();
                    string groupname = LoadSiteGroupFromDT(shgroupname.Trim(), SecurityGroupsDT);

                    log.LogInformation("Aplicamos seguridad a las carpetas. [" + groupname + "]");

                    SetSecurityFolders(cfg, groupname, list, folder, subfolder, ownersgroup, membersgroup, log);

                    log.LogWarning("Seguridad aplicada");

                    foreach (DataRow doctyperow in DocumentTypesDT.Rows)
                    {
                        if (doctyperow[0].ToString() == "Statutory")
                        {
                            Microsoft.SharePoint.Client.ContentType ctype = ctypes.Where(x => x.Name == doctyperow[1].ToString() && x.Group == k_contenttype_group_statutory).First();

                            log.LogInformation(ctype.Name);

                            ListItemCreationInformation ListItemInfo = new ListItemCreationInformation();
                            ListItemInfo.FolderUrl = subfolder.ServerRelativeUrl;
                            ListItemInfo.LeafName = ctype.Name;
                            Microsoft.SharePoint.Client.ListItem item = list.AddItem(ListItemInfo);

                            for (int i = 0; i < statutorycols.Count; i++)
                            {
                                if (statutorycols[i].Title == "AI_E_Tipo" && row[i].ToString() == "x")
                                {
                                    item[statutorycols[i].Title] = "Individual";
                                    consolidated = true;
                                    log.LogInformation("Consolidated: TRUE");
                                }
                                else
                                {
                                    item[statutorycols[i].Title] = row[i].ToString().Trim();
                                }
                            }

                            item["ContentTypeId"] = ctype.Id;
                            item.Update();
                            ctx.ExecuteQuery();

                            // Comprobamos si el content type es de solo lectura. En ese caso, el grupo tendrá acceso de solo lectura a ese item
                            if (doctyperow[2].ToString() == "x")
                            {
                                log.LogWarning("Acceso de reader al grupo de seguridad '" + groupname + "' (Individual)");
                                SetSecurityItem(cfg, groupname, item, ownersgroup, membersgroup, Microsoft.SharePoint.Client.RoleType.Reader, log);
                            }
                            else
                            {
                                log.LogWarning("Acceso de contributor al grupo de seguridad '" + groupname + "' (Individual)");
                                SetSecurityItem(cfg, groupname, item, ownersgroup, membersgroup, Microsoft.SharePoint.Client.RoleType.Contributor, log);
                            }


                            // Si es consolidated entonces tenemos que guardar los datos duplicados (Individual y Consolidado)
                            if (consolidated)
                            {
                                ctype = ctypes.Where(x => x.Name == doctyperow[1].ToString() && x.Group == k_contenttype_group_statutory).First();

                                log.LogInformation(ctype.Name + " [Consolidated]");

                                ListItemInfo = new ListItemCreationInformation();
                                ListItemInfo.FolderUrl = subfolder.ServerRelativeUrl;
                                item = list.AddItem(ListItemInfo);

                                for (int i = 0; i < statutorycols.Count; i++)
                                {
                                    if (statutorycols[i].Title == "AI_E_Tipo" && row[i].ToString() == "x")
                                    {
                                        item[statutorycols[i].Title] = "Consolidated";
                                        consolidated = false;
                                    }
                                    else
                                    {
                                        item[statutorycols[i].Title] = row[i].ToString().Trim();
                                    }

                                }

                                item["ContentTypeId"] = ctype.Id;
                                item.Update();
                                ctx.ExecuteQuery();

                                // Comprobamos si el content type es de solo lectura. En ese caso, el grupo tendrá acceso de solo lectura a ese item
                                if (doctyperow[2].ToString() == "x")
                                {
                                    log.LogWarning("Acceso de reader al grupo de seguridad '" + groupname + "' (Consolidated)");
                                    SetSecurityItem(cfg, groupname, item, ownersgroup, membersgroup, Microsoft.SharePoint.Client.RoleType.Reader, log);
                                }
                                else
                                {
                                    log.LogWarning("Acceso de contributor al grupo de seguridad '" + groupname + "' (Consolidated)");
                                    SetSecurityItem(cfg, groupname, item, ownersgroup, membersgroup, Microsoft.SharePoint.Client.RoleType.Contributor, log);
                                }
                            }

                        } // END if Statutory content type

                    }// END foreach Doc Type
                }
                else
                {
                    log.LogError("La carpeta o subcarpeta no se ha podido crear");
                }
            } // END foreach Statutory row

        }

        private static void LoadDataReporting(AZConfig cfg, ClientContext ctx, DataTable ReportingDT, DataTable DocumentTypesDT, DataTable SecurityGroupsDT, Microsoft.Extensions.Logging.ILogger log)
        {
            List<SiteColumnSettings> sitecolumns = LoadSiteColumnsMaster(ctx);
            List<SiteColumnSettings> reportingcols = sitecolumns.Where(x => x.Reporting == true).OrderBy(x => x.ReportingExcelColumn).ToList();

            List list = ctx.Web.Lists.GetByTitle(k_list_reporting);
            ctx.Load(list);
            ctx.Load(list.Fields);
            ctx.Load(list.RootFolder);
            ctx.ExecuteQuery();

            // Cargamos los content type de statutory
            ContentTypeCollection ctypes = ctx.Web.ContentTypes;
            ctx.Load(ctypes);
            ctx.ExecuteQuery();

            Microsoft.SharePoint.Client.Group ownersgroup = ctx.Web.AssociatedOwnerGroup;
            Microsoft.SharePoint.Client.Group membersgroup = ctx.Web.AssociatedMemberGroup;
            ctx.Load(ownersgroup);
            ctx.Load(membersgroup);
            ctx.ExecuteQuery();

            foreach (DataRow row in ReportingDT.Rows)
            {
                log.LogInformation("Creando carpeta principal y secundaria.");

                string foldername = AddFolder(ctx, list, row, reportingcols, log);
                string subfoldername = AddSubfolder(ctx, list, foldername, row, reportingcols, log);

                if ((!string.IsNullOrEmpty(foldername)) && (!string.IsNullOrEmpty(subfoldername)))
                {
                    string folderurl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, foldername);
                    string subfolderurl = String.Format("{0}/{1}/{2}", list.RootFolder.ServerRelativeUrl, foldername, subfoldername);

                    log.LogInformation("Carpetas creadas.");

                    log.LogWarning(folderurl);
                    log.LogWarning(subfolderurl);

                    Folder folder = ctx.Web.GetFolderByServerRelativeUrl(folderurl);
                    Folder subfolder = ctx.Web.GetFolderByServerRelativeUrl(subfolderurl);
                    ctx.Load(folder);
                    ctx.Load(subfolder);
                    ctx.ExecuteQuery();


                    string shgroupname = row[k_reporting_groupindex].ToString();
                    string groupname = LoadSiteGroupFromDT(shgroupname, SecurityGroupsDT);

                    log.LogInformation("Aplicamos seguridad a las carpetas. [" + groupname + "]");

                    SetSecurityFolders(cfg, groupname, list, folder, subfolder, ownersgroup, membersgroup, log);

                    log.LogWarning("Seguridad aplicada");

                    foreach (DataRow doctyperow in DocumentTypesDT.Rows)
                    {
                        if (doctyperow[0].ToString() == "Reporting")
                        {
                            Microsoft.SharePoint.Client.ContentType ctype = ctypes.Where(x => x.Name == doctyperow[1].ToString() && x.Group == k_contenttype_group_reporting).First();

                            log.LogInformation(ctype.Name);

                            ListItemCreationInformation ListItemInfo = new ListItemCreationInformation();
                            ListItemInfo.FolderUrl = subfolder.ServerRelativeUrl;
                            ListItemInfo.LeafName = ctype.Name;
                            Microsoft.SharePoint.Client.ListItem item = list.AddItem(ListItemInfo);

                            for (int i = 0; i < reportingcols.Count; i++)
                            {
                                item[reportingcols[i].Title] = row[i].ToString().Trim();
                            }

                            item["ContentTypeId"] = ctype.Id;
                            item.Update();
                            ctx.ExecuteQuery();

                            // Comprobamos si el content type es de solo lectura. En ese caso, el grupo tendrá acceso de solo lectura a ese item
                            if (doctyperow[2].ToString() == "x")
                            {
                                log.LogWarning("Acceso de reader al grupo de seguridad '" + groupname + "' (Individual)");
                                SetSecurityItem(cfg, groupname, item, ownersgroup, membersgroup, Microsoft.SharePoint.Client.RoleType.Reader, log);
                            }
                            else
                            {
                                log.LogWarning("Acceso de contributor al grupo de seguridad '" + groupname + "' (Individual)");
                                SetSecurityItem(cfg, groupname, item, ownersgroup, membersgroup, Microsoft.SharePoint.Client.RoleType.Contributor, log);
                            }

                        } // END if Reporting content type

                    }// END foreach Doc Type
                }
                else
                {
                    log.LogError("La carpeta o subcarpeta no se ha podido crear");
                }

            } // END foreach Statutory row

        }

        private static void LoadDataCertificates(AZConfig cfg, ClientContext ctx, DataTable CertificatesDT, DataTable DocumentTypesDT, DataTable SecurityGroupsDT, Microsoft.Extensions.Logging.ILogger log)
        {
            List<SiteColumnSettings> sitecolumns = LoadSiteColumnsMaster(ctx);
            List<SiteColumnSettings> certificatescols = sitecolumns.Where(x => x.Certificates == true).OrderBy(x => x.CertificatesExcelColumn).ToList();

            List list = ctx.Web.Lists.GetByTitle(k_list_certificates);
            ctx.Load(list);
            ctx.Load(list.Fields);
            ctx.Load(list.RootFolder);
            ctx.ExecuteQuery();

            // Cargamos los content type de statutory
            ContentTypeCollection ctypes = ctx.Web.ContentTypes;
            ctx.Load(ctypes);
            ctx.ExecuteQuery();

            Microsoft.SharePoint.Client.Group ownersgroup = ctx.Web.AssociatedOwnerGroup;
            Microsoft.SharePoint.Client.Group membersgroup = ctx.Web.AssociatedMemberGroup;
            ctx.Load(ownersgroup);
            ctx.Load(membersgroup);
            ctx.ExecuteQuery();

            foreach (DataRow row in CertificatesDT.Rows)
            {
                log.LogInformation("Creando carpeta principal y secundaria.");

                string foldername = AddFolder(ctx, list, row, certificatescols, log);
                string subfoldername = AddSubfolder(ctx, list, foldername, row, certificatescols, log);

                if ((!string.IsNullOrEmpty(foldername)) && (!string.IsNullOrEmpty(subfoldername)))
                {
                    string folderurl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, foldername);
                    string subfolderurl = String.Format("{0}/{1}/{2}", list.RootFolder.ServerRelativeUrl, foldername, subfoldername);

                    log.LogInformation("Carpetas creadas.");

                    log.LogWarning(folderurl);
                    log.LogWarning(subfolderurl);

                    Folder folder = ctx.Web.GetFolderByServerRelativeUrl(folderurl);
                    Folder subfolder = ctx.Web.GetFolderByServerRelativeUrl(subfolderurl);
                    ctx.Load(folder);
                    ctx.Load(subfolder);
                    ctx.ExecuteQuery();


                    string shgroupname = row[k_certificates_groupindex].ToString();
                    string groupname = LoadSiteGroupFromDT(shgroupname, SecurityGroupsDT);

                    log.LogInformation("Aplicamos seguridad a las carpetas. [" + groupname + "]");

                    SetSecurityFolders(cfg, groupname, list, folder, subfolder, ownersgroup, membersgroup, log);

                    log.LogWarning("Seguridad aplicada");

                    foreach (DataRow doctyperow in DocumentTypesDT.Rows)
                    {
                        if (doctyperow[0].ToString() == "Certificates")
                        {
                            Microsoft.SharePoint.Client.ContentType ctype = ctypes.Where(x => x.Name == doctyperow[1].ToString() && x.Group == k_contenttype_group_certificates).First();

                            log.LogInformation(ctype.Name);

                            ListItemCreationInformation ListItemInfo = new ListItemCreationInformation();
                            ListItemInfo.FolderUrl = subfolder.ServerRelativeUrl;
                            ListItemInfo.LeafName = ctype.Name;
                            Microsoft.SharePoint.Client.ListItem item = list.AddItem(ListItemInfo);

                            for (int i = 0; i < certificatescols.Count; i++)
                            {
                                item[certificatescols[i].Title] = row[i].ToString().Trim();
                            }

                            item["ContentTypeId"] = ctype.Id;
                            item.Update();
                            ctx.ExecuteQuery();

                            // Comprobamos si el content type es de solo lectura. En ese caso, el grupo tendrá acceso de solo lectura a ese item
                            if (doctyperow[2].ToString() == "x")
                            {
                                log.LogWarning("Acceso de reader al grupo de seguridad '" + groupname + "' (Individual)");
                                SetSecurityItem(cfg, groupname, item, ownersgroup, membersgroup, Microsoft.SharePoint.Client.RoleType.Reader, log);
                            }
                            else
                            {
                                log.LogWarning("Acceso de contributor al grupo de seguridad '" + groupname + "' (Individual)");
                                SetSecurityItem(cfg, groupname, item, ownersgroup, membersgroup, Microsoft.SharePoint.Client.RoleType.Contributor, log);
                            }

                        } // END if Reporting content type

                    }// END foreach Doc Type
                }
                else
                {
                    log.LogError("La carpeta o subcarpeta no se ha podido crear");
                }
            } // END foreach Statutory row

        }

    }
        
}
