namespace b2xtranslator.OpenXmlLib.SpreadsheetML
{
    public class Sml
    {
        public const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        #region Names defined in sml-autoFilter.xsd
        public class AutoFilter
        {
            /// <summary>
            /// AutoFilter Column
            /// </summary>
            public const string ElFilterColumn = "filterColumn";

            /// <summary>
            /// Sort State for Auto Filter
            /// </summary>
            public const string ElSortState = "sortState";

            /// <summary>
            /// 
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Filter Criteria
            /// </summary>
            public const string ElFilters = "filters";

            /// <summary>
            /// Top 10
            /// </summary>
            public const string ElTop10 = "top10";

            /// <summary>
            /// Custom Filters
            /// </summary>
            public const string ElCustomFilters = "customFilters";

            /// <summary>
            /// Dynamic Filter
            /// </summary>
            public const string ElDynamicFilter = "dynamicFilter";

            /// <summary>
            /// Color Filter Criteria
            /// </summary>
            public const string ElColorFilter = "colorFilter";

            /// <summary>
            /// Icon Filter
            /// </summary>
            public const string ElIconFilter = "iconFilter";

            /// <summary>
            /// Filter
            /// </summary>
            public const string ElFilter = "filter";

            /// <summary>
            /// Date Grouping
            /// </summary>
            public const string ElDateGroupItem = "dateGroupItem";

            /// <summary>
            /// Custom Filter Criteria
            /// </summary>
            public const string ElCustomFilter = "customFilter";

            /// <summary>
            /// Sort Condition
            /// </summary>
            public const string ElSortCondition = "sortCondition";

            /// <summary>
            /// Cell or Range Reference
            /// </summary>
            public const string AttrRef = "ref";

            /// <summary>
            /// Filter Column Data
            /// </summary>
            public const string AttrColId = "colId";

            /// <summary>
            /// Hidden AutoFilter Button
            /// </summary>
            public const string AttrHiddenButton = "hiddenButton";

            /// <summary>
            /// Show Filter Button
            /// </summary>
            public const string AttrShowButton = "showButton";

            /// <summary>
            /// Filter by Blank
            /// </summary>
            public const string AttrBlank = "blank";

            /// <summary>
            /// Calendar Type
            /// </summary>
            public const string AttrCalendarType = "calendarType";

            /// <summary>
            /// Filter Value
            /// </summary>
            public const string AttrVal = "val";

            /// <summary>
            /// And
            /// </summary>
            public const string AttrAnd = "and";

            /// <summary>
            /// Filter Comparison Operator
            /// </summary>
            public const string AttrOperator = "operator";

            /// <summary>
            /// Top
            /// </summary>
            public const string AttrTop = "top";

            /// <summary>
            /// Filter by Percent
            /// </summary>
            public const string AttrPercent = "percent";

            /// <summary>
            /// Filter Value
            /// </summary>
            public const string AttrFilterVal = "filterVal";

            /// <summary>
            /// Differential Format Record Id
            /// </summary>
            public const string AttrDxfId = "dxfId";

            /// <summary>
            /// Filter By Cell Color
            /// </summary>
            public const string AttrCellColor = "cellColor";

            /// <summary>
            /// Icon Set
            /// </summary>
            public const string AttrIconSet = "iconSet";

            /// <summary>
            /// Icon Id
            /// </summary>
            public const string AttrIconId = "iconId";

            /// <summary>
            /// Dynamic filter type
            /// </summary>
            public const string AttrType = "type";

            /// <summary>
            /// Max Value
            /// </summary>
            public const string AttrMaxVal = "maxVal";

            /// <summary>
            /// Sort by Columns
            /// </summary>
            public const string AttrColumnSort = "columnSort";

            /// <summary>
            /// Case Sensitive
            /// </summary>
            public const string AttrCaseSensitive = "caseSensitive";

            /// <summary>
            /// Sort Method
            /// </summary>
            public const string AttrSortMethod = "sortMethod";

            /// <summary>
            /// Descending
            /// </summary>
            public const string AttrDescending = "descending";

            /// <summary>
            /// Sort By
            /// </summary>
            public const string AttrSortBy = "sortBy";

            /// <summary>
            /// Custom List
            /// </summary>
            public const string AttrCustomList = "customList";

            /// <summary>
            /// Year
            /// </summary>
            public const string AttrYear = "year";

            /// <summary>
            /// Month
            /// </summary>
            public const string AttrMonth = "month";

            /// <summary>
            /// Day
            /// </summary>
            public const string AttrDay = "day";

            /// <summary>
            /// Hour
            /// </summary>
            public const string AttrHour = "hour";

            /// <summary>
            /// Minute
            /// </summary>
            public const string AttrMinute = "minute";

            /// <summary>
            /// Second
            /// </summary>
            public const string AttrSecond = "second";

            /// <summary>
            /// Date Time Grouping
            /// </summary>
            public const string AttrDateTimeGrouping = "dateTimeGrouping";

        }
        #endregion

        #region Names defined in sml-baseTypes.xsd
        public class BaseTypes
        {
            /// <summary>
            /// Extension
            /// </summary>
            public const string ElExt = "ext";

            /// <summary>
            /// Value
            /// </summary>
            public const string AttrV = "v";

            /// <summary>
            /// URI
            /// </summary>
            public const string AttrUri = "uri";

        }
        #endregion

        #region Names defined in sml-calculationChain.xsd
        public class CalculationChain
        {
            /// <summary>
            /// Calculation Chain Info
            /// </summary>
            public const string ElCalcChain = "calcChain";

            /// <summary>
            /// Cell
            /// </summary>
            public const string ElC = "c";

            /// <summary>
            /// 
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Cell Reference
            /// </summary>
            public const string AttrR = "r";

            /// <summary>
            /// Sheet Id
            /// </summary>
            public const string AttrI = "i";

            /// <summary>
            /// Child Chain
            /// </summary>
            public const string AttrS = "s";

            /// <summary>
            /// New Dependency Level
            /// </summary>
            public const string AttrL = "l";

            /// <summary>
            /// New Thread
            /// </summary>
            public const string AttrT = "t";

            /// <summary>
            /// Array
            /// </summary>
            public const string AttrA = "a";

        }
        #endregion

        #region Names defined in sml-comments.xsd
        public class Comments
        {
            /// <summary>
            /// Comments
            /// </summary>
            public const string ElComments = "comments";

            /// <summary>
            /// Authors
            /// </summary>
            public const string ElAuthors = "authors";

            /// <summary>
            /// List of Comments
            /// </summary>
            public const string ElCommentList = "commentList";

            /// <summary>
            /// 
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Author
            /// </summary>
            public const string ElAuthor = "author";

            /// <summary>
            /// Comment
            /// </summary>
            public const string ElComment = "comment";

            /// <summary>
            /// Comment Text
            /// </summary>
            public const string ElText = "text";

            /// <summary>
            /// Cell Reference
            /// </summary>
            public const string AttrRef = "ref";

            /// <summary>
            /// Author Id
            /// </summary>
            public const string AttrAuthorId = "authorId";

            /// <summary>
            /// Unique Identifier for Comment
            /// </summary>
            public const string AttrGuid = "guid";

        }
        #endregion

        #region Names defined in sml-customXmlMappings.xsd
        public class CustomXmlMappings
        {
            /// <summary>
            /// XML Mapping
            /// </summary>
            public const string ElMapInfo = "MapInfo";

            /// <summary>
            /// XML Schema
            /// </summary>
            public const string ElSchema = "Schema";

            /// <summary>
            /// XML Mapping Properties
            /// </summary>
            public const string ElMap = "Map";

            /// <summary>
            /// XML Mapping
            /// </summary>
            public const string ElDataBinding = "DataBinding";

            /// <summary>
            /// Prefix Mappings for XPath Expressions
            /// </summary>
            public const string AttrSelectionNamespaces = "SelectionNamespaces";

            /// <summary>
            /// Schema ID
            /// </summary>
            public const string AttrID = "ID";

            /// <summary>
            /// Schema Reference
            /// </summary>
            public const string AttrSchemaRef = "SchemaRef";

            /// <summary>
            /// Schema Root Namespace
            /// </summary>
            public const string AttrNamespace = "Namespace";

            /// <summary>
            /// XML Mapping Name
            /// </summary>
            public const string AttrName = "Name";

            /// <summary>
            /// Root Element Name
            /// </summary>
            public const string AttrRootElement = "RootElement";

            /// <summary>
            /// Schema Name
            /// </summary>
            public const string AttrSchemaID = "SchemaID";

            /// <summary>
            /// Show Validation Errors
            /// </summary>
            public const string AttrShowImportExportValidationErrors = "ShowImportExportValidationErrors";

            /// <summary>
            /// AutoFit Table on Refresh
            /// </summary>
            public const string AttrAutoFit = "AutoFit";

            /// <summary>
            /// Append Data to Table
            /// </summary>
            public const string AttrAppend = "Append";

            /// <summary>
            /// Preserve AutoFilter State
            /// </summary>
            public const string AttrPreserveSortAFLayout = "PreserveSortAFLayout";

            /// <summary>
            /// Preserve Cell Formatting
            /// </summary>
            public const string AttrPreserveFormat = "PreserveFormat";

            /// <summary>
            /// Unique Identifer
            /// </summary>
            public const string AttrDataBindingName = "DataBindingName";

            /// <summary>
            /// Binding to External File
            /// </summary>
            public const string AttrFileBinding = "FileBinding";

            /// <summary>
            /// Reference to Connection ID
            /// </summary>
            public const string AttrConnectionID = "ConnectionID";

            /// <summary>
            /// File Binding Name
            /// </summary>
            public const string AttrFileBindingName = "FileBindingName";

            /// <summary>
            /// XML Data Loading Behavior
            /// </summary>
            public const string AttrDataBindingLoadMode = "DataBindingLoadMode";

        }
        #endregion

        #region Names defined in sml-externalConnections.xsd
        public class ExternalConnections
        {
            /// <summary>
            /// Connections
            /// </summary>
            public const string ElConnections = "connections";

            /// <summary>
            /// Connection
            /// </summary>
            public const string ElConnection = "connection";

            /// <summary>
            /// ODBC & OLE DB Properties
            /// </summary>
            public const string ElDbPr = "dbPr";

            /// <summary>
            /// OLAP Properties
            /// </summary>
            public const string ElOlapPr = "olapPr";

            /// <summary>
            /// Web Query Properties
            /// </summary>
            public const string ElWebPr = "webPr";

            /// <summary>
            /// Text Import Settings
            /// </summary>
            public const string ElTextPr = "textPr";

            /// <summary>
            /// Query Parameters
            /// </summary>
            public const string ElParameters = "parameters";

            /// <summary>
            /// Future Feature Data Storage
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Tables
            /// </summary>
            public const string ElTables = "tables";

            /// <summary>
            /// Parameter Properties
            /// </summary>
            public const string ElParameter = "parameter";

            /// <summary>
            /// No Value
            /// </summary>
            public const string ElM = "m";

            /// <summary>
            /// Character Value
            /// </summary>
            public const string ElS = "s";

            /// <summary>
            /// Index
            /// </summary>
            public const string ElX = "x";

            /// <summary>
            /// Fields
            /// </summary>
            public const string ElTextFields = "textFields";

            /// <summary>
            /// Text Import Field Settings
            /// </summary>
            public const string ElTextField = "textField";

            /// <summary>
            /// Connection Id
            /// </summary>
            public const string AttrId = "id";

            /// <summary>
            /// Source Database File
            /// </summary>
            public const string AttrSourceFile = "sourceFile";

            /// <summary>
            /// Connection File
            /// </summary>
            public const string AttrOdcFile = "odcFile";

            /// <summary>
            /// Keep Connection Open
            /// </summary>
            public const string AttrKeepAlive = "keepAlive";

            /// <summary>
            /// Automatic Refresh Interval
            /// </summary>
            public const string AttrInterval = "interval";

            /// <summary>
            /// Connection Name
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// Connection Description
            /// </summary>
            public const string AttrDescription = "description";

            /// <summary>
            /// Database Source Type
            /// </summary>
            public const string AttrType = "type";

            /// <summary>
            /// Reconnection Method
            /// </summary>
            public const string AttrReconnectionMethod = "reconnectionMethod";

            /// <summary>
            /// Last Refresh Version
            /// </summary>
            public const string AttrRefreshedVersion = "refreshedVersion";

            /// <summary>
            /// Minimum Version Required for Refresh
            /// </summary>
            public const string AttrMinRefreshableVersion = "minRefreshableVersion";

            /// <summary>
            /// Save Password
            /// </summary>
            public const string AttrSavePassword = "savePassword";

            /// <summary>
            /// New Connection
            /// </summary>
            public const string AttrNew = "new";

            /// <summary>
            /// Deleted Connection
            /// </summary>
            public const string AttrDeleted = "deleted";

            /// <summary>
            /// Only Use Connection File
            /// </summary>
            public const string AttrOnlyUseConnectionFile = "onlyUseConnectionFile";

            /// <summary>
            /// Background Refresh
            /// </summary>
            public const string AttrBackground = "background";

            /// <summary>
            /// Refresh on Open
            /// </summary>
            public const string AttrRefreshOnLoad = "refreshOnLoad";

            /// <summary>
            /// Save Data
            /// </summary>
            public const string AttrSaveData = "saveData";

            /// <summary>
            /// Reconnection Method
            /// </summary>
            public const string AttrCredentials = "credentials";

            /// <summary>
            /// SSO Id
            /// </summary>
            public const string AttrSingleSignOnId = "singleSignOnId";

            /// <summary>
            /// Command Text
            /// </summary>
            public const string AttrCommand = "command";

            /// <summary>
            /// Command Text
            /// </summary>
            public const string AttrServerCommand = "serverCommand";

            /// <summary>
            /// OLE DB Command Type
            /// </summary>
            public const string AttrCommandType = "commandType";

            /// <summary>
            /// Local Cube
            /// </summary>
            public const string AttrLocal = "local";

            /// <summary>
            /// Local Cube Connection
            /// </summary>
            public const string AttrLocalConnection = "localConnection";

            /// <summary>
            /// Local Refresh
            /// </summary>
            public const string AttrLocalRefresh = "localRefresh";

            /// <summary>
            /// Send Locale to OLAP
            /// </summary>
            public const string AttrSendLocale = "sendLocale";

            /// <summary>
            /// Drill Through Count
            /// </summary>
            public const string AttrRowDrillCount = "rowDrillCount";

            /// <summary>
            /// OLAP Fill Formatting
            /// </summary>
            public const string AttrServerFill = "serverFill";

            /// <summary>
            /// OLAP Number Format
            /// </summary>
            public const string AttrServerNumberFormat = "serverNumberFormat";

            /// <summary>
            /// OLAP Server Font
            /// </summary>
            public const string AttrServerFont = "serverFont";

            /// <summary>
            /// OLAP Font Formatting
            /// </summary>
            public const string AttrServerFontColor = "serverFontColor";

            /// <summary>
            /// XML Source
            /// </summary>
            public const string AttrXml = "xml";

            /// <summary>
            /// Import XML Source Data
            /// </summary>
            public const string AttrSourceData = "sourceData";

            /// <summary>
            /// Parse PRE
            /// </summary>
            public const string AttrParsePre = "parsePre";

            /// <summary>
            /// Consecutive Delimiters
            /// </summary>
            public const string AttrConsecutive = "consecutive";

            /// <summary>
            /// Use First Row
            /// </summary>
            public const string AttrFirstRow = "firstRow";

            /// <summary>
            /// Created in Excel 97
            /// </summary>
            public const string AttrXl97 = "xl97";

            /// <summary>
            /// Dates as Text
            /// </summary>
            public const string AttrTextDates = "textDates";

            /// <summary>
            /// Refreshed in Excel 2000
            /// </summary>
            public const string AttrXl2000 = "xl2000";

            /// <summary>
            /// URL
            /// </summary>
            public const string AttrUrl = "url";

            /// <summary>
            /// Web Post
            /// </summary>
            public const string AttrPost = "post";

            /// <summary>
            /// HTML Tables Only
            /// </summary>
            public const string AttrHtmlTables = "htmlTables";

            /// <summary>
            /// HTML Formatting Handling
            /// </summary>
            public const string AttrHtmlFormat = "htmlFormat";

            /// <summary>
            /// Edit Query URL
            /// </summary>
            public const string AttrEditPage = "editPage";

            /// <summary>
            /// Parameter Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// SQL Data Type
            /// </summary>
            public const string AttrSqlType = "sqlType";

            /// <summary>
            /// Parameter Type
            /// </summary>
            public const string AttrParameterType = "parameterType";

            /// <summary>
            /// Refresh on Change
            /// </summary>
            public const string AttrRefreshOnChange = "refreshOnChange";

            /// <summary>
            /// Parameter Prompt String
            /// </summary>
            public const string AttrPrompt = "prompt";

            /// <summary>
            /// Boolean
            /// </summary>
            public const string AttrBoolean = "boolean";

            /// <summary>
            /// Double
            /// </summary>
            public const string AttrDouble = "double";

            /// <summary>
            /// Integer
            /// </summary>
            public const string AttrInteger = "integer";

            /// <summary>
            /// String
            /// </summary>
            public const string AttrString = "string";

            /// <summary>
            /// Cell Reference
            /// </summary>
            public const string AttrCell = "cell";

            /// <summary>
            /// File Type
            /// </summary>
            public const string AttrFileType = "fileType";

            /// <summary>
            /// Code Page
            /// </summary>
            public const string AttrCodePage = "codePage";

            /// <summary>
            /// Delimited File
            /// </summary>
            public const string AttrDelimited = "delimited";

            /// <summary>
            /// Decimal Separator
            /// </summary>
            public const string AttrDecimal = "decimal";

            /// <summary>
            /// Thousands Separator
            /// </summary>
            public const string AttrThousands = "thousands";

            /// <summary>
            /// Tab as Delimiter
            /// </summary>
            public const string AttrTab = "tab";

            /// <summary>
            /// Space is Delimiter
            /// </summary>
            public const string AttrSpace = "space";

            /// <summary>
            /// Comma is Delimiter
            /// </summary>
            public const string AttrComma = "comma";

            /// <summary>
            /// Semicolon is Delimiter
            /// </summary>
            public const string AttrSemicolon = "semicolon";

            /// <summary>
            /// Qualifier
            /// </summary>
            public const string AttrQualifier = "qualifier";

            /// <summary>
            /// Custom Delimiter
            /// </summary>
            public const string AttrDelimiter = "delimiter";

            /// <summary>
            /// Position
            /// </summary>
            public const string AttrPosition = "position";

        }
        #endregion

        #region Names defined in sml-pivotTable.xsd
        public class PivotTable
        {
            /// <summary>
            /// PivotCache Definition
            /// </summary>
            public const string ElPivotCacheDefinition = "pivotCacheDefinition";

            /// <summary>
            /// PivotCache Records
            /// </summary>
            public const string ElPivotCacheRecords = "pivotCacheRecords";

            /// <summary>
            /// PivotTable Definition
            /// </summary>
            public const string ElPivotTableDefinition = "pivotTableDefinition";

            /// <summary>
            /// PivotCache Source Description
            /// </summary>
            public const string ElCacheSource = "cacheSource";

            /// <summary>
            /// PivotCache Fields
            /// </summary>
            public const string ElCacheFields = "cacheFields";

            /// <summary>
            /// PivotCache Hierarchies
            /// </summary>
            public const string ElCacheHierarchies = "cacheHierarchies";

            /// <summary>
            /// OLAP KPIs
            /// </summary>
            public const string ElKpis = "kpis";

            /// <summary>
            /// Tuple Cache
            /// </summary>
            public const string ElTupleCache = "tupleCache";

            /// <summary>
            /// Calculated Items
            /// </summary>
            public const string ElCalculatedItems = "calculatedItems";

            /// <summary>
            /// Calculated Members
            /// </summary>
            public const string ElCalculatedMembers = "calculatedMembers";

            /// <summary>
            /// OLAP Dimensions
            /// </summary>
            public const string ElDimensions = "dimensions";

            /// <summary>
            /// OLAP Measure Groups
            /// </summary>
            public const string ElMeasureGroups = "measureGroups";

            /// <summary>
            /// OLAP Measure Group
            /// </summary>
            public const string ElMaps = "maps";

            /// <summary>
            /// Future Feature Data Storage Area
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// PivotCache Field
            /// </summary>
            public const string ElCacheField = "cacheField";

            /// <summary>
            /// Shared Items
            /// </summary>
            public const string ElSharedItems = "sharedItems";

            /// <summary>
            /// Field Group Properties
            /// </summary>
            public const string ElFieldGroup = "fieldGroup";

            /// <summary>
            /// Member Properties Map
            /// </summary>
            public const string ElMpMap = "mpMap";

            /// <summary>
            /// Worksheet PivotCache Source
            /// </summary>
            public const string ElWorksheetSource = "worksheetSource";

            /// <summary>
            /// Consolidation Source
            /// </summary>
            public const string ElConsolidation = "consolidation";

            /// <summary>
            /// Page Item Values
            /// </summary>
            public const string ElPages = "pages";

            /// <summary>
            /// Range Sets
            /// </summary>
            public const string ElRangeSets = "rangeSets";

            /// <summary>
            /// Page Items
            /// </summary>
            public const string ElPage = "page";

            /// <summary>
            /// Page Item
            /// </summary>
            public const string ElPageItem = "pageItem";

            /// <summary>
            /// Range Set
            /// </summary>
            public const string ElRangeSet = "rangeSet";

            /// <summary>
            /// No Value
            /// </summary>
            public const string ElM = "m";

            /// <summary>
            /// Numeric
            /// </summary>
            public const string ElN = "n";

            /// <summary>
            /// Boolean
            /// </summary>
            public const string ElB = "b";

            /// <summary>
            /// Error Value
            /// </summary>
            public const string ElE = "e";

            /// <summary>
            /// Character Value
            /// </summary>
            public const string ElS = "s";

            /// <summary>
            /// Date Time
            /// </summary>
            public const string ElD = "d";

            /// <summary>
            /// Tuples
            /// </summary>
            public const string ElTpls = "tpls";

            /// <summary>
            /// Member Property Indexes
            /// </summary>
            public const string ElX = "x";

            /// <summary>
            /// Range Grouping Properties
            /// </summary>
            public const string ElRangePr = "rangePr";

            /// <summary>
            /// Discrete Grouping Properties
            /// </summary>
            public const string ElDiscretePr = "discretePr";

            /// <summary>
            /// OLAP Group Items
            /// </summary>
            public const string ElGroupItems = "groupItems";

            /// <summary>
            /// PivotCache Record
            /// </summary>
            public const string ElR = "r";

            /// <summary>
            /// OLAP KPI
            /// </summary>
            public const string ElKpi = "kpi";

            /// <summary>
            /// PivotCache Hierarchy
            /// </summary>
            public const string ElCacheHierarchy = "cacheHierarchy";

            /// <summary>
            /// Fields Usage
            /// </summary>
            public const string ElFieldsUsage = "fieldsUsage";

            /// <summary>
            /// OLAP Grouping Levels
            /// </summary>
            public const string ElGroupLevels = "groupLevels";

            /// <summary>
            /// PivotCache Field Id
            /// </summary>
            public const string ElFieldUsage = "fieldUsage";

            /// <summary>
            /// OLAP Grouping Levels
            /// </summary>
            public const string ElGroupLevel = "groupLevel";

            /// <summary>
            /// OLAP Level Groups
            /// </summary>
            public const string ElGroups = "groups";

            /// <summary>
            /// OLAP Group
            /// </summary>
            public const string ElGroup = "group";

            /// <summary>
            /// OLAP Group Members
            /// </summary>
            public const string ElGroupMembers = "groupMembers";

            /// <summary>
            /// OLAP Group Member
            /// </summary>
            public const string ElGroupMember = "groupMember";

            /// <summary>
            /// Entries
            /// </summary>
            public const string ElEntries = "entries";

            /// <summary>
            /// Sets
            /// </summary>
            public const string ElSets = "sets";

            /// <summary>
            /// OLAP Query Cache
            /// </summary>
            public const string ElQueryCache = "queryCache";

            /// <summary>
            /// Server Formats
            /// </summary>
            public const string ElServerFormats = "serverFormats";

            /// <summary>
            /// Server Format
            /// </summary>
            public const string ElServerFormat = "serverFormat";

            /// <summary>
            /// Tuple
            /// </summary>
            public const string ElTpl = "tpl";

            /// <summary>
            /// OLAP Set
            /// </summary>
            public const string ElSet = "set";

            /// <summary>
            /// Sort By Tuple
            /// </summary>
            public const string ElSortByTuple = "sortByTuple";

            /// <summary>
            /// Query
            /// </summary>
            public const string ElQuery = "query";

            /// <summary>
            /// Calculated Item
            /// </summary>
            public const string ElCalculatedItem = "calculatedItem";

            /// <summary>
            /// Calculated Item Location
            /// </summary>
            public const string ElPivotArea = "pivotArea";

            /// <summary>
            /// Calculated Member
            /// </summary>
            public const string ElCalculatedMember = "calculatedMember";

            /// <summary>
            /// PivotTable Location
            /// </summary>
            public const string ElLocation = "location";

            /// <summary>
            /// PivotTable Fields
            /// </summary>
            public const string ElPivotFields = "pivotFields";

            /// <summary>
            /// Row Fields
            /// </summary>
            public const string ElRowFields = "rowFields";

            /// <summary>
            /// Row Items
            /// </summary>
            public const string ElRowItems = "rowItems";

            /// <summary>
            /// Column Fields
            /// </summary>
            public const string ElColFields = "colFields";

            /// <summary>
            /// Column Items
            /// </summary>
            public const string ElColItems = "colItems";

            /// <summary>
            /// Page Field Items
            /// </summary>
            public const string ElPageFields = "pageFields";

            /// <summary>
            /// Data Fields
            /// </summary>
            public const string ElDataFields = "dataFields";

            /// <summary>
            /// PivotTable Formats
            /// </summary>
            public const string ElFormats = "formats";

            /// <summary>
            /// Conditional Formats
            /// </summary>
            public const string ElConditionalFormats = "conditionalFormats";

            /// <summary>
            /// PivotChart Formats
            /// </summary>
            public const string ElChartFormats = "chartFormats";

            /// <summary>
            /// PivotTable OLAP Hierarchies
            /// </summary>
            public const string ElPivotHierarchies = "pivotHierarchies";

            /// <summary>
            /// PivotTable Style
            /// </summary>
            public const string ElPivotTableStyleInfo = "pivotTableStyleInfo";

            /// <summary>
            /// Filters
            /// </summary>
            public const string ElFilters = "filters";

            /// <summary>
            /// Row OLAP Hierarchy References
            /// </summary>
            public const string ElRowHierarchiesUsage = "rowHierarchiesUsage";

            /// <summary>
            /// Column OLAP Hierarchy References
            /// </summary>
            public const string ElColHierarchiesUsage = "colHierarchiesUsage";

            /// <summary>
            /// PivotTable Field
            /// </summary>
            public const string ElPivotField = "pivotField";

            /// <summary>
            /// Field Items
            /// </summary>
            public const string ElItems = "items";

            /// <summary>
            /// AutoSort Scope
            /// </summary>
            public const string ElAutoSortScope = "autoSortScope";

            /// <summary>
            /// PivotTable Field Item
            /// </summary>
            public const string ElItem = "item";

            /// <summary>
            /// Page Field
            /// </summary>
            public const string ElPageField = "pageField";

            /// <summary>
            /// Data Field Item
            /// </summary>
            public const string ElDataField = "dataField";

            /// <summary>
            /// Row Items
            /// </summary>
            public const string ElI = "i";

            /// <summary>
            /// Row Items
            /// </summary>
            public const string ElField = "field";

            /// <summary>
            /// PivotTable Format
            /// </summary>
            public const string ElFormat = "format";

            /// <summary>
            /// Conditional Formatting
            /// </summary>
            public const string ElConditionalFormat = "conditionalFormat";

            /// <summary>
            /// Pivot Areas
            /// </summary>
            public const string ElPivotAreas = "pivotAreas";

            /// <summary>
            /// PivotChart Format
            /// </summary>
            public const string ElChartFormat = "chartFormat";

            /// <summary>
            /// OLAP Hierarchy
            /// </summary>
            public const string ElPivotHierarchy = "pivotHierarchy";

            /// <summary>
            /// OLAP Member Properties
            /// </summary>
            public const string ElMps = "mps";

            /// <summary>
            /// Members
            /// </summary>
            public const string ElMembers = "members";

            /// <summary>
            /// Row OLAP Hierarchies
            /// </summary>
            public const string ElRowHierarchyUsage = "rowHierarchyUsage";

            /// <summary>
            /// Column OLAP Hierarchies
            /// </summary>
            public const string ElColHierarchyUsage = "colHierarchyUsage";

            /// <summary>
            /// OLAP Member Property
            /// </summary>
            public const string ElMp = "mp";

            /// <summary>
            /// Member
            /// </summary>
            public const string ElMember = "member";

            /// <summary>
            /// OLAP Dimension
            /// </summary>
            public const string ElDimension = "dimension";

            /// <summary>
            /// OLAP Measure Group
            /// </summary>
            public const string ElMeasureGroup = "measureGroup";

            /// <summary>
            /// OLAP Measure Group
            /// </summary>
            public const string ElMap = "map";

            /// <summary>
            /// PivotTable Advanced Filter
            /// </summary>
            public const string ElFilter = "filter";

            /// <summary>
            /// Auto Filter
            /// </summary>
            public const string ElAutoFilter = "autoFilter";

            /// <summary>
            /// Invalid Cache
            /// </summary>
            public const string AttrInvalid = "invalid";

            /// <summary>
            /// Save Pivot Records
            /// </summary>
            public const string AttrSaveData = "saveData";

            /// <summary>
            /// Refresh On Load
            /// </summary>
            public const string AttrRefreshOnLoad = "refreshOnLoad";

            /// <summary>
            /// Optimize Cache for Memory
            /// </summary>
            public const string AttrOptimizeMemory = "optimizeMemory";

            /// <summary>
            /// Enable PivotCache Refresh
            /// </summary>
            public const string AttrEnableRefresh = "enableRefresh";

            /// <summary>
            /// Last Refreshed By
            /// </summary>
            public const string AttrRefreshedBy = "refreshedBy";

            /// <summary>
            /// PivotCache Last Refreshed Date
            /// </summary>
            public const string AttrRefreshedDate = "refreshedDate";

            /// <summary>
            /// Background Query
            /// </summary>
            public const string AttrBackgroundQuery = "backgroundQuery";

            /// <summary>
            /// Missing Items Limit
            /// </summary>
            public const string AttrMissingItemsLimit = "missingItemsLimit";

            /// <summary>
            /// PivotCache Created Version
            /// </summary>
            public const string AttrCreatedVersion = "createdVersion";

            /// <summary>
            /// PivotCache Last Refreshed Version
            /// </summary>
            public const string AttrRefreshedVersion = "refreshedVersion";

            /// <summary>
            /// Minimum Version Required for Refresh
            /// </summary>
            public const string AttrMinRefreshableVersion = "minRefreshableVersion";

            /// <summary>
            /// PivotCache Record Count
            /// </summary>
            public const string AttrRecordCount = "recordCount";

            /// <summary>
            /// Upgrade PivotCache on Refresh
            /// </summary>
            public const string AttrUpgradeOnRefresh = "upgradeOnRefresh";

            /// <summary>
            /// Supports Subqueries
            /// </summary>
            public const string AttrSupportSubquery = "supportSubquery";

            /// <summary>
            /// Supports Attribute Drilldown
            /// </summary>
            public const string AttrSupportAdvancedDrill = "supportAdvancedDrill";

            /// <summary>
            /// Field Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// PivotCache Field Name
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// PivotCache Field Caption
            /// </summary>
            public const string AttrCaption = "caption";

            /// <summary>
            /// Property Name
            /// </summary>
            public const string AttrPropertyName = "propertyName";

            /// <summary>
            /// Server-based Field
            /// </summary>
            public const string AttrServerField = "serverField";

            /// <summary>
            /// Unique List Retrieved
            /// </summary>
            public const string AttrUniqueList = "uniqueList";

            /// <summary>
            /// Number Format Id
            /// </summary>
            public const string AttrNumFmtId = "numFmtId";

            /// <summary>
            /// Calculated Field Formula
            /// </summary>
            public const string AttrFormula = "formula";

            /// <summary>
            /// SQL Data Type
            /// </summary>
            public const string AttrSqlType = "sqlType";

            /// <summary>
            /// Hierarchy
            /// </summary>
            public const string AttrHierarchy = "hierarchy";

            /// <summary>
            /// Hierarchy Level
            /// </summary>
            public const string AttrLevel = "level";

            /// <summary>
            /// Database Field
            /// </summary>
            public const string AttrDatabaseField = "databaseField";

            /// <summary>
            /// Member Property Count
            /// </summary>
            public const string AttrMappingCount = "mappingCount";

            /// <summary>
            /// Member Property Field
            /// </summary>
            public const string AttrMemberPropertyField = "memberPropertyField";

            /// <summary>
            /// Cache Type
            /// </summary>
            public const string AttrType = "type";

            /// <summary>
            /// Connection Index
            /// </summary>
            public const string AttrConnectionId = "connectionId";

            /// <summary>
            /// Reference
            /// </summary>
            public const string AttrRef = "ref";

            /// <summary>
            /// Sheet Name
            /// </summary>
            public const string AttrSheet = "sheet";

            /// <summary>
            /// Auto Page
            /// </summary>
            public const string AttrAutoPage = "autoPage";

            /// <summary>
            /// Field Item Index Page 1
            /// </summary>
            public const string AttrI1 = "i1";

            /// <summary>
            /// Field Item Index Page 2
            /// </summary>
            public const string AttrI2 = "i2";

            /// <summary>
            /// Field Item index Page 3
            /// </summary>
            public const string AttrI3 = "i3";

            /// <summary>
            /// Field Item Index Page 4
            /// </summary>
            public const string AttrI4 = "i4";

            /// <summary>
            /// Contains Semi Mixed Data Types
            /// </summary>
            public const string AttrContainsSemiMixedTypes = "containsSemiMixedTypes";

            /// <summary>
            /// Contains Non Date
            /// </summary>
            public const string AttrContainsNonDate = "containsNonDate";

            /// <summary>
            /// Contains Date
            /// </summary>
            public const string AttrContainsDate = "containsDate";

            /// <summary>
            /// Contains String
            /// </summary>
            public const string AttrContainsString = "containsString";

            /// <summary>
            /// Contains Blank
            /// </summary>
            public const string AttrContainsBlank = "containsBlank";

            /// <summary>
            /// Contains Mixed Data Types
            /// </summary>
            public const string AttrContainsMixedTypes = "containsMixedTypes";

            /// <summary>
            /// Contains Numbers
            /// </summary>
            public const string AttrContainsNumber = "containsNumber";

            /// <summary>
            /// Contains Integer
            /// </summary>
            public const string AttrContainsInteger = "containsInteger";

            /// <summary>
            /// Minimum Numeric Value
            /// </summary>
            public const string AttrMinValue = "minValue";

            /// <summary>
            /// Maximum Numeric Value
            /// </summary>
            public const string AttrMaxValue = "maxValue";

            /// <summary>
            /// Minimum Date Time
            /// </summary>
            public const string AttrMinDate = "minDate";

            /// <summary>
            /// Maximum Date Time Value
            /// </summary>
            public const string AttrMaxDate = "maxDate";

            /// <summary>
            /// Long Text
            /// </summary>
            public const string AttrLongText = "longText";

            /// <summary>
            /// Unused Item
            /// </summary>
            public const string AttrU = "u";

            /// <summary>
            /// Calculated Item
            /// </summary>
            public const string AttrF = "f";

            /// <summary>
            /// Caption
            /// </summary>
            public const string AttrC = "c";

            /// <summary>
            /// Member Property Count
            /// </summary>
            public const string AttrCp = "cp";

            /// <summary>
            /// Format Index
            /// </summary>
            public const string AttrIn = "in";

            /// <summary>
            /// background Color
            /// </summary>
            public const string AttrBc = "bc";

            /// <summary>
            /// Foreground Color
            /// </summary>
            public const string AttrFc = "fc";

            /// <summary>
            /// Underline
            /// </summary>
            public const string AttrUn = "un";

            /// <summary>
            /// Strikethrough
            /// </summary>
            public const string AttrSt = "st";

            /// <summary>
            /// Value
            /// </summary>
            public const string AttrV = "v";

            /// <summary>
            /// Parent
            /// </summary>
            public const string AttrPar = "par";

            /// <summary>
            /// Field Base
            /// </summary>
            public const string AttrBase = "base";

            /// <summary>
            /// Source Data Set Beginning Range
            /// </summary>
            public const string AttrAutoStart = "autoStart";

            /// <summary>
            /// Source Data Ending Range
            /// </summary>
            public const string AttrAutoEnd = "autoEnd";

            /// <summary>
            /// Group By
            /// </summary>
            public const string AttrGroupBy = "groupBy";

            /// <summary>
            /// Numeric Grouping Start Value
            /// </summary>
            public const string AttrStartNum = "startNum";

            /// <summary>
            /// Numeric Grouping End Value
            /// </summary>
            public const string AttrEndNum = "endNum";

            /// <summary>
            /// Date Grouping Start Value
            /// </summary>
            public const string AttrStartDate = "startDate";

            /// <summary>
            /// Date Grouping End Value
            /// </summary>
            public const string AttrEndDate = "endDate";

            /// <summary>
            /// Grouping Interval
            /// </summary>
            public const string AttrGroupInterval = "groupInterval";

            /// <summary>
            /// KPI Unique Name
            /// </summary>
            public const string AttrUniqueName = "uniqueName";

            /// <summary>
            /// KPI Display Folder
            /// </summary>
            public const string AttrDisplayFolder = "displayFolder";

            /// <summary>
            /// Parent KPI
            /// </summary>
            public const string AttrParent = "parent";

            /// <summary>
            /// KPI Value Unique Name
            /// </summary>
            public const string AttrValue = "value";

            /// <summary>
            /// KPI Goal Unique Name
            /// </summary>
            public const string AttrGoal = "goal";

            /// <summary>
            /// KPI Status Unique Name
            /// </summary>
            public const string AttrStatus = "status";

            /// <summary>
            /// KPI Trend Unique Name
            /// </summary>
            public const string AttrTrend = "trend";

            /// <summary>
            /// KPI Weight Unique Name
            /// </summary>
            public const string AttrWeight = "weight";

            /// <summary>
            /// Time Member KPI Unique Name
            /// </summary>
            public const string AttrTime = "time";

            /// <summary>
            /// Measure Hierarchy
            /// </summary>
            public const string AttrMeasure = "measure";

            /// <summary>
            /// Parent Set
            /// </summary>
            public const string AttrParentSet = "parentSet";

            /// <summary>
            /// KPI Icon Set
            /// </summary>
            public const string AttrIconSet = "iconSet";

            /// <summary>
            /// Attribute Hierarchy
            /// </summary>
            public const string AttrAttribute = "attribute";

            /// <summary>
            /// Key Attribute Hierarchy
            /// </summary>
            public const string AttrKeyAttribute = "keyAttribute";

            /// <summary>
            /// Default Member Unique Name
            /// </summary>
            public const string AttrDefaultMemberUniqueName = "defaultMemberUniqueName";

            /// <summary>
            /// Unique Name of 'All'
            /// </summary>
            public const string AttrAllUniqueName = "allUniqueName";

            /// <summary>
            /// Display Name of 'All'
            /// </summary>
            public const string AttrAllCaption = "allCaption";

            /// <summary>
            /// Dimension Unique Name
            /// </summary>
            public const string AttrDimensionUniqueName = "dimensionUniqueName";

            /// <summary>
            /// Measures
            /// </summary>
            public const string AttrMeasures = "measures";

            /// <summary>
            /// One Field
            /// </summary>
            public const string AttrOneField = "oneField";

            /// <summary>
            /// Member Value Data Type
            /// </summary>
            public const string AttrMemberValueDatatype = "memberValueDatatype";

            /// <summary>
            /// Unbalanced
            /// </summary>
            public const string AttrUnbalanced = "unbalanced";

            /// <summary>
            /// Unbalanced Group
            /// </summary>
            public const string AttrUnbalancedGroup = "unbalancedGroup";

            /// <summary>
            /// Hidden
            /// </summary>
            public const string AttrHidden = "hidden";

            /// <summary>
            /// User-Defined Group Level
            /// </summary>
            public const string AttrUser = "user";

            /// <summary>
            /// Custom Roll Up
            /// </summary>
            public const string AttrCustomRollUp = "customRollUp";

            /// <summary>
            /// Parent Unique Name
            /// </summary>
            public const string AttrUniqueParent = "uniqueParent";

            /// <summary>
            /// Group Id
            /// </summary>
            public const string AttrId = "id";

            /// <summary>
            /// Culture
            /// </summary>
            public const string AttrCulture = "culture";

            /// <summary>
            /// Field Index
            /// </summary>
            public const string AttrFld = "fld";

            /// <summary>
            /// Hierarchy Index
            /// </summary>
            public const string AttrHier = "hier";

            /// <summary>
            /// Maximum Rank Requested
            /// </summary>
            public const string AttrMaxRank = "maxRank";

            /// <summary>
            /// MDX Set Definition
            /// </summary>
            public const string AttrSetDefinition = "setDefinition";

            /// <summary>
            /// Set Sort Order
            /// </summary>
            public const string AttrSortType = "sortType";

            /// <summary>
            /// Query Failed
            /// </summary>
            public const string AttrQueryFailed = "queryFailed";

            /// <summary>
            /// MDX Query String
            /// </summary>
            public const string AttrMdx = "mdx";

            /// <summary>
            /// OLAP Calculated Member Name
            /// </summary>
            public const string AttrMemberName = "memberName";

            /// <summary>
            /// Calculated Members Solve Order
            /// </summary>
            public const string AttrSolveOrder = "solveOrder";

            /// <summary>
            /// PivotCache Definition Id
            /// </summary>
            public const string AttrCacheId = "cacheId";

            /// <summary>
            /// Data On Rows
            /// </summary>
            public const string AttrDataOnRows = "dataOnRows";

            /// <summary>
            /// Default Data Field Position
            /// </summary>
            public const string AttrDataPosition = "dataPosition";

            /// <summary>
            /// Data Field Header Name
            /// </summary>
            public const string AttrDataCaption = "dataCaption";

            /// <summary>
            /// Grand Totals Caption
            /// </summary>
            public const string AttrGrandTotalCaption = "grandTotalCaption";

            /// <summary>
            /// Error Caption
            /// </summary>
            public const string AttrErrorCaption = "errorCaption";

            /// <summary>
            /// Show Error
            /// </summary>
            public const string AttrShowError = "showError";

            /// <summary>
            /// Caption for Missing Values
            /// </summary>
            public const string AttrMissingCaption = "missingCaption";

            /// <summary>
            /// Show Missing
            /// </summary>
            public const string AttrShowMissing = "showMissing";

            /// <summary>
            /// Page Header Style Name
            /// </summary>
            public const string AttrPageStyle = "pageStyle";

            /// <summary>
            /// Table Style Name
            /// </summary>
            public const string AttrPivotTableStyle = "pivotTableStyle";

            /// <summary>
            /// Vacated Style
            /// </summary>
            public const string AttrVacatedStyle = "vacatedStyle";

            /// <summary>
            /// PivotTable Custom String
            /// </summary>
            public const string AttrTag = "tag";

            /// <summary>
            /// PivotTable Last Updated Version
            /// </summary>
            public const string AttrUpdatedVersion = "updatedVersion";

            /// <summary>
            /// Asterisk Totals
            /// </summary>
            public const string AttrAsteriskTotals = "asteriskTotals";

            /// <summary>
            /// Show Item Names
            /// </summary>
            public const string AttrShowItems = "showItems";

            /// <summary>
            /// Allow Edit Data
            /// </summary>
            public const string AttrEditData = "editData";

            /// <summary>
            /// Disable Field List
            /// </summary>
            public const string AttrDisableFieldList = "disableFieldList";

            /// <summary>
            /// Show Calculated Members
            /// </summary>
            public const string AttrShowCalcMbrs = "showCalcMbrs";

            /// <summary>
            /// Total Visual Data
            /// </summary>
            public const string AttrVisualTotals = "visualTotals";

            /// <summary>
            /// Show Multiple Labels
            /// </summary>
            public const string AttrShowMultipleLabel = "showMultipleLabel";

            /// <summary>
            /// Show Drop Down
            /// </summary>
            public const string AttrShowDataDropDown = "showDataDropDown";

            /// <summary>
            /// Show Expand Collapse
            /// </summary>
            public const string AttrShowDrill = "showDrill";

            /// <summary>
            /// Print Drill Indicators
            /// </summary>
            public const string AttrPrintDrill = "printDrill";

            /// <summary>
            /// Show Member Property ToolTips
            /// </summary>
            public const string AttrShowMemberPropertyTips = "showMemberPropertyTips";

            /// <summary>
            /// Show ToolTips on Data
            /// </summary>
            public const string AttrShowDataTips = "showDataTips";

            /// <summary>
            /// Enable PivotTable Wizard
            /// </summary>
            public const string AttrEnableWizard = "enableWizard";

            /// <summary>
            /// Enable Drill Down
            /// </summary>
            public const string AttrEnableDrill = "enableDrill";

            /// <summary>
            /// Enable Field Properties
            /// </summary>
            public const string AttrEnableFieldProperties = "enableFieldProperties";

            /// <summary>
            /// Preserve Formatting
            /// </summary>
            public const string AttrPreserveFormatting = "preserveFormatting";

            /// <summary>
            /// Auto Formatting
            /// </summary>
            public const string AttrUseAutoFormatting = "useAutoFormatting";

            /// <summary>
            /// Page Wrap
            /// </summary>
            public const string AttrPageWrap = "pageWrap";

            /// <summary>
            /// Page Over Then Down
            /// </summary>
            public const string AttrPageOverThenDown = "pageOverThenDown";

            /// <summary>
            /// Subtotal Hidden Items
            /// </summary>
            public const string AttrSubtotalHiddenItems = "subtotalHiddenItems";

            /// <summary>
            /// Row Grand Totals
            /// </summary>
            public const string AttrRowGrandTotals = "rowGrandTotals";

            /// <summary>
            /// Grand Totals On Columns
            /// </summary>
            public const string AttrColGrandTotals = "colGrandTotals";

            /// <summary>
            /// Field Print Titles
            /// </summary>
            public const string AttrFieldPrintTitles = "fieldPrintTitles";

            /// <summary>
            /// Item Print Titles
            /// </summary>
            public const string AttrItemPrintTitles = "itemPrintTitles";

            /// <summary>
            /// Merge Titles
            /// </summary>
            public const string AttrMergeItem = "mergeItem";

            /// <summary>
            /// Show Drop Zones
            /// </summary>
            public const string AttrShowDropZones = "showDropZones";

            /// <summary>
            /// Indentation for Compact Axis
            /// </summary>
            public const string AttrIndent = "indent";

            /// <summary>
            /// Show Empty Row
            /// </summary>
            public const string AttrShowEmptyRow = "showEmptyRow";

            /// <summary>
            /// Show Empty Column
            /// </summary>
            public const string AttrShowEmptyCol = "showEmptyCol";

            /// <summary>
            /// Show Field Headers
            /// </summary>
            public const string AttrShowHeaders = "showHeaders";

            /// <summary>
            /// Compact New Fields
            /// </summary>
            public const string AttrCompact = "compact";

            /// <summary>
            /// Outline New Fields
            /// </summary>
            public const string AttrOutline = "outline";

            /// <summary>
            /// Outline Data Fields
            /// </summary>
            public const string AttrOutlineData = "outlineData";

            /// <summary>
            /// Compact Data
            /// </summary>
            public const string AttrCompactData = "compactData";

            /// <summary>
            /// Data Fields Published
            /// </summary>
            public const string AttrPublished = "published";

            /// <summary>
            /// Enable Drop Zones
            /// </summary>
            public const string AttrGridDropZones = "gridDropZones";

            /// <summary>
            /// Stop Immersive UI
            /// </summary>
            public const string AttrImmersive = "immersive";

            /// <summary>
            /// Multiple Field Filters
            /// </summary>
            public const string AttrMultipleFieldFilters = "multipleFieldFilters";

            /// <summary>
            /// Row Header Caption
            /// </summary>
            public const string AttrRowHeaderCaption = "rowHeaderCaption";

            /// <summary>
            /// Column Header Caption
            /// </summary>
            public const string AttrColHeaderCaption = "colHeaderCaption";

            /// <summary>
            /// Default Sort Order
            /// </summary>
            public const string AttrFieldListSortAscending = "fieldListSortAscending";

            /// <summary>
            /// MDX Subqueries Supported
            /// </summary>
            public const string AttrMdxSubqueries = "mdxSubqueries";

            /// <summary>
            /// Custom List AutoSort
            /// </summary>
            public const string AttrCustomListSort = "customListSort";

            /// <summary>
            /// First Header Row
            /// </summary>
            public const string AttrFirstHeaderRow = "firstHeaderRow";

            /// <summary>
            /// PivotTable Data First Row
            /// </summary>
            public const string AttrFirstDataRow = "firstDataRow";

            /// <summary>
            /// First Data Column
            /// </summary>
            public const string AttrFirstDataCol = "firstDataCol";

            /// <summary>
            /// Rows Per Page Count
            /// </summary>
            public const string AttrRowPageCount = "rowPageCount";

            /// <summary>
            /// Columns Per Page
            /// </summary>
            public const string AttrColPageCount = "colPageCount";

            /// <summary>
            /// Axis
            /// </summary>
            public const string AttrAxis = "axis";

            /// <summary>
            /// Custom Subtotal Caption
            /// </summary>
            public const string AttrSubtotalCaption = "subtotalCaption";

            /// <summary>
            /// Show PivotField Header Drop Downs
            /// </summary>
            public const string AttrShowDropDowns = "showDropDowns";

            /// <summary>
            /// Hidden Level
            /// </summary>
            public const string AttrHiddenLevel = "hiddenLevel";

            /// <summary>
            /// Unique Member Property
            /// </summary>
            public const string AttrUniqueMemberProperty = "uniqueMemberProperty";

            /// <summary>
            /// All Items Expanded
            /// </summary>
            public const string AttrAllDrilled = "allDrilled";

            /// <summary>
            /// Subtotals At Top
            /// </summary>
            public const string AttrSubtotalTop = "subtotalTop";

            /// <summary>
            /// Drag To Row
            /// </summary>
            public const string AttrDragToRow = "dragToRow";

            /// <summary>
            /// Drag To Column
            /// </summary>
            public const string AttrDragToCol = "dragToCol";

            /// <summary>
            /// Multiple Field Filters
            /// </summary>
            public const string AttrMultipleItemSelectionAllowed = "multipleItemSelectionAllowed";

            /// <summary>
            /// Drag Field to Page
            /// </summary>
            public const string AttrDragToPage = "dragToPage";

            /// <summary>
            /// Field Can Drag to Data
            /// </summary>
            public const string AttrDragToData = "dragToData";

            /// <summary>
            /// Drag Off
            /// </summary>
            public const string AttrDragOff = "dragOff";

            /// <summary>
            /// Show All Items
            /// </summary>
            public const string AttrShowAll = "showAll";

            /// <summary>
            /// Insert Blank Row
            /// </summary>
            public const string AttrInsertBlankRow = "insertBlankRow";

            /// <summary>
            /// Insert Item Page Break
            /// </summary>
            public const string AttrInsertPageBreak = "insertPageBreak";

            /// <summary>
            /// Auto Show
            /// </summary>
            public const string AttrAutoShow = "autoShow";

            /// <summary>
            /// Top Auto Show
            /// </summary>
            public const string AttrTopAutoShow = "topAutoShow";

            /// <summary>
            /// Hide New Items
            /// </summary>
            public const string AttrHideNewItems = "hideNewItems";

            /// <summary>
            /// Measure Filter
            /// </summary>
            public const string AttrMeasureFilter = "measureFilter";

            /// <summary>
            /// Inclusive Manual Filter
            /// </summary>
            public const string AttrIncludeNewItemsInFilter = "includeNewItemsInFilter";

            /// <summary>
            /// Items Per Page Count
            /// </summary>
            public const string AttrItemPageCount = "itemPageCount";

            /// <summary>
            /// Data Source Sort
            /// </summary>
            public const string AttrDataSourceSort = "dataSourceSort";

            /// <summary>
            /// Auto Sort
            /// </summary>
            public const string AttrNonAutoSortDefault = "nonAutoSortDefault";

            /// <summary>
            /// Auto Show Rank By
            /// </summary>
            public const string AttrRankBy = "rankBy";

            /// <summary>
            /// Show Default Subtotal
            /// </summary>
            public const string AttrDefaultSubtotal = "defaultSubtotal";

            /// <summary>
            /// Sum Subtotal
            /// </summary>
            public const string AttrSumSubtotal = "sumSubtotal";

            /// <summary>
            /// CountA
            /// </summary>
            public const string AttrCountASubtotal = "countASubtotal";

            /// <summary>
            /// Average
            /// </summary>
            public const string AttrAvgSubtotal = "avgSubtotal";

            /// <summary>
            /// Max Subtotal
            /// </summary>
            public const string AttrMaxSubtotal = "maxSubtotal";

            /// <summary>
            /// Min Subtotal
            /// </summary>
            public const string AttrMinSubtotal = "minSubtotal";

            /// <summary>
            /// Product Subtotal
            /// </summary>
            public const string AttrProductSubtotal = "productSubtotal";

            /// <summary>
            /// Count
            /// </summary>
            public const string AttrCountSubtotal = "countSubtotal";

            /// <summary>
            /// StdDev Subtotal
            /// </summary>
            public const string AttrStdDevSubtotal = "stdDevSubtotal";

            /// <summary>
            /// StdDevP Subtotal
            /// </summary>
            public const string AttrStdDevPSubtotal = "stdDevPSubtotal";

            /// <summary>
            /// Variance Subtotal
            /// </summary>
            public const string AttrVarSubtotal = "varSubtotal";

            /// <summary>
            /// VarP Subtotal
            /// </summary>
            public const string AttrVarPSubtotal = "varPSubtotal";

            /// <summary>
            /// Show Member Property in Cell
            /// </summary>
            public const string AttrShowPropCell = "showPropCell";

            /// <summary>
            /// Show Member Property ToolTip
            /// </summary>
            public const string AttrShowPropTip = "showPropTip";

            /// <summary>
            /// Show As Caption
            /// </summary>
            public const string AttrShowPropAsCaption = "showPropAsCaption";

            /// <summary>
            /// Drill State
            /// </summary>
            public const string AttrDefaultAttributeDrillState = "defaultAttributeDrillState";

            /// <summary>
            /// Item Type
            /// </summary>
            public const string AttrT = "t";

            /// <summary>
            /// Hidden
            /// </summary>
            public const string AttrH = "h";

            /// <summary>
            /// Hide Details
            /// </summary>
            public const string AttrSd = "sd";

            /// <summary>
            /// Hierarchy Display Name
            /// </summary>
            public const string AttrCap = "cap";

            /// <summary>
            /// Subtotal
            /// </summary>
            public const string AttrSubtotal = "subtotal";

            /// <summary>
            /// Show Data As Display Format
            /// </summary>
            public const string AttrShowDataAs = "showDataAs";

            /// <summary>
            /// 'Show Data As' Base Field
            /// </summary>
            public const string AttrBaseField = "baseField";

            /// <summary>
            /// 'Show Data As' Base Setting
            /// </summary>
            public const string AttrBaseItem = "baseItem";

            /// <summary>
            /// Format Action
            /// </summary>
            public const string AttrAction = "action";

            /// <summary>
            /// Format Id
            /// </summary>
            public const string AttrDxfId = "dxfId";

            /// <summary>
            /// Conditional Formatting Scope
            /// </summary>
            public const string AttrScope = "scope";

            /// <summary>
            /// Priority
            /// </summary>
            public const string AttrPriority = "priority";

            /// <summary>
            /// Chart Index
            /// </summary>
            public const string AttrChart = "chart";

            /// <summary>
            /// Series Format
            /// </summary>
            public const string AttrSeries = "series";

            /// <summary>
            /// Show In Field List
            /// </summary>
            public const string AttrShowInFieldList = "showInFieldList";

            /// <summary>
            /// Hierarchy Usage
            /// </summary>
            public const string AttrHierarchyUsage = "hierarchyUsage";

            /// <summary>
            /// Show Cell
            /// </summary>
            public const string AttrShowCell = "showCell";

            /// <summary>
            /// Show Tooltip
            /// </summary>
            public const string AttrShowTip = "showTip";

            /// <summary>
            /// Show As Caption
            /// </summary>
            public const string AttrShowAsCaption = "showAsCaption";

            /// <summary>
            /// Name Length
            /// </summary>
            public const string AttrNameLen = "nameLen";

            /// <summary>
            /// Property Name Character Index
            /// </summary>
            public const string AttrPPos = "pPos";

            /// <summary>
            /// Property Name Length
            /// </summary>
            public const string AttrPLen = "pLen";

            /// <summary>
            /// Show Row Header Formatting
            /// </summary>
            public const string AttrShowRowHeaders = "showRowHeaders";

            /// <summary>
            /// Show Table Style Column Header Formatting
            /// </summary>
            public const string AttrShowColHeaders = "showColHeaders";

            /// <summary>
            /// Show Row Stripes
            /// </summary>
            public const string AttrShowRowStripes = "showRowStripes";

            /// <summary>
            /// Show Column Stripes
            /// </summary>
            public const string AttrShowColStripes = "showColStripes";

            /// <summary>
            /// Show Last Column
            /// </summary>
            public const string AttrShowLastColumn = "showLastColumn";

            /// <summary>
            /// Member Property Field Id
            /// </summary>
            public const string AttrMpFld = "mpFld";

            /// <summary>
            /// Evaluation Order
            /// </summary>
            public const string AttrEvalOrder = "evalOrder";

            /// <summary>
            /// Measure Index
            /// </summary>
            public const string AttrIMeasureHier = "iMeasureHier";

            /// <summary>
            /// Measure Field Index
            /// </summary>
            public const string AttrIMeasureFld = "iMeasureFld";

            /// <summary>
            /// Pivot Filter Description
            /// </summary>
            public const string AttrDescription = "description";

            /// <summary>
            /// Label Pivot
            /// </summary>
            public const string AttrStringValue1 = "stringValue1";

            /// <summary>
            /// Label Pivot Filter String Value 2
            /// </summary>
            public const string AttrStringValue2 = "stringValue2";

        }
        #endregion

        #region Names defined in sml-pivotTableShared.xsd
        public class PivotTableShared
        {
            /// <summary>
            /// References
            /// </summary>
            public const string ElReferences = "references";

            /// <summary>
            /// Future Feature Data Storage Area
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Reference
            /// </summary>
            public const string ElReference = "reference";

            /// <summary>
            /// Field Item
            /// </summary>
            public const string ElX = "x";

            /// <summary>
            /// Field Index
            /// </summary>
            public const string AttrField = "field";

            /// <summary>
            /// Rule Type
            /// </summary>
            public const string AttrType = "type";

            /// <summary>
            /// Data Only
            /// </summary>
            public const string AttrDataOnly = "dataOnly";

            /// <summary>
            /// Labels Only
            /// </summary>
            public const string AttrLabelOnly = "labelOnly";

            /// <summary>
            /// Include Row Grand Total
            /// </summary>
            public const string AttrGrandRow = "grandRow";

            /// <summary>
            /// Include Column Grand Total
            /// </summary>
            public const string AttrGrandCol = "grandCol";

            /// <summary>
            /// Cache Index
            /// </summary>
            public const string AttrCacheIndex = "cacheIndex";

            /// <summary>
            /// Outline
            /// </summary>
            public const string AttrOutline = "outline";

            /// <summary>
            /// Offset Reference
            /// </summary>
            public const string AttrOffset = "offset";

            /// <summary>
            /// Collapsed Levels Are Subtotals
            /// </summary>
            public const string AttrCollapsedLevelsAreSubtotals = "collapsedLevelsAreSubtotals";

            /// <summary>
            /// Axis
            /// </summary>
            public const string AttrAxis = "axis";

            /// <summary>
            /// Field Position
            /// </summary>
            public const string AttrFieldPosition = "fieldPosition";

            /// <summary>
            /// Pivot Filter Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// Selected
            /// </summary>
            public const string AttrSelected = "selected";

            /// <summary>
            /// Positional Reference
            /// </summary>
            public const string AttrByPosition = "byPosition";

            /// <summary>
            /// Relative Reference
            /// </summary>
            public const string AttrRelative = "relative";

            /// <summary>
            /// Include Default Filter
            /// </summary>
            public const string AttrDefaultSubtotal = "defaultSubtotal";

            /// <summary>
            /// Include Sum Filter
            /// </summary>
            public const string AttrSumSubtotal = "sumSubtotal";

            /// <summary>
            /// Include CountA Filter
            /// </summary>
            public const string AttrCountASubtotal = "countASubtotal";

            /// <summary>
            /// Include Average Filter
            /// </summary>
            public const string AttrAvgSubtotal = "avgSubtotal";

            /// <summary>
            /// Include Maximum Filter
            /// </summary>
            public const string AttrMaxSubtotal = "maxSubtotal";

            /// <summary>
            /// Include Minimum Filter
            /// </summary>
            public const string AttrMinSubtotal = "minSubtotal";

            /// <summary>
            /// Include Product Filter
            /// </summary>
            public const string AttrProductSubtotal = "productSubtotal";

            /// <summary>
            /// Include Count Subtotal
            /// </summary>
            public const string AttrCountSubtotal = "countSubtotal";

            /// <summary>
            /// Include StdDev Filter
            /// </summary>
            public const string AttrStdDevSubtotal = "stdDevSubtotal";

            /// <summary>
            /// Include StdDevP Filter
            /// </summary>
            public const string AttrStdDevPSubtotal = "stdDevPSubtotal";

            /// <summary>
            /// Include Var Filter
            /// </summary>
            public const string AttrVarSubtotal = "varSubtotal";

            /// <summary>
            /// Include VarP Filter
            /// </summary>
            public const string AttrVarPSubtotal = "varPSubtotal";

            /// <summary>
            /// Shared Items Index
            /// </summary>
            public const string AttrV = "v";

        }
        #endregion

        #region Names defined in sml-queryTable.xsd
        public class QueryTable
        {
            /// <summary>
            /// Query Table
            /// </summary>
            public const string ElQueryTable = "queryTable";

            /// <summary>
            /// QueryTable Refresh Information
            /// </summary>
            public const string ElQueryTableRefresh = "queryTableRefresh";

            /// <summary>
            /// Future Feature Data Storage Area
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Query table fields
            /// </summary>
            public const string ElQueryTableFields = "queryTableFields";

            /// <summary>
            /// Deleted Fields
            /// </summary>
            public const string ElQueryTableDeletedFields = "queryTableDeletedFields";

            /// <summary>
            /// Sort State
            /// </summary>
            public const string ElSortState = "sortState";

            /// <summary>
            /// Deleted Field
            /// </summary>
            public const string ElDeletedField = "deletedField";

            /// <summary>
            /// QueryTable Field
            /// </summary>
            public const string ElQueryTableField = "queryTableField";

            /// <summary>
            /// QueryTable Name
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// First Row Column Titles
            /// </summary>
            public const string AttrHeaders = "headers";

            /// <summary>
            /// Row Numbers
            /// </summary>
            public const string AttrRowNumbers = "rowNumbers";

            /// <summary>
            /// Disable Refresh
            /// </summary>
            public const string AttrDisableRefresh = "disableRefresh";

            /// <summary>
            /// Background Refresh
            /// </summary>
            public const string AttrBackgroundRefresh = "backgroundRefresh";

            /// <summary>
            /// First Background Refresh
            /// </summary>
            public const string AttrFirstBackgroundRefresh = "firstBackgroundRefresh";

            /// <summary>
            /// Refresh On Load
            /// </summary>
            public const string AttrRefreshOnLoad = "refreshOnLoad";

            /// <summary>
            /// Grow Shrink Type
            /// </summary>
            public const string AttrGrowShrinkType = "growShrinkType";

            /// <summary>
            /// Fill Adjacent Formulas
            /// </summary>
            public const string AttrFillFormulas = "fillFormulas";

            /// <summary>
            /// Remove Data On Save
            /// </summary>
            public const string AttrRemoveDataOnSave = "removeDataOnSave";

            /// <summary>
            /// Disable Edit
            /// </summary>
            public const string AttrDisableEdit = "disableEdit";

            /// <summary>
            /// Preserve Formatting On Refresh
            /// </summary>
            public const string AttrPreserveFormatting = "preserveFormatting";

            /// <summary>
            /// Adjust Column Width On Refresh
            /// </summary>
            public const string AttrAdjustColumnWidth = "adjustColumnWidth";

            /// <summary>
            /// Intermediate
            /// </summary>
            public const string AttrIntermediate = "intermediate";

            /// <summary>
            /// Connection Id
            /// </summary>
            public const string AttrConnectionId = "connectionId";

            /// <summary>
            /// Preserve Sort & Filter Layout
            /// </summary>
            public const string AttrPreserveSortFilterLayout = "preserveSortFilterLayout";

            /// <summary>
            /// Next Field Id Wrapped
            /// </summary>
            public const string AttrFieldIdWrapped = "fieldIdWrapped";

            /// <summary>
            /// Headers In Last Refresh
            /// </summary>
            public const string AttrHeadersInLastRefresh = "headersInLastRefresh";

            /// <summary>
            /// Minimum Refresh Version
            /// </summary>
            public const string AttrMinimumVersion = "minimumVersion";

            /// <summary>
            /// Next field id
            /// </summary>
            public const string AttrNextId = "nextId";

            /// <summary>
            /// Columns Left
            /// </summary>
            public const string AttrUnboundColumnsLeft = "unboundColumnsLeft";

            /// <summary>
            /// Columns Right
            /// </summary>
            public const string AttrUnboundColumnsRight = "unboundColumnsRight";

            /// <summary>
            /// Deleted Fields Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// Field Id
            /// </summary>
            public const string AttrId = "id";

            /// <summary>
            /// Data Bound Column
            /// </summary>
            public const string AttrDataBound = "dataBound";

            /// <summary>
            /// Clipped Column
            /// </summary>
            public const string AttrClipped = "clipped";

            /// <summary>
            /// Table Column Id
            /// </summary>
            public const string AttrTableColumnId = "tableColumnId";

        }
        #endregion

        #region Names defined in sml-sharedStringTable.xsd
        public class SharedStringTable
        {
            /// <summary>
            /// Shared String Table
            /// </summary>
            public const string ElSst = "sst";

            /// <summary>
            /// String Item
            /// </summary>
            public const string ElSi = "si";

            /// <summary>
            /// 
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Text
            /// </summary>
            public const string ElT = "t";

            /// <summary>
            /// Run Properties
            /// </summary>
            public const string ElRPr = "rPr";

            /// <summary>
            /// Font
            /// </summary>
            public const string ElRFont = "rFont";

            /// <summary>
            /// Character Set
            /// </summary>
            public const string ElCharset = "charset";

            /// <summary>
            /// Font Family
            /// </summary>
            public const string ElFamily = "family";

            /// <summary>
            /// Bold
            /// </summary>
            public const string ElB = "b";

            /// <summary>
            /// Italic
            /// </summary>
            public const string ElI = "i";

            /// <summary>
            /// Strike Through
            /// </summary>
            public const string ElStrike = "strike";

            /// <summary>
            /// Outline
            /// </summary>
            public const string ElOutline = "outline";

            /// <summary>
            /// Shadow
            /// </summary>
            public const string ElShadow = "shadow";

            /// <summary>
            /// Condense
            /// </summary>
            public const string ElCondense = "condense";

            /// <summary>
            /// Extend
            /// </summary>
            public const string ElExtend = "extend";

            /// <summary>
            /// Text Color
            /// </summary>
            public const string ElColor = "color";

            /// <summary>
            /// Font Size
            /// </summary>
            public const string ElSz = "sz";

            /// <summary>
            /// Underline
            /// </summary>
            public const string ElU = "u";

            /// <summary>
            /// Vertical Alignment
            /// </summary>
            public const string ElVertAlign = "vertAlign";

            /// <summary>
            /// Font Scheme
            /// </summary>
            public const string ElScheme = "scheme";

            /// <summary>
            /// Rich Text Run
            /// </summary>
            public const string ElR = "r";

            /// <summary>
            /// Phonetic Run
            /// </summary>
            public const string ElRPh = "rPh";

            /// <summary>
            /// Phonetic Properties
            /// </summary>
            public const string ElPhoneticPr = "phoneticPr";

            /// <summary>
            /// String Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// Unique String Count
            /// </summary>
            public const string AttrUniqueCount = "uniqueCount";

            /// <summary>
            /// Base Text Start Index
            /// </summary>
            public const string AttrSb = "sb";

            /// <summary>
            /// Base Text End Index
            /// </summary>
            public const string AttrEb = "eb";

            /// <summary>
            /// Font Id
            /// </summary>
            public const string AttrFontId = "fontId";

            /// <summary>
            /// Character Type
            /// </summary>
            public const string AttrType = "type";

            /// <summary>
            /// Alignment
            /// </summary>
            public const string AttrAlignment = "alignment";

        }
        #endregion

        #region Names defined in sml-sharedWorkbookRevisions.xsd
        public class SharedWorkbookRevisions
        {
            /// <summary>
            /// Revision Headers
            /// </summary>
            public const string ElHeaders = "headers";

            /// <summary>
            /// Revisions
            /// </summary>
            public const string ElRevisions = "revisions";

            /// <summary>
            /// Header
            /// </summary>
            public const string ElHeader = "header";

            /// <summary>
            /// Revision Row Column Insert Delete
            /// </summary>
            public const string ElRrc = "rrc";

            /// <summary>
            /// Revision Cell Move
            /// </summary>
            public const string ElRm = "rm";

            /// <summary>
            /// Revision Custom View
            /// </summary>
            public const string ElRcv = "rcv";

            /// <summary>
            /// Revision Sheet Name
            /// </summary>
            public const string ElRsnm = "rsnm";

            /// <summary>
            /// Revision Insert Sheet
            /// </summary>
            public const string ElRis = "ris";

            /// <summary>
            /// Revision Cell Change
            /// </summary>
            public const string ElRcc = "rcc";

            /// <summary>
            /// Revision Format
            /// </summary>
            public const string ElRfmt = "rfmt";

            /// <summary>
            /// Revision AutoFormat
            /// </summary>
            public const string ElRaf = "raf";

            /// <summary>
            /// Revision Defined Name
            /// </summary>
            public const string ElRdn = "rdn";

            /// <summary>
            /// Revision Cell Comment
            /// </summary>
            public const string ElRcmt = "rcmt";

            /// <summary>
            /// Revision Query Table
            /// </summary>
            public const string ElRqt = "rqt";

            /// <summary>
            /// Revision Merge Conflict
            /// </summary>
            public const string ElRcft = "rcft";

            /// <summary>
            /// Sheet Id Map
            /// </summary>
            public const string ElSheetIdMap = "sheetIdMap";

            /// <summary>
            /// Reviewed List
            /// </summary>
            public const string ElReviewedList = "reviewedList";

            /// <summary>
            /// 
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Sheet Id
            /// </summary>
            public const string ElSheetId = "sheetId";

            /// <summary>
            /// Reviewed
            /// </summary>
            public const string ElReviewed = "reviewed";

            /// <summary>
            /// Undo
            /// </summary>
            public const string ElUndo = "undo";

            /// <summary>
            /// Old Cell Data
            /// </summary>
            public const string ElOc = "oc";

            /// <summary>
            /// New Cell Data
            /// </summary>
            public const string ElNc = "nc";

            /// <summary>
            /// Old Formatting Information
            /// </summary>
            public const string ElOdxf = "odxf";

            /// <summary>
            /// New Formatting Information
            /// </summary>
            public const string ElNdxf = "ndxf";

            /// <summary>
            /// Formatting
            /// </summary>
            public const string ElDxf = "dxf";

            /// <summary>
            /// Formula
            /// </summary>
            public const string ElFormula = "formula";

            /// <summary>
            /// Old Formula
            /// </summary>
            public const string ElOldFormula = "oldFormula";

            /// <summary>
            /// Last Revision GUID
            /// </summary>
            public const string AttrGuid = "guid";

            /// <summary>
            /// Last GUID
            /// </summary>
            public const string AttrLastGuid = "lastGuid";

            /// <summary>
            /// Shared Workbook
            /// </summary>
            public const string AttrShared = "shared";

            /// <summary>
            /// Disk Revisions
            /// </summary>
            public const string AttrDiskRevisions = "diskRevisions";

            /// <summary>
            /// History
            /// </summary>
            public const string AttrHistory = "history";

            /// <summary>
            /// Track Revisions
            /// </summary>
            public const string AttrTrackRevisions = "trackRevisions";

            /// <summary>
            /// Exclusive Mode
            /// </summary>
            public const string AttrExclusive = "exclusive";

            /// <summary>
            /// Revision Id
            /// </summary>
            public const string AttrRevisionId = "revisionId";

            /// <summary>
            /// Version
            /// </summary>
            public const string AttrVersion = "version";

            /// <summary>
            /// Keep Change History
            /// </summary>
            public const string AttrKeepChangeHistory = "keepChangeHistory";

            /// <summary>
            /// Protected
            /// </summary>
            public const string AttrProtected = "protected";

            /// <summary>
            /// Preserve History
            /// </summary>
            public const string AttrPreserveHistory = "preserveHistory";

            /// <summary>
            /// Revision Id
            /// </summary>
            public const string AttrRId = "rId";

            /// <summary>
            /// Revision From Rejection
            /// </summary>
            public const string AttrUa = "ua";

            /// <summary>
            /// Revision Undo Rejected
            /// </summary>
            public const string AttrRa = "ra";

            /// <summary>
            /// Date Time
            /// </summary>
            public const string AttrDateTime = "dateTime";

            /// <summary>
            /// Last Sheet Id
            /// </summary>
            public const string AttrMaxSheetId = "maxSheetId";

            /// <summary>
            /// User Name
            /// </summary>
            public const string AttrUserName = "userName";

            /// <summary>
            /// Minimum Revision Id
            /// </summary>
            public const string AttrMinRId = "minRId";

            /// <summary>
            /// Max Revision Id
            /// </summary>
            public const string AttrMaxRId = "maxRId";

            /// <summary>
            /// Sheet Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// Sheet Id
            /// </summary>
            public const string AttrVal = "val";

            /// <summary>
            /// Index
            /// </summary>
            public const string AttrIndex = "index";

            /// <summary>
            /// Expression
            /// </summary>
            public const string AttrExp = "exp";

            /// <summary>
            /// Reference 3D
            /// </summary>
            public const string AttrRef3D = "ref3D";

            /// <summary>
            /// Array Entered
            /// </summary>
            public const string AttrArray = "array";

            /// <summary>
            /// Value Needed
            /// </summary>
            public const string AttrV = "v";

            /// <summary>
            /// Defined Name Formula
            /// </summary>
            public const string AttrNf = "nf";

            /// <summary>
            /// Cross Sheet Move
            /// </summary>
            public const string AttrCs = "cs";

            /// <summary>
            /// Range
            /// </summary>
            public const string AttrDr = "dr";

            /// <summary>
            /// Defined Name
            /// </summary>
            public const string AttrDn = "dn";

            /// <summary>
            /// Cell Reference
            /// </summary>
            public const string AttrR = "r";

            /// <summary>
            /// Sheet Id
            /// </summary>
            public const string AttrSId = "sId";

            /// <summary>
            /// End Of List
            /// </summary>
            public const string AttrEol = "eol";

            /// <summary>
            /// Reference
            /// </summary>
            public const string AttrRef = "ref";

            /// <summary>
            /// User Action
            /// </summary>
            public const string AttrAction = "action";

            /// <summary>
            /// Edge Deleted
            /// </summary>
            public const string AttrEdge = "edge";

            /// <summary>
            /// Source
            /// </summary>
            public const string AttrSource = "source";

            /// <summary>
            /// Destination
            /// </summary>
            public const string AttrDestination = "destination";

            /// <summary>
            /// Source Sheet Id
            /// </summary>
            public const string AttrSourceSheetId = "sourceSheetId";

            /// <summary>
            /// Old Sheet Name
            /// </summary>
            public const string AttrOldName = "oldName";

            /// <summary>
            /// New Sheet Name
            /// </summary>
            public const string AttrNewName = "newName";

            /// <summary>
            /// Sheet Name
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// Sheet Position
            /// </summary>
            public const string AttrSheetPosition = "sheetPosition";

            /// <summary>
            /// Row Column Formatting Change
            /// </summary>
            public const string AttrXfDxf = "xfDxf";

            /// <summary>
            /// Style Revision
            /// </summary>
            public const string AttrS = "s";

            /// <summary>
            /// Number Format Id
            /// </summary>
            public const string AttrNumFmtId = "numFmtId";

            /// <summary>
            /// Quote Prefix
            /// </summary>
            public const string AttrQuotePrefix = "quotePrefix";

            /// <summary>
            /// Old Quote Prefix
            /// </summary>
            public const string AttrOldQuotePrefix = "oldQuotePrefix";

            /// <summary>
            /// Phonetic Text
            /// </summary>
            public const string AttrPh = "ph";

            /// <summary>
            /// Old Phonetic Text
            /// </summary>
            public const string AttrOldPh = "oldPh";

            /// <summary>
            /// End of List  Formula Update
            /// </summary>
            public const string AttrEndOfListFormulaUpdate = "endOfListFormulaUpdate";

            /// <summary>
            /// Sequence Of References
            /// </summary>
            public const string AttrSqref = "sqref";

            /// <summary>
            /// Start index
            /// </summary>
            public const string AttrStart = "start";

            /// <summary>
            /// Length
            /// </summary>
            public const string AttrLength = "length";

            /// <summary>
            /// Cell
            /// </summary>
            public const string AttrCell = "cell";

            /// <summary>
            /// Always Show Comment
            /// </summary>
            public const string AttrAlwaysShow = "alwaysShow";

            /// <summary>
            /// Old Comment
            /// </summary>
            public const string AttrOld = "old";

            /// <summary>
            /// Comment In Hidden Row
            /// </summary>
            public const string AttrHiddenRow = "hiddenRow";

            /// <summary>
            /// Hidden Column
            /// </summary>
            public const string AttrHiddenColumn = "hiddenColumn";

            /// <summary>
            /// Author
            /// </summary>
            public const string AttrAuthor = "author";

            /// <summary>
            /// Original Comment Length
            /// </summary>
            public const string AttrOldLength = "oldLength";

            /// <summary>
            /// New Comment Length
            /// </summary>
            public const string AttrNewLength = "newLength";

            /// <summary>
            /// Local Name Sheet Id
            /// </summary>
            public const string AttrLocalSheetId = "localSheetId";

            /// <summary>
            /// Custom View
            /// </summary>
            public const string AttrCustomView = "customView";

            /// <summary>
            /// Function
            /// </summary>
            public const string AttrFunction = "function";

            /// <summary>
            /// Old Function
            /// </summary>
            public const string AttrOldFunction = "oldFunction";

            /// <summary>
            /// Function Group Id
            /// </summary>
            public const string AttrFunctionGroupId = "functionGroupId";

            /// <summary>
            /// Old Function Group Id
            /// </summary>
            public const string AttrOldFunctionGroupId = "oldFunctionGroupId";

            /// <summary>
            /// Shortcut Key
            /// </summary>
            public const string AttrShortcutKey = "shortcutKey";

            /// <summary>
            /// Old Short Cut Key
            /// </summary>
            public const string AttrOldShortcutKey = "oldShortcutKey";

            /// <summary>
            /// Named Range Hidden
            /// </summary>
            public const string AttrHidden = "hidden";

            /// <summary>
            /// Old Hidden
            /// </summary>
            public const string AttrOldHidden = "oldHidden";

            /// <summary>
            /// New Custom Menu
            /// </summary>
            public const string AttrCustomMenu = "customMenu";

            /// <summary>
            /// Old Custom Menu Text
            /// </summary>
            public const string AttrOldCustomMenu = "oldCustomMenu";

            /// <summary>
            /// Description
            /// </summary>
            public const string AttrDescription = "description";

            /// <summary>
            /// Old Description
            /// </summary>
            public const string AttrOldDescription = "oldDescription";

            /// <summary>
            /// New Help Topic
            /// </summary>
            public const string AttrHelp = "help";

            /// <summary>
            /// Old Help Topic
            /// </summary>
            public const string AttrOldHelp = "oldHelp";

            /// <summary>
            /// Status Bar
            /// </summary>
            public const string AttrStatusBar = "statusBar";

            /// <summary>
            /// Old Status Bar
            /// </summary>
            public const string AttrOldStatusBar = "oldStatusBar";

            /// <summary>
            /// Name Comment
            /// </summary>
            public const string AttrComment = "comment";

            /// <summary>
            /// Old Name Comment
            /// </summary>
            public const string AttrOldComment = "oldComment";

            /// <summary>
            /// Field Id
            /// </summary>
            public const string AttrFieldId = "fieldId";

        }
        #endregion

        #region Names defined in sml-sharedWorkbookUserNames.xsd
        public class SharedWorkbookUserNames
        {
            /// <summary>
            /// User List
            /// </summary>
            public const string ElUsers = "users";

            /// <summary>
            /// User Information
            /// </summary>
            public const string ElUserInfo = "userInfo";

            /// <summary>
            /// 
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Active User Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// User Revisions GUID
            /// </summary>
            public const string AttrGuid = "guid";

            /// <summary>
            /// User Name
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// User Id
            /// </summary>
            public const string AttrId = "id";

            /// <summary>
            /// Date Time
            /// </summary>
            public const string AttrDateTime = "dateTime";

        }
        #endregion

        #region Names defined in sml-sheet.xsd
        public class Sheet
        {
            /// <summary>
            /// Worksheet
            /// </summary>
            public const string ElWorksheet = "worksheet";

            /// <summary>
            /// Chart Sheet
            /// </summary>
            public const string ElChartsheet = "chartsheet";

            /// <summary>
            /// Dialog Sheet
            /// </summary>
            public const string ElDialogsheet = "dialogsheet";

            /// <summary>
            /// Sheet Properties
            /// </summary>
            public const string ElSheetPr = "sheetPr";

            /// <summary>
            /// Macro Sheet Dimensions
            /// </summary>
            public const string ElDimension = "dimension";

            /// <summary>
            /// Macro Sheet Views
            /// </summary>
            public const string ElSheetViews = "sheetViews";

            /// <summary>
            /// Sheet Format Properties
            /// </summary>
            public const string ElSheetFormatPr = "sheetFormatPr";

            /// <summary>
            /// Column Information
            /// </summary>
            public const string ElCols = "cols";

            /// <summary>
            /// Sheet Data
            /// </summary>
            public const string ElSheetData = "sheetData";

            /// <summary>
            /// Sheet Protection Options
            /// </summary>
            public const string ElSheetProtection = "sheetProtection";

            /// <summary>
            /// AutoFilter
            /// </summary>
            public const string ElAutoFilter = "autoFilter";

            /// <summary>
            /// Sort State
            /// </summary>
            public const string ElSortState = "sortState";

            /// <summary>
            /// Data Consolidation
            /// </summary>
            public const string ElDataConsolidate = "dataConsolidate";

            /// <summary>
            /// Custom Sheet Views
            /// </summary>
            public const string ElCustomSheetViews = "customSheetViews";

            /// <summary>
            /// Phonetic Properties
            /// </summary>
            public const string ElPhoneticPr = "phoneticPr";

            /// <summary>
            /// Conditional Formatting
            /// </summary>
            public const string ElConditionalFormatting = "conditionalFormatting";

            /// <summary>
            /// Print Options
            /// </summary>
            public const string ElPrintOptions = "printOptions";

            /// <summary>
            /// Page Margins
            /// </summary>
            public const string ElPageMargins = "pageMargins";

            /// <summary>
            /// Page Setup Settings
            /// </summary>
            public const string ElPageSetup = "pageSetup";

            /// <summary>
            /// Header Footer Settings
            /// </summary>
            public const string ElHeaderFooter = "headerFooter";

            /// <summary>
            /// Horizontal Page Breaks (Row)
            /// </summary>
            public const string ElRowBreaks = "rowBreaks";

            /// <summary>
            /// Vertical Page Breaks
            /// </summary>
            public const string ElColBreaks = "colBreaks";

            /// <summary>
            /// Custom Properties
            /// </summary>
            public const string ElCustomProperties = "customProperties";

            /// <summary>
            /// Drawing
            /// </summary>
            public const string ElDrawing = "drawing";

            /// <summary>
            /// Legacy Drawing Reference
            /// </summary>
            public const string ElLegacyDrawing = "legacyDrawing";

            /// <summary>
            /// Legacy Drawing Header Footer
            /// </summary>
            public const string ElLegacyDrawingHF = "legacyDrawingHF";

            /// <summary>
            /// Background Image
            /// </summary>
            public const string ElPicture = "picture";

            /// <summary>
            /// OLE Objects
            /// </summary>
            public const string ElOleObjects = "oleObjects";

            /// <summary>
            /// Future Feature Data Storage Area
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Sheet Calculation Properties
            /// </summary>
            public const string ElSheetCalcPr = "sheetCalcPr";

            /// <summary>
            /// Protected Ranges
            /// </summary>
            public const string ElProtectedRanges = "protectedRanges";

            /// <summary>
            /// Scenarios
            /// </summary>
            public const string ElScenarios = "scenarios";

            /// <summary>
            /// Merge Cells
            /// </summary>
            public const string ElMergeCells = "mergeCells";

            /// <summary>
            /// Data Validations
            /// </summary>
            public const string ElDataValidations = "dataValidations";

            /// <summary>
            /// Hyperlinks
            /// </summary>
            public const string ElHyperlinks = "hyperlinks";

            /// <summary>
            /// Cell Watch Items
            /// </summary>
            public const string ElCellWatches = "cellWatches";

            /// <summary>
            /// Ignored Errors
            /// </summary>
            public const string ElIgnoredErrors = "ignoredErrors";

            /// <summary>
            /// Smart Tags
            /// </summary>
            public const string ElSmartTags = "smartTags";

            /// <summary>
            /// Embedded Controls
            /// </summary>
            public const string ElControls = "controls";

            /// <summary>
            /// Web Publishing Items
            /// </summary>
            public const string ElWebPublishItems = "webPublishItems";

            /// <summary>
            /// Table Parts
            /// </summary>
            public const string ElTableParts = "tableParts";

            /// <summary>
            /// Row
            /// </summary>
            public const string ElRow = "row";

            /// <summary>
            /// Column Width & Formatting
            /// </summary>
            public const string ElCol = "col";

            /// <summary>
            /// Cell
            /// </summary>
            public const string ElC = "c";

            /// <summary>
            /// Formula
            /// </summary>
            public const string ElF = "f";

            /// <summary>
            /// Cell Value
            /// </summary>
            public const string ElV = "v";

            /// <summary>
            /// Rich Text Inline
            /// </summary>
            public const string ElIs = "is";

            /// <summary>
            /// Sheet Tab Color
            /// </summary>
            public const string ElTabColor = "tabColor";

            /// <summary>
            /// Outline Properties
            /// </summary>
            public const string ElOutlinePr = "outlinePr";

            /// <summary>
            /// Page Setup Properties
            /// </summary>
            public const string ElPageSetUpPr = "pageSetUpPr";

            /// <summary>
            /// Worksheet View
            /// </summary>
            public const string ElSheetView = "sheetView";

            /// <summary>
            /// View Pane
            /// </summary>
            public const string ElPane = "pane";

            /// <summary>
            /// Selection
            /// </summary>
            public const string ElSelection = "selection";

            /// <summary>
            /// PivotTable Selection
            /// </summary>
            public const string ElPivotSelection = "pivotSelection";

            /// <summary>
            /// Pivot Area
            /// </summary>
            public const string ElPivotArea = "pivotArea";

            /// <summary>
            /// Break
            /// </summary>
            public const string ElBrk = "brk";

            /// <summary>
            /// Data Consolidation References
            /// </summary>
            public const string ElDataRefs = "dataRefs";

            /// <summary>
            /// Data Consolidation Reference
            /// </summary>
            public const string ElDataRef = "dataRef";

            /// <summary>
            /// Merged Cell
            /// </summary>
            public const string ElMergeCell = "mergeCell";

            /// <summary>
            /// Cell Smart Tags
            /// </summary>
            public const string ElCellSmartTags = "cellSmartTags";

            /// <summary>
            /// Cell Smart Tag
            /// </summary>
            public const string ElCellSmartTag = "cellSmartTag";

            /// <summary>
            /// Smart Tag Properties
            /// </summary>
            public const string ElCellSmartTagPr = "cellSmartTagPr";

            /// <summary>
            /// Custom Sheet View
            /// </summary>
            public const string ElCustomSheetView = "customSheetView";

            /// <summary>
            /// Data Validation
            /// </summary>
            public const string ElDataValidation = "dataValidation";

            /// <summary>
            /// Formula 1
            /// </summary>
            public const string ElFormula1 = "formula1";

            /// <summary>
            /// Formula 2
            /// </summary>
            public const string ElFormula2 = "formula2";

            /// <summary>
            /// Conditional Formatting Rule
            /// </summary>
            public const string ElCfRule = "cfRule";

            /// <summary>
            /// Formula
            /// </summary>
            public const string ElFormula = "formula";

            /// <summary>
            /// Color Scale
            /// </summary>
            public const string ElColorScale = "colorScale";

            /// <summary>
            /// Data Bar
            /// </summary>
            public const string ElDataBar = "dataBar";

            /// <summary>
            /// Icon Set
            /// </summary>
            public const string ElIconSet = "iconSet";

            /// <summary>
            /// Hyperlink
            /// </summary>
            public const string ElHyperlink = "hyperlink";

            /// <summary>
            /// Conditional Format Value Object
            /// </summary>
            public const string ElCfvo = "cfvo";

            /// <summary>
            /// Color Gradiant Interpolation
            /// </summary>
            public const string ElColor = "color";

            /// <summary>
            /// Odd Header
            /// </summary>
            public const string ElOddHeader = "oddHeader";

            /// <summary>
            /// Odd Page Footer
            /// </summary>
            public const string ElOddFooter = "oddFooter";

            /// <summary>
            /// Even Page Header
            /// </summary>
            public const string ElEvenHeader = "evenHeader";

            /// <summary>
            /// Even Page Footer
            /// </summary>
            public const string ElEvenFooter = "evenFooter";

            /// <summary>
            /// First Page Header
            /// </summary>
            public const string ElFirstHeader = "firstHeader";

            /// <summary>
            /// First Page Footer
            /// </summary>
            public const string ElFirstFooter = "firstFooter";

            /// <summary>
            /// Scenario
            /// </summary>
            public const string ElScenario = "scenario";

            /// <summary>
            /// Protected Range
            /// </summary>
            public const string ElProtectedRange = "protectedRange";

            /// <summary>
            /// Input Cells
            /// </summary>
            public const string ElInputCells = "inputCells";

            /// <summary>
            /// Cell Watch Item
            /// </summary>
            public const string ElCellWatch = "cellWatch";

            /// <summary>
            /// Custom Property
            /// </summary>
            public const string ElCustomPr = "customPr";

            /// <summary>
            /// OLE Object
            /// </summary>
            public const string ElOleObject = "oleObject";

            /// <summary>
            /// Web Publishing Item
            /// </summary>
            public const string ElWebPublishItem = "webPublishItem";

            /// <summary>
            /// Embedded Control
            /// </summary>
            public const string ElControl = "control";

            /// <summary>
            /// Ignored Error
            /// </summary>
            public const string ElIgnoredError = "ignoredError";

            /// <summary>
            /// Table Part
            /// </summary>
            public const string ElTablePart = "tablePart";

            /// <summary>
            /// Full Calculation On Load
            /// </summary>
            public const string AttrFullCalcOnLoad = "fullCalcOnLoad";

            /// <summary>
            /// Base Column Width
            /// </summary>
            public const string AttrBaseColWidth = "baseColWidth";

            /// <summary>
            /// Default Column Width
            /// </summary>
            public const string AttrDefaultColWidth = "defaultColWidth";

            /// <summary>
            /// Default Row Height
            /// </summary>
            public const string AttrDefaultRowHeight = "defaultRowHeight";

            /// <summary>
            /// Custom Height
            /// </summary>
            public const string AttrCustomHeight = "customHeight";

            /// <summary>
            /// Hidden By Default
            /// </summary>
            public const string AttrZeroHeight = "zeroHeight";

            /// <summary>
            /// Thick Top Border
            /// </summary>
            public const string AttrThickTop = "thickTop";

            /// <summary>
            /// Thick Bottom Border
            /// </summary>
            public const string AttrThickBottom = "thickBottom";

            /// <summary>
            /// Maximum Outline Row
            /// </summary>
            public const string AttrOutlineLevelRow = "outlineLevelRow";

            /// <summary>
            /// Column Outline Level
            /// </summary>
            public const string AttrOutlineLevelCol = "outlineLevelCol";

            /// <summary>
            /// Minimum Column
            /// </summary>
            public const string AttrMin = "min";

            /// <summary>
            /// Maximum Column
            /// </summary>
            public const string AttrMax = "max";

            /// <summary>
            /// Column Width
            /// </summary>
            public const string AttrWidth = "width";

            /// <summary>
            /// Style
            /// </summary>
            public const string AttrStyle = "style";

            /// <summary>
            /// Hidden Columns
            /// </summary>
            public const string AttrHidden = "hidden";

            /// <summary>
            /// Best Fit Column Width
            /// </summary>
            public const string AttrBestFit = "bestFit";

            /// <summary>
            /// Custom Width
            /// </summary>
            public const string AttrCustomWidth = "customWidth";

            /// <summary>
            /// Show Phonetic Information
            /// </summary>
            public const string AttrPhonetic = "phonetic";

            /// <summary>
            /// Outline Level
            /// </summary>
            public const string AttrOutlineLevel = "outlineLevel";

            /// <summary>
            /// Collapsed
            /// </summary>
            public const string AttrCollapsed = "collapsed";

            /// <summary>
            /// Row Index
            /// </summary>
            public const string AttrR = "r";

            /// <summary>
            /// Spans
            /// </summary>
            public const string AttrSpans = "spans";

            /// <summary>
            /// Style Index
            /// </summary>
            public const string AttrS = "s";

            /// <summary>
            /// Custom Format
            /// </summary>
            public const string AttrCustomFormat = "customFormat";

            /// <summary>
            /// Row Height
            /// </summary>
            public const string AttrHt = "ht";

            /// <summary>
            /// Thick Bottom
            /// </summary>
            public const string AttrThickBot = "thickBot";

            /// <summary>
            /// Show Phonetic
            /// </summary>
            public const string AttrPh = "ph";

            /// <summary>
            /// Cell Data Type
            /// </summary>
            public const string AttrT = "t";

            /// <summary>
            /// Cell Metadata Index
            /// </summary>
            public const string AttrCm = "cm";

            /// <summary>
            /// Value Metadata Index
            /// </summary>
            public const string AttrVm = "vm";

            /// <summary>
            /// Synch Horizontal
            /// </summary>
            public const string AttrSyncHorizontal = "syncHorizontal";

            /// <summary>
            /// Synch Vertical
            /// </summary>
            public const string AttrSyncVertical = "syncVertical";

            /// <summary>
            /// Synch Reference
            /// </summary>
            public const string AttrSyncRef = "syncRef";

            /// <summary>
            /// Transition Formula Evaluation
            /// </summary>
            public const string AttrTransitionEvaluation = "transitionEvaluation";

            /// <summary>
            /// Transition Formula Entry
            /// </summary>
            public const string AttrTransitionEntry = "transitionEntry";

            /// <summary>
            /// Published
            /// </summary>
            public const string AttrPublished = "published";

            /// <summary>
            /// Code Name
            /// </summary>
            public const string AttrCodeName = "codeName";

            /// <summary>
            /// Filter Mode
            /// </summary>
            public const string AttrFilterMode = "filterMode";

            /// <summary>
            /// Enable Conditional Formatting Calculations
            /// </summary>
            public const string AttrEnableFormatConditionsCalculation = "enableFormatConditionsCalculation";

            /// <summary>
            /// Reference
            /// </summary>
            public const string AttrRef = "ref";

            /// <summary>
            /// Window Protection
            /// </summary>
            public const string AttrWindowProtection = "windowProtection";

            /// <summary>
            /// Show Formulas
            /// </summary>
            public const string AttrShowFormulas = "showFormulas";

            /// <summary>
            /// Show Grid Lines
            /// </summary>
            public const string AttrShowGridLines = "showGridLines";

            /// <summary>
            /// Show Headers
            /// </summary>
            public const string AttrShowRowColHeaders = "showRowColHeaders";

            /// <summary>
            /// Show Zero Values
            /// </summary>
            public const string AttrShowZeros = "showZeros";

            /// <summary>
            /// Right To Left
            /// </summary>
            public const string AttrRightToLeft = "rightToLeft";

            /// <summary>
            /// Sheet Tab Selected
            /// </summary>
            public const string AttrTabSelected = "tabSelected";

            /// <summary>
            /// Show Ruler
            /// </summary>
            public const string AttrShowRuler = "showRuler";

            /// <summary>
            /// Show Outline Symbols
            /// </summary>
            public const string AttrShowOutlineSymbols = "showOutlineSymbols";

            /// <summary>
            /// Default Grid Color
            /// </summary>
            public const string AttrDefaultGridColor = "defaultGridColor";

            /// <summary>
            /// Show White Space
            /// </summary>
            public const string AttrShowWhiteSpace = "showWhiteSpace";

            /// <summary>
            /// View Type
            /// </summary>
            public const string AttrView = "view";

            /// <summary>
            /// Top Left Visible Cell
            /// </summary>
            public const string AttrTopLeftCell = "topLeftCell";

            /// <summary>
            /// Color Id
            /// </summary>
            public const string AttrColorId = "colorId";

            /// <summary>
            /// Zoom Scale
            /// </summary>
            public const string AttrZoomScale = "zoomScale";

            /// <summary>
            /// Zoom Scale Normal View
            /// </summary>
            public const string AttrZoomScaleNormal = "zoomScaleNormal";

            /// <summary>
            /// Zoom Scale Page Break Preview
            /// </summary>
            public const string AttrZoomScaleSheetLayoutView = "zoomScaleSheetLayoutView";

            /// <summary>
            /// Zoom Scale Page Layout View
            /// </summary>
            public const string AttrZoomScalePageLayoutView = "zoomScalePageLayoutView";

            /// <summary>
            /// Workbook View Index
            /// </summary>
            public const string AttrWorkbookViewId = "workbookViewId";

            /// <summary>
            /// Horizontal Split Position
            /// </summary>
            public const string AttrXSplit = "xSplit";

            /// <summary>
            /// Vertical Split Position
            /// </summary>
            public const string AttrYSplit = "ySplit";

            /// <summary>
            /// Active Pane
            /// </summary>
            public const string AttrActivePane = "activePane";

            /// <summary>
            /// Split State
            /// </summary>
            public const string AttrState = "state";

            /// <summary>
            /// Show Header
            /// </summary>
            public const string AttrShowHeader = "showHeader";

            /// <summary>
            /// Label
            /// </summary>
            public const string AttrLabel = "label";

            /// <summary>
            /// Data Selection
            /// </summary>
            public const string AttrData = "data";

            /// <summary>
            /// Extendable
            /// </summary>
            public const string AttrExtendable = "extendable";

            /// <summary>
            /// Selection Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// Axis
            /// </summary>
            public const string AttrAxis = "axis";

            /// <summary>
            /// Start
            /// </summary>
            public const string AttrStart = "start";

            /// <summary>
            /// Active Row
            /// </summary>
            public const string AttrActiveRow = "activeRow";

            /// <summary>
            /// Active Column
            /// </summary>
            public const string AttrActiveCol = "activeCol";

            /// <summary>
            /// Previous Row
            /// </summary>
            public const string AttrPreviousRow = "previousRow";

            /// <summary>
            /// Previous Column Selection
            /// </summary>
            public const string AttrPreviousCol = "previousCol";

            /// <summary>
            /// Click Count
            /// </summary>
            public const string AttrClick = "click";

            /// <summary>
            /// Active Cell Location
            /// </summary>
            public const string AttrActiveCell = "activeCell";

            /// <summary>
            /// Active Cell Index
            /// </summary>
            public const string AttrActiveCellId = "activeCellId";

            /// <summary>
            /// Sequence of References
            /// </summary>
            public const string AttrSqref = "sqref";

            /// <summary>
            /// Manual Break Count
            /// </summary>
            public const string AttrManualBreakCount = "manualBreakCount";

            /// <summary>
            /// Id
            /// </summary>
            public const string AttrId = "id";

            /// <summary>
            /// Manual Page Break
            /// </summary>
            public const string AttrMan = "man";

            /// <summary>
            /// Pivot-Created Page Break
            /// </summary>
            public const string AttrPt = "pt";

            /// <summary>
            /// Apply Styles in Outline
            /// </summary>
            public const string AttrApplyStyles = "applyStyles";

            /// <summary>
            /// Summary Below
            /// </summary>
            public const string AttrSummaryBelow = "summaryBelow";

            /// <summary>
            /// Summary Right
            /// </summary>
            public const string AttrSummaryRight = "summaryRight";

            /// <summary>
            /// Show Auto Page Breaks
            /// </summary>
            public const string AttrAutoPageBreaks = "autoPageBreaks";

            /// <summary>
            /// Fit To Page
            /// </summary>
            public const string AttrFitToPage = "fitToPage";

            /// <summary>
            /// Function Index
            /// </summary>
            public const string AttrFunction = "function";

            /// <summary>
            /// Use Left Column Labels
            /// </summary>
            public const string AttrLeftLabels = "leftLabels";

            /// <summary>
            /// Labels In Top Row
            /// </summary>
            public const string AttrTopLabels = "topLabels";

            /// <summary>
            /// Link
            /// </summary>
            public const string AttrLink = "link";

            /// <summary>
            /// Named Range
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// Sheet Name
            /// </summary>
            public const string AttrSheet = "sheet";

            /// <summary>
            /// Smart Tag Type Index
            /// </summary>
            public const string AttrType = "type";

            /// <summary>
            /// Deleted
            /// </summary>
            public const string AttrDeleted = "deleted";

            /// <summary>
            /// XML Based
            /// </summary>
            public const string AttrXmlBased = "xmlBased";

            /// <summary>
            /// Key Name
            /// </summary>
            public const string AttrKey = "key";

            /// <summary>
            /// Value
            /// </summary>
            public const string AttrVal = "val";

            /// <summary>
            /// GUID
            /// </summary>
            public const string AttrGuid = "guid";

            /// <summary>
            /// Print Scale
            /// </summary>
            public const string AttrScale = "scale";

            /// <summary>
            /// Show Page Breaks
            /// </summary>
            public const string AttrShowPageBreaks = "showPageBreaks";

            /// <summary>
            /// Show Headers
            /// </summary>
            public const string AttrShowRowCol = "showRowCol";

            /// <summary>
            /// Show Outline Symbols
            /// </summary>
            public const string AttrOutlineSymbols = "outlineSymbols";

            /// <summary>
            /// Show Zero Values
            /// </summary>
            public const string AttrZeroValues = "zeroValues";

            /// <summary>
            /// Print Area Defined
            /// </summary>
            public const string AttrPrintArea = "printArea";

            /// <summary>
            /// Filtered List
            /// </summary>
            public const string AttrFilter = "filter";

            /// <summary>
            /// Show AutoFitler Drop Down Controls
            /// </summary>
            public const string AttrShowAutoFilter = "showAutoFilter";

            /// <summary>
            /// Hidden Rows
            /// </summary>
            public const string AttrHiddenRows = "hiddenRows";

            /// <summary>
            /// Hidden Columns
            /// </summary>
            public const string AttrHiddenColumns = "hiddenColumns";

            /// <summary>
            /// Filter
            /// </summary>
            public const string AttrFilterUnique = "filterUnique";

            /// <summary>
            /// Disable Prompts
            /// </summary>
            public const string AttrDisablePrompts = "disablePrompts";

            /// <summary>
            /// Top Left Corner (X Coodrinate)
            /// </summary>
            public const string AttrXWindow = "xWindow";

            /// <summary>
            /// Top Left Corner (Y Coordinate)
            /// </summary>
            public const string AttrYWindow = "yWindow";

            /// <summary>
            /// Data Validation Error Style
            /// </summary>
            public const string AttrErrorStyle = "errorStyle";

            /// <summary>
            /// IME Mode Enforced
            /// </summary>
            public const string AttrImeMode = "imeMode";

            /// <summary>
            /// Operator
            /// </summary>
            public const string AttrOperator = "operator";

            /// <summary>
            /// Allow Blank
            /// </summary>
            public const string AttrAllowBlank = "allowBlank";

            /// <summary>
            /// Show Drop Down
            /// </summary>
            public const string AttrShowDropDown = "showDropDown";

            /// <summary>
            /// Show Input Message
            /// </summary>
            public const string AttrShowInputMessage = "showInputMessage";

            /// <summary>
            /// Show Error Message
            /// </summary>
            public const string AttrShowErrorMessage = "showErrorMessage";

            /// <summary>
            /// Error Alert Text
            /// </summary>
            public const string AttrErrorTitle = "errorTitle";

            /// <summary>
            /// Error Message
            /// </summary>
            public const string AttrError = "error";

            /// <summary>
            /// Prompt Title
            /// </summary>
            public const string AttrPromptTitle = "promptTitle";

            /// <summary>
            /// Input Prompt
            /// </summary>
            public const string AttrPrompt = "prompt";

            /// <summary>
            /// PivotTable Conditional Formatting
            /// </summary>
            public const string AttrPivot = "pivot";

            /// <summary>
            /// Differential Formatting Id
            /// </summary>
            public const string AttrDxfId = "dxfId";

            /// <summary>
            /// Priority
            /// </summary>
            public const string AttrPriority = "priority";

            /// <summary>
            /// Stop If True
            /// </summary>
            public const string AttrStopIfTrue = "stopIfTrue";

            /// <summary>
            /// Above Or Below Average
            /// </summary>
            public const string AttrAboveAverage = "aboveAverage";

            /// <summary>
            /// Top 10 Percent
            /// </summary>
            public const string AttrPercent = "percent";

            /// <summary>
            /// Bottom N
            /// </summary>
            public const string AttrBottom = "bottom";

            /// <summary>
            /// Text
            /// </summary>
            public const string AttrText = "text";

            /// <summary>
            /// Time Period
            /// </summary>
            public const string AttrTimePeriod = "timePeriod";

            /// <summary>
            /// Rank
            /// </summary>
            public const string AttrRank = "rank";

            /// <summary>
            /// StdDev
            /// </summary>
            public const string AttrStdDev = "stdDev";

            /// <summary>
            /// Equal Average
            /// </summary>
            public const string AttrEqualAverage = "equalAverage";

            /// <summary>
            /// Location
            /// </summary>
            public const string AttrLocation = "location";

            /// <summary>
            /// Tool Tip
            /// </summary>
            public const string AttrTooltip = "tooltip";

            /// <summary>
            /// Display String
            /// </summary>
            public const string AttrDisplay = "display";

            /// <summary>
            /// Always Calculate Array
            /// </summary>
            public const string AttrAca = "aca";

            /// <summary>
            /// Data Table 2-D
            /// </summary>
            public const string AttrDt2D = "dt2D";

            /// <summary>
            /// Data Table Row
            /// </summary>
            public const string AttrDtr = "dtr";

            /// <summary>
            /// Input 1 Deleted
            /// </summary>
            public const string AttrDel1 = "del1";

            /// <summary>
            /// Input 2 Deleted
            /// </summary>
            public const string AttrDel2 = "del2";

            /// <summary>
            /// Data Table Cell 1
            /// </summary>
            public const string AttrR1 = "r1";

            /// <summary>
            /// Input Cell 2
            /// </summary>
            public const string AttrR2 = "r2";

            /// <summary>
            /// Calculate Cell
            /// </summary>
            public const string AttrCa = "ca";

            /// <summary>
            /// Shared Group Index
            /// </summary>
            public const string AttrSi = "si";

            /// <summary>
            /// Assigns Value to Name
            /// </summary>
            public const string AttrBx = "bx";

            /// <summary>
            /// Minimum Length
            /// </summary>
            public const string AttrMinLength = "minLength";

            /// <summary>
            /// Maximum Length
            /// </summary>
            public const string AttrMaxLength = "maxLength";

            /// <summary>
            /// Show Values
            /// </summary>
            public const string AttrShowValue = "showValue";

            /// <summary>
            /// Reverse Icons
            /// </summary>
            public const string AttrReverse = "reverse";

            /// <summary>
            /// Greater Than Or Equal
            /// </summary>
            public const string AttrGte = "gte";

            /// <summary>
            /// Left Page Margin
            /// </summary>
            public const string AttrLeft = "left";

            /// <summary>
            /// Right Page Margin
            /// </summary>
            public const string AttrRight = "right";

            /// <summary>
            /// Top Page Margin
            /// </summary>
            public const string AttrTop = "top";

            /// <summary>
            /// Header Page Margin
            /// </summary>
            public const string AttrHeader = "header";

            /// <summary>
            /// Footer Page Margin
            /// </summary>
            public const string AttrFooter = "footer";

            /// <summary>
            /// Horizontal Centered
            /// </summary>
            public const string AttrHorizontalCentered = "horizontalCentered";

            /// <summary>
            /// Vertical Centered
            /// </summary>
            public const string AttrVerticalCentered = "verticalCentered";

            /// <summary>
            /// Print Headings
            /// </summary>
            public const string AttrHeadings = "headings";

            /// <summary>
            /// Print Grid Lines
            /// </summary>
            public const string AttrGridLines = "gridLines";

            /// <summary>
            /// Grid Lines Set
            /// </summary>
            public const string AttrGridLinesSet = "gridLinesSet";

            /// <summary>
            /// Paper Size
            /// </summary>
            public const string AttrPaperSize = "paperSize";

            /// <summary>
            /// First Page Number
            /// </summary>
            public const string AttrFirstPageNumber = "firstPageNumber";

            /// <summary>
            /// Fit To Width
            /// </summary>
            public const string AttrFitToWidth = "fitToWidth";

            /// <summary>
            /// Fit To Height
            /// </summary>
            public const string AttrFitToHeight = "fitToHeight";

            /// <summary>
            /// Page Order
            /// </summary>
            public const string AttrPageOrder = "pageOrder";

            /// <summary>
            /// Orientation
            /// </summary>
            public const string AttrOrientation = "orientation";

            /// <summary>
            /// Use Printer Defaults
            /// </summary>
            public const string AttrUsePrinterDefaults = "usePrinterDefaults";

            /// <summary>
            /// Black And White
            /// </summary>
            public const string AttrBlackAndWhite = "blackAndWhite";

            /// <summary>
            /// Draft
            /// </summary>
            public const string AttrDraft = "draft";

            /// <summary>
            /// Print Cell Comments
            /// </summary>
            public const string AttrCellComments = "cellComments";

            /// <summary>
            /// Use First Page Number
            /// </summary>
            public const string AttrUseFirstPageNumber = "useFirstPageNumber";

            /// <summary>
            /// Print Error Handling
            /// </summary>
            public const string AttrErrors = "errors";

            /// <summary>
            /// Horizontal DPI
            /// </summary>
            public const string AttrHorizontalDpi = "horizontalDpi";

            /// <summary>
            /// Vertical DPI
            /// </summary>
            public const string AttrVerticalDpi = "verticalDpi";

            /// <summary>
            /// Number Of Copies
            /// </summary>
            public const string AttrCopies = "copies";

            /// <summary>
            /// Different Odd Even Header Footer
            /// </summary>
            public const string AttrDifferentOddEven = "differentOddEven";

            /// <summary>
            /// Different First Page
            /// </summary>
            public const string AttrDifferentFirst = "differentFirst";

            /// <summary>
            /// Scale Header & Footer With Document
            /// </summary>
            public const string AttrScaleWithDoc = "scaleWithDoc";

            /// <summary>
            /// Align Margins
            /// </summary>
            public const string AttrAlignWithMargins = "alignWithMargins";

            /// <summary>
            /// Current Scenario
            /// </summary>
            public const string AttrCurrent = "current";

            /// <summary>
            /// Last Shown Scenario
            /// </summary>
            public const string AttrShow = "show";

            /// <summary>
            /// Password
            /// </summary>
            public const string AttrPassword = "password";

            /// <summary>
            /// Objects Locked
            /// </summary>
            public const string AttrObjects = "objects";

            /// <summary>
            /// Format Cells Locked
            /// </summary>
            public const string AttrFormatCells = "formatCells";

            /// <summary>
            /// Format Columns Locked
            /// </summary>
            public const string AttrFormatColumns = "formatColumns";

            /// <summary>
            /// Format Rows Locked
            /// </summary>
            public const string AttrFormatRows = "formatRows";

            /// <summary>
            /// Insert Columns Locked
            /// </summary>
            public const string AttrInsertColumns = "insertColumns";

            /// <summary>
            /// Insert Rows Locked
            /// </summary>
            public const string AttrInsertRows = "insertRows";

            /// <summary>
            /// Insert Hyperlinks Locked
            /// </summary>
            public const string AttrInsertHyperlinks = "insertHyperlinks";

            /// <summary>
            /// Delete Columns Locked
            /// </summary>
            public const string AttrDeleteColumns = "deleteColumns";

            /// <summary>
            /// Delete Rows Locked
            /// </summary>
            public const string AttrDeleteRows = "deleteRows";

            /// <summary>
            /// Select Locked Cells Locked
            /// </summary>
            public const string AttrSelectLockedCells = "selectLockedCells";

            /// <summary>
            /// Sort Locked
            /// </summary>
            public const string AttrSort = "sort";

            /// <summary>
            /// Pivot Tables Locked
            /// </summary>
            public const string AttrPivotTables = "pivotTables";

            /// <summary>
            /// Select Unlocked Cells Locked
            /// </summary>
            public const string AttrSelectUnlockedCells = "selectUnlockedCells";

            /// <summary>
            /// Security Descriptor
            /// </summary>
            public const string AttrSecurityDescriptor = "securityDescriptor";

            /// <summary>
            /// Scenario Locked
            /// </summary>
            public const string AttrLocked = "locked";

            /// <summary>
            /// User Name
            /// </summary>
            public const string AttrUser = "user";

            /// <summary>
            /// Scenario Comment
            /// </summary>
            public const string AttrComment = "comment";

            /// <summary>
            /// Undone
            /// </summary>
            public const string AttrUndone = "undone";

            /// <summary>
            /// Number Format Id
            /// </summary>
            public const string AttrNumFmtId = "numFmtId";

            /// <summary>
            /// Zoom To Fit
            /// </summary>
            public const string AttrZoomToFit = "zoomToFit";

            /// <summary>
            /// Contents
            /// </summary>
            public const string AttrContent = "content";

            /// <summary>
            /// OLE ProgId
            /// </summary>
            public const string AttrProgId = "progId";

            /// <summary>
            /// Data or View Aspect
            /// </summary>
            public const string AttrDvAspect = "dvAspect";

            /// <summary>
            /// OLE Update
            /// </summary>
            public const string AttrOleUpdate = "oleUpdate";

            /// <summary>
            /// Auto Load
            /// </summary>
            public const string AttrAutoLoad = "autoLoad";

            /// <summary>
            /// Shape Id
            /// </summary>
            public const string AttrShapeId = "shapeId";

            /// <summary>
            /// Destination Bookmark
            /// </summary>
            public const string AttrDivId = "divId";

            /// <summary>
            /// Web Source Type
            /// </summary>
            public const string AttrSourceType = "sourceType";

            /// <summary>
            /// Source Id
            /// </summary>
            public const string AttrSourceRef = "sourceRef";

            /// <summary>
            /// Source Object Name
            /// </summary>
            public const string AttrSourceObject = "sourceObject";

            /// <summary>
            /// Destination File Name
            /// </summary>
            public const string AttrDestinationFile = "destinationFile";

            /// <summary>
            /// Title
            /// </summary>
            public const string AttrTitle = "title";

            /// <summary>
            /// Automatically Publish
            /// </summary>
            public const string AttrAutoRepublish = "autoRepublish";

            /// <summary>
            /// Evaluation Error
            /// </summary>
            public const string AttrEvalError = "evalError";

            /// <summary>
            /// Two Digit Text Year
            /// </summary>
            public const string AttrTwoDigitTextYear = "twoDigitTextYear";

            /// <summary>
            /// Number Stored As Text
            /// </summary>
            public const string AttrNumberStoredAsText = "numberStoredAsText";

            /// <summary>
            /// Formula Range
            /// </summary>
            public const string AttrFormulaRange = "formulaRange";

            /// <summary>
            /// Unlocked Formula
            /// </summary>
            public const string AttrUnlockedFormula = "unlockedFormula";

            /// <summary>
            /// Empty Cell Reference
            /// </summary>
            public const string AttrEmptyCellReference = "emptyCellReference";

            /// <summary>
            /// List Data Validation
            /// </summary>
            public const string AttrListDataValidation = "listDataValidation";

            /// <summary>
            /// Calculated Column
            /// </summary>
            public const string AttrCalculatedColumn = "calculatedColumn";

        }
        #endregion

        #region Names defined in sml-sheetMetadata.xsd
        public class SheetMetadata
        {
            /// <summary>
            /// Metadata
            /// </summary>
            public const string ElMetadata = "metadata";

            /// <summary>
            /// Metadata Types Collection
            /// </summary>
            public const string ElMetadataTypes = "metadataTypes";

            /// <summary>
            /// Metadata String Store
            /// </summary>
            public const string ElMetadataStrings = "metadataStrings";

            /// <summary>
            /// MDX Metadata Information
            /// </summary>
            public const string ElMdxMetadata = "mdxMetadata";

            /// <summary>
            /// Future Metadata
            /// </summary>
            public const string ElFutureMetadata = "futureMetadata";

            /// <summary>
            /// Cell Metadata
            /// </summary>
            public const string ElCellMetadata = "cellMetadata";

            /// <summary>
            /// Value Metadata
            /// </summary>
            public const string ElValueMetadata = "valueMetadata";

            /// <summary>
            /// Future Feature Storage Area
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Metadata Type Information
            /// </summary>
            public const string ElMetadataType = "metadataType";

            /// <summary>
            /// Metadata Block
            /// </summary>
            public const string ElBk = "bk";

            /// <summary>
            /// Metadata Record
            /// </summary>
            public const string ElRc = "rc";

            /// <summary>
            /// MDX Metadata Record
            /// </summary>
            public const string ElMdx = "mdx";

            /// <summary>
            /// Tuple MDX Metadata
            /// </summary>
            public const string ElT = "t";

            /// <summary>
            /// Set MDX Metadata
            /// </summary>
            public const string ElMs = "ms";

            /// <summary>
            /// Member Property MDX Metadata
            /// </summary>
            public const string ElP = "p";

            /// <summary>
            /// KPI MDX Metadata
            /// </summary>
            public const string ElK = "k";

            /// <summary>
            /// Member Unique Name Index
            /// </summary>
            public const string ElN = "n";

            /// <summary>
            /// MDX Metadata String
            /// </summary>
            public const string ElS = "s";

            /// <summary>
            /// Metadata Type Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// Metadata Type Name
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// Minimum Supported Version
            /// </summary>
            public const string AttrMinSupportedVersion = "minSupportedVersion";

            /// <summary>
            /// Metadata Ghost Row
            /// </summary>
            public const string AttrGhostRow = "ghostRow";

            /// <summary>
            /// Metadata Ghost Column
            /// </summary>
            public const string AttrGhostCol = "ghostCol";

            /// <summary>
            /// Metadata Edit
            /// </summary>
            public const string AttrEdit = "edit";

            /// <summary>
            /// Metadata Cell Value Delete
            /// </summary>
            public const string AttrDelete = "delete";

            /// <summary>
            /// Metadata Copy
            /// </summary>
            public const string AttrCopy = "copy";

            /// <summary>
            /// Metadata Paste All
            /// </summary>
            public const string AttrPasteAll = "pasteAll";

            /// <summary>
            /// Metadata Paste Formulas
            /// </summary>
            public const string AttrPasteFormulas = "pasteFormulas";

            /// <summary>
            /// Metadata Paste Special Values
            /// </summary>
            public const string AttrPasteValues = "pasteValues";

            /// <summary>
            /// Metadata Paste Formats
            /// </summary>
            public const string AttrPasteFormats = "pasteFormats";

            /// <summary>
            /// Metadata Paste Comments
            /// </summary>
            public const string AttrPasteComments = "pasteComments";

            /// <summary>
            /// Metadata Paste Data Validation
            /// </summary>
            public const string AttrPasteDataValidation = "pasteDataValidation";

            /// <summary>
            /// Metadata Paste Borders
            /// </summary>
            public const string AttrPasteBorders = "pasteBorders";

            /// <summary>
            /// Metadata Paste Column Widths
            /// </summary>
            public const string AttrPasteColWidths = "pasteColWidths";

            /// <summary>
            /// Metadata Paste Number Formats
            /// </summary>
            public const string AttrPasteNumberFormats = "pasteNumberFormats";

            /// <summary>
            /// Metadata Merge
            /// </summary>
            public const string AttrMerge = "merge";

            /// <summary>
            /// Meatadata Split First
            /// </summary>
            public const string AttrSplitFirst = "splitFirst";

            /// <summary>
            /// Metadata Split All
            /// </summary>
            public const string AttrSplitAll = "splitAll";

            /// <summary>
            /// Metadata Insert Delete
            /// </summary>
            public const string AttrRowColShift = "rowColShift";

            /// <summary>
            /// Metadata Clear All
            /// </summary>
            public const string AttrClearAll = "clearAll";

            /// <summary>
            /// Metadata Clear Formats
            /// </summary>
            public const string AttrClearFormats = "clearFormats";

            /// <summary>
            /// Metadata Clear Contents
            /// </summary>
            public const string AttrClearContents = "clearContents";

            /// <summary>
            /// Metadata Clear Comments
            /// </summary>
            public const string AttrClearComments = "clearComments";

            /// <summary>
            /// Metadata Formula Assignment
            /// </summary>
            public const string AttrAssign = "assign";

            /// <summary>
            /// Metadata Coercion
            /// </summary>
            public const string AttrCoerce = "coerce";

            /// <summary>
            /// Adjust Metadata
            /// </summary>
            public const string AttrAdjust = "adjust";

            /// <summary>
            /// Cell Metadata
            /// </summary>
            public const string AttrCellMeta = "cellMeta";

            /// <summary>
            /// Metadata Record Value Index
            /// </summary>
            public const string AttrV = "v";

            /// <summary>
            /// Cube Function Tag
            /// </summary>
            public const string AttrF = "f";

            /// <summary>
            /// Member Index Count
            /// </summary>
            public const string AttrC = "c";

            /// <summary>
            /// Server Formatting Culture Currency
            /// </summary>
            public const string AttrCt = "ct";

            /// <summary>
            /// Server Formatting String Index
            /// </summary>
            public const string AttrSi = "si";

            /// <summary>
            /// Server Formatting Built-In Number Format Index
            /// </summary>
            public const string AttrFi = "fi";

            /// <summary>
            /// Server Formatting Background Color
            /// </summary>
            public const string AttrBc = "bc";

            /// <summary>
            /// Server Formatting Foreground Color
            /// </summary>
            public const string AttrFc = "fc";

            /// <summary>
            /// Server Formatting Italic Font
            /// </summary>
            public const string AttrI = "i";

            /// <summary>
            /// Server Formatting Underline Font
            /// </summary>
            public const string AttrU = "u";

            /// <summary>
            /// Server Formatting Strikethrough Font
            /// </summary>
            public const string AttrSt = "st";

            /// <summary>
            /// Server Formatting Bold Font
            /// </summary>
            public const string AttrB = "b";

            /// <summary>
            /// Set Definition Index
            /// </summary>
            public const string AttrNs = "ns";

            /// <summary>
            /// Set Sort Order
            /// </summary>
            public const string AttrO = "o";

            /// <summary>
            /// Property Name Index
            /// </summary>
            public const string AttrNp = "np";

            /// <summary>
            /// Index Value
            /// </summary>
            public const string AttrX = "x";

        }
        #endregion

        #region Names defined in sml-singleCellTable.xsd
        public class SingleCellTable
        {
            /// <summary>
            /// Single Cells
            /// </summary>
            public const string ElSingleXmlCells = "singleXmlCells";

            /// <summary>
            /// Table Properties
            /// </summary>
            public const string ElSingleXmlCell = "singleXmlCell";

            /// <summary>
            /// Cell Properties
            /// </summary>
            public const string ElXmlCellPr = "xmlCellPr";

            /// <summary>
            /// Future Feature Data Storage Area
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Column XML Properties
            /// </summary>
            public const string ElXmlPr = "xmlPr";

            /// <summary>
            /// Table Id
            /// </summary>
            public const string AttrId = "id";

            /// <summary>
            /// Reference
            /// </summary>
            public const string AttrR = "r";

            /// <summary>
            /// Connection ID
            /// </summary>
            public const string AttrConnectionId = "connectionId";

            /// <summary>
            /// Unique Table Name
            /// </summary>
            public const string AttrUniqueName = "uniqueName";

            /// <summary>
            /// XML Map Id
            /// </summary>
            public const string AttrMapId = "mapId";

            /// <summary>
            /// XPath
            /// </summary>
            public const string AttrXpath = "xpath";

            /// <summary>
            /// XML Data Type
            /// </summary>
            public const string AttrXmlDataType = "xmlDataType";

        }
        #endregion

        #region Names defined in sml-styles.xsd
        public class Styles
        {
            /// <summary>
            /// Style Sheet
            /// </summary>
            public const string ElStyleSheet = "styleSheet";

            /// <summary>
            /// Number Formats
            /// </summary>
            public const string ElNumFmts = "numFmts";

            /// <summary>
            /// Fonts
            /// </summary>
            public const string ElFonts = "fonts";

            /// <summary>
            /// Fills
            /// </summary>
            public const string ElFills = "fills";

            /// <summary>
            /// Borders
            /// </summary>
            public const string ElBorders = "borders";

            /// <summary>
            /// Formatting Records
            /// </summary>
            public const string ElCellStyleXfs = "cellStyleXfs";

            /// <summary>
            /// Cell Formats
            /// </summary>
            public const string ElCellXfs = "cellXfs";

            /// <summary>
            /// Cell Styles
            /// </summary>
            public const string ElCellStyles = "cellStyles";

            /// <summary>
            /// Formats
            /// </summary>
            public const string ElDxfs = "dxfs";

            /// <summary>
            /// Table Styles
            /// </summary>
            public const string ElTableStyles = "tableStyles";

            /// <summary>
            /// Colors
            /// </summary>
            public const string ElColors = "colors";

            /// <summary>
            /// Future Feature Data Storage Area
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Border
            /// </summary>
            public const string ElBorder = "border";

            /// <summary>
            /// Left Border
            /// </summary>
            public const string ElLeft = "left";

            /// <summary>
            /// Right Border
            /// </summary>
            public const string ElRight = "right";

            /// <summary>
            /// Top Border
            /// </summary>
            public const string ElTop = "top";

            /// <summary>
            /// Bottom Border
            /// </summary>
            public const string ElBottom = "bottom";

            /// <summary>
            /// Diagonal
            /// </summary>
            public const string ElDiagonal = "diagonal";

            /// <summary>
            /// Vertical Inner Border
            /// </summary>
            public const string ElVertical = "vertical";

            /// <summary>
            /// Horizontal Inner Borders
            /// </summary>
            public const string ElHorizontal = "horizontal";

            /// <summary>
            /// Color
            /// </summary>
            public const string ElColor = "color";

            /// <summary>
            /// Font
            /// </summary>
            public const string ElFont = "font";

            /// <summary>
            /// Fill
            /// </summary>
            public const string ElFill = "fill";

            /// <summary>
            /// Pattern
            /// </summary>
            public const string ElPatternFill = "patternFill";

            /// <summary>
            /// Gradient
            /// </summary>
            public const string ElGradientFill = "gradientFill";

            /// <summary>
            /// Foreground Color
            /// </summary>
            public const string ElFgColor = "fgColor";

            /// <summary>
            /// Background Color
            /// </summary>
            public const string ElBgColor = "bgColor";

            /// <summary>
            /// Gradient Stop
            /// </summary>
            public const string ElStop = "stop";

            /// <summary>
            /// Number Formats
            /// </summary>
            public const string ElNumFmt = "numFmt";

            /// <summary>
            /// Formatting Elements
            /// </summary>
            public const string ElXf = "xf";

            /// <summary>
            /// Alignment
            /// </summary>
            public const string ElAlignment = "alignment";

            /// <summary>
            /// Protection
            /// </summary>
            public const string ElProtection = "protection";

            /// <summary>
            /// Cell Style
            /// </summary>
            public const string ElCellStyle = "cellStyle";

            /// <summary>
            /// Formatting
            /// </summary>
            public const string ElDxf = "dxf";

            /// <summary>
            /// Color Indexes
            /// </summary>
            public const string ElIndexedColors = "indexedColors";

            /// <summary>
            /// MRU Colors
            /// </summary>
            public const string ElMruColors = "mruColors";

            /// <summary>
            /// RGB Color
            /// </summary>
            public const string ElRgbColor = "rgbColor";

            /// <summary>
            /// Table Style
            /// </summary>
            public const string ElTableStyle = "tableStyle";

            /// <summary>
            /// Table Style
            /// </summary>
            public const string ElTableStyleElement = "tableStyleElement";

            /// <summary>
            /// Font Name
            /// </summary>
            public const string ElName = "name";

            /// <summary>
            /// Character Set
            /// </summary>
            public const string ElCharset = "charset";

            /// <summary>
            /// Font Family
            /// </summary>
            public const string ElFamily = "family";

            /// <summary>
            /// Bold
            /// </summary>
            public const string ElB = "b";

            /// <summary>
            /// Italic
            /// </summary>
            public const string ElI = "i";

            /// <summary>
            /// Strike Through
            /// </summary>
            public const string ElStrike = "strike";

            /// <summary>
            /// Outline
            /// </summary>
            public const string ElOutline = "outline";

            /// <summary>
            /// Shadow
            /// </summary>
            public const string ElShadow = "shadow";

            /// <summary>
            /// Condense
            /// </summary>
            public const string ElCondense = "condense";

            /// <summary>
            /// Extend
            /// </summary>
            public const string ElExtend = "extend";

            /// <summary>
            /// Font Size
            /// </summary>
            public const string ElSz = "sz";

            /// <summary>
            /// Underline
            /// </summary>
            public const string ElU = "u";

            /// <summary>
            /// Text Vertical Alignment
            /// </summary>
            public const string ElVertAlign = "vertAlign";

            /// <summary>
            /// Scheme
            /// </summary>
            public const string ElScheme = "scheme";

            /// <summary>
            /// Text Rotation
            /// </summary>
            public const string AttrTextRotation = "textRotation";

            /// <summary>
            /// Wrap Text
            /// </summary>
            public const string AttrWrapText = "wrapText";

            /// <summary>
            /// Indent
            /// </summary>
            public const string AttrIndent = "indent";

            /// <summary>
            /// Relative Indent
            /// </summary>
            public const string AttrRelativeIndent = "relativeIndent";

            /// <summary>
            /// Justify Last Line
            /// </summary>
            public const string AttrJustifyLastLine = "justifyLastLine";

            /// <summary>
            /// Shrink To Fit
            /// </summary>
            public const string AttrShrinkToFit = "shrinkToFit";

            /// <summary>
            /// Reading Order
            /// </summary>
            public const string AttrReadingOrder = "readingOrder";

            /// <summary>
            /// Border Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// Diagonal Up
            /// </summary>
            public const string AttrDiagonalUp = "diagonalUp";

            /// <summary>
            /// Diagonal Down
            /// </summary>
            public const string AttrDiagonalDown = "diagonalDown";

            /// <summary>
            /// Line Style
            /// </summary>
            public const string AttrStyle = "style";

            /// <summary>
            /// Cell Locked
            /// </summary>
            public const string AttrLocked = "locked";

            /// <summary>
            /// Hidden Cell
            /// </summary>
            public const string AttrHidden = "hidden";

            /// <summary>
            /// Pattern Type
            /// </summary>
            public const string AttrPatternType = "patternType";

            /// <summary>
            /// Automatic
            /// </summary>
            public const string AttrAuto = "auto";

            /// <summary>
            /// Index
            /// </summary>
            public const string AttrIndexed = "indexed";

            /// <summary>
            /// Alpha Red Green Blue Color Value
            /// </summary>
            public const string AttrRgb = "rgb";

            /// <summary>
            /// Theme Color
            /// </summary>
            public const string AttrTheme = "theme";

            /// <summary>
            /// Tint
            /// </summary>
            public const string AttrTint = "tint";

            /// <summary>
            /// Gradient Fill Type
            /// </summary>
            public const string AttrType = "type";

            /// <summary>
            /// Linear Gradient Degree
            /// </summary>
            public const string AttrDegree = "degree";

            /// <summary>
            /// Gradient Stop Position
            /// </summary>
            public const string AttrPosition = "position";

            /// <summary>
            /// Number Format Id
            /// </summary>
            public const string AttrNumFmtId = "numFmtId";

            /// <summary>
            /// Number Format Code
            /// </summary>
            public const string AttrFormatCode = "formatCode";

            /// <summary>
            /// Font Id
            /// </summary>
            public const string AttrFontId = "fontId";

            /// <summary>
            /// Fill Id
            /// </summary>
            public const string AttrFillId = "fillId";

            /// <summary>
            /// Border Id
            /// </summary>
            public const string AttrBorderId = "borderId";

            /// <summary>
            /// Format Id
            /// </summary>
            public const string AttrXfId = "xfId";

            /// <summary>
            /// Quote Prefix
            /// </summary>
            public const string AttrQuotePrefix = "quotePrefix";

            /// <summary>
            /// Pivot Button
            /// </summary>
            public const string AttrPivotButton = "pivotButton";

            /// <summary>
            /// Apply Number Format
            /// </summary>
            public const string AttrApplyNumberFormat = "applyNumberFormat";

            /// <summary>
            /// Apply Font
            /// </summary>
            public const string AttrApplyFont = "applyFont";

            /// <summary>
            /// Apply Fill
            /// </summary>
            public const string AttrApplyFill = "applyFill";

            /// <summary>
            /// Apply Border
            /// </summary>
            public const string AttrApplyBorder = "applyBorder";

            /// <summary>
            /// Apply Alignment
            /// </summary>
            public const string AttrApplyAlignment = "applyAlignment";

            /// <summary>
            /// Apply Protection
            /// </summary>
            public const string AttrApplyProtection = "applyProtection";

            /// <summary>
            /// Built-In Style Id
            /// </summary>
            public const string AttrBuiltinId = "builtinId";

            /// <summary>
            /// Outline Style
            /// </summary>
            public const string AttrILevel = "iLevel";

            /// <summary>
            /// Custom Built In
            /// </summary>
            public const string AttrCustomBuiltin = "customBuiltin";

            /// <summary>
            /// Default Table Style
            /// </summary>
            public const string AttrDefaultTableStyle = "defaultTableStyle";

            /// <summary>
            /// Default Pivot Style
            /// </summary>
            public const string AttrDefaultPivotStyle = "defaultPivotStyle";

            /// <summary>
            /// Pivot Style
            /// </summary>
            public const string AttrPivot = "pivot";

            /// <summary>
            /// Table
            /// </summary>
            public const string AttrTable = "table";

            /// <summary>
            /// Band Size
            /// </summary>
            public const string AttrSize = "size";

            /// <summary>
            /// Formatting Id
            /// </summary>
            public const string AttrDxfId = "dxfId";

            /// <summary>
            /// Value
            /// </summary>
            public const string AttrVal = "val";

            /// <summary>
            /// Auto Format Id
            /// </summary>
            public const string AttrAutoFormatId = "autoFormatId";

            /// <summary>
            /// Apply Number Formats
            /// </summary>
            public const string AttrApplyNumberFormats = "applyNumberFormats";

            /// <summary>
            /// Apply Border Formats
            /// </summary>
            public const string AttrApplyBorderFormats = "applyBorderFormats";

            /// <summary>
            /// Apply Font Formats
            /// </summary>
            public const string AttrApplyFontFormats = "applyFontFormats";

            /// <summary>
            /// Apply Pattern Formats
            /// </summary>
            public const string AttrApplyPatternFormats = "applyPatternFormats";

            /// <summary>
            /// Apply Alignment Formats
            /// </summary>
            public const string AttrApplyAlignmentFormats = "applyAlignmentFormats";

            /// <summary>
            /// Apply Width / Height Formats
            /// </summary>
            public const string AttrApplyWidthHeightFormats = "applyWidthHeightFormats";

        }
        #endregion

        #region Names defined in sml-supplementaryWorkbooks.xsd
        public class SupplementaryWorkbooks
        {
            /// <summary>
            /// External Reference
            /// </summary>
            public const string ElExternalLink = "externalLink";

            /// <summary>
            /// External Workbook
            /// </summary>
            public const string ElExternalBook = "externalBook";

            /// <summary>
            /// DDE Connection
            /// </summary>
            public const string ElDdeLink = "ddeLink";

            /// <summary>
            /// OLE Link
            /// </summary>
            public const string ElOleLink = "oleLink";

            /// <summary>
            /// 
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Supporting Workbook Sheet Names
            /// </summary>
            public const string ElSheetNames = "sheetNames";

            /// <summary>
            /// Named Links
            /// </summary>
            public const string ElDefinedNames = "definedNames";

            /// <summary>
            /// Cached Worksheet Data
            /// </summary>
            public const string ElSheetDataSet = "sheetDataSet";

            /// <summary>
            /// Sheet Name
            /// </summary>
            public const string ElSheetName = "sheetName";

            /// <summary>
            /// Defined Name
            /// </summary>
            public const string ElDefinedName = "definedName";

            /// <summary>
            /// External Sheet Data Set
            /// </summary>
            public const string ElSheetData = "sheetData";

            /// <summary>
            /// Row
            /// </summary>
            public const string ElRow = "row";

            /// <summary>
            /// External Cell Data
            /// </summary>
            public const string ElCell = "cell";

            /// <summary>
            /// Value
            /// </summary>
            public const string ElV = "v";

            /// <summary>
            /// DDE Items Collection
            /// </summary>
            public const string ElDdeItems = "ddeItems";

            /// <summary>
            /// DDE Item definition
            /// </summary>
            public const string ElDdeItem = "ddeItem";

            /// <summary>
            /// DDE Name Values
            /// </summary>
            public const string ElValues = "values";

            /// <summary>
            /// Value
            /// </summary>
            public const string ElValue = "value";

            /// <summary>
            /// DDE Link Value
            /// </summary>
            public const string ElVal = "val";

            /// <summary>
            /// OLE Link Items
            /// </summary>
            public const string ElOleItems = "oleItems";

            /// <summary>
            /// OLE Link Item
            /// </summary>
            public const string ElOleItem = "oleItem";

            /// <summary>
            /// Defined Name
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// Refers To
            /// </summary>
            public const string AttrRefersTo = "refersTo";

            /// <summary>
            /// Sheet Id
            /// </summary>
            public const string AttrSheetId = "sheetId";

            /// <summary>
            /// Last Refresh Resulted in Error
            /// </summary>
            public const string AttrRefreshError = "refreshError";

            /// <summary>
            /// Row
            /// </summary>
            public const string AttrR = "r";

            /// <summary>
            /// Type
            /// </summary>
            public const string AttrT = "t";

            /// <summary>
            /// Value Metadata
            /// </summary>
            public const string AttrVm = "vm";

            /// <summary>
            /// Service name
            /// </summary>
            public const string AttrDdeService = "ddeService";

            /// <summary>
            /// Topic for DDE server
            /// </summary>
            public const string AttrDdeTopic = "ddeTopic";

            /// <summary>
            /// OLE
            /// </summary>
            public const string AttrOle = "ole";

            /// <summary>
            /// Advise
            /// </summary>
            public const string AttrAdvise = "advise";

            /// <summary>
            /// Data is an Image
            /// </summary>
            public const string AttrPreferPic = "preferPic";

            /// <summary>
            /// Rows
            /// </summary>
            public const string AttrRows = "rows";

            /// <summary>
            /// Columns
            /// </summary>
            public const string AttrCols = "cols";

            /// <summary>
            /// OLE Link ProgID
            /// </summary>
            public const string AttrProgId = "progId";

            /// <summary>
            /// Icon
            /// </summary>
            public const string AttrIcon = "icon";

        }
        #endregion

        #region Names defined in sml-table.xsd
        public class Table
        {
            /// <summary>
            /// Table
            /// </summary>
            public const string ElTable = "table";

            /// <summary>
            /// Table AutoFilter
            /// </summary>
            public const string ElAutoFilter = "autoFilter";

            /// <summary>
            /// Sort State
            /// </summary>
            public const string ElSortState = "sortState";

            /// <summary>
            /// Table Columns
            /// </summary>
            public const string ElTableColumns = "tableColumns";

            /// <summary>
            /// Table Style
            /// </summary>
            public const string ElTableStyleInfo = "tableStyleInfo";

            /// <summary>
            /// Future Feature Data Storage Area
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Table Column
            /// </summary>
            public const string ElTableColumn = "tableColumn";

            /// <summary>
            /// Calculated Column Formula
            /// </summary>
            public const string ElCalculatedColumnFormula = "calculatedColumnFormula";

            /// <summary>
            /// Totals Row Formula
            /// </summary>
            public const string ElTotalsRowFormula = "totalsRowFormula";

            /// <summary>
            /// XML Column Properties
            /// </summary>
            public const string ElXmlColumnPr = "xmlColumnPr";

            /// <summary>
            /// Table Id
            /// </summary>
            public const string AttrId = "id";

            /// <summary>
            /// Name
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// Table Name
            /// </summary>
            public const string AttrDisplayName = "displayName";

            /// <summary>
            /// Table Comment
            /// </summary>
            public const string AttrComment = "comment";

            /// <summary>
            /// Reference
            /// </summary>
            public const string AttrRef = "ref";

            /// <summary>
            /// Table Type
            /// </summary>
            public const string AttrTableType = "tableType";

            /// <summary>
            /// Header Row Count
            /// </summary>
            public const string AttrHeaderRowCount = "headerRowCount";

            /// <summary>
            /// Insert Row Showing
            /// </summary>
            public const string AttrInsertRow = "insertRow";

            /// <summary>
            /// Insert Row Shift
            /// </summary>
            public const string AttrInsertRowShift = "insertRowShift";

            /// <summary>
            /// Totals Row Count
            /// </summary>
            public const string AttrTotalsRowCount = "totalsRowCount";

            /// <summary>
            /// Totals Row Shown
            /// </summary>
            public const string AttrTotalsRowShown = "totalsRowShown";

            /// <summary>
            /// Published
            /// </summary>
            public const string AttrPublished = "published";

            /// <summary>
            /// Header Row Format Id
            /// </summary>
            public const string AttrHeaderRowDxfId = "headerRowDxfId";

            /// <summary>
            /// Data Area Format Id
            /// </summary>
            public const string AttrDataDxfId = "dataDxfId";

            /// <summary>
            /// Totals Row Format Id
            /// </summary>
            public const string AttrTotalsRowDxfId = "totalsRowDxfId";

            /// <summary>
            /// Header Row Border Format Id
            /// </summary>
            public const string AttrHeaderRowBorderDxfId = "headerRowBorderDxfId";

            /// <summary>
            /// Table Border Format Id
            /// </summary>
            public const string AttrTableBorderDxfId = "tableBorderDxfId";

            /// <summary>
            /// Totals Row Border Format Id
            /// </summary>
            public const string AttrTotalsRowBorderDxfId = "totalsRowBorderDxfId";

            /// <summary>
            /// Header Row Style
            /// </summary>
            public const string AttrHeaderRowCellStyle = "headerRowCellStyle";

            /// <summary>
            /// Data Style Name
            /// </summary>
            public const string AttrDataCellStyle = "dataCellStyle";

            /// <summary>
            /// Totals Row Style
            /// </summary>
            public const string AttrTotalsRowCellStyle = "totalsRowCellStyle";

            /// <summary>
            /// Connection ID
            /// </summary>
            public const string AttrConnectionId = "connectionId";

            /// <summary>
            /// Show First Column
            /// </summary>
            public const string AttrShowFirstColumn = "showFirstColumn";

            /// <summary>
            /// Show Last Column
            /// </summary>
            public const string AttrShowLastColumn = "showLastColumn";

            /// <summary>
            /// Show Row Stripes
            /// </summary>
            public const string AttrShowRowStripes = "showRowStripes";

            /// <summary>
            /// Show Column Stripes
            /// </summary>
            public const string AttrShowColumnStripes = "showColumnStripes";

            /// <summary>
            /// Column Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// Unique Name
            /// </summary>
            public const string AttrUniqueName = "uniqueName";

            /// <summary>
            /// Totals Row Function
            /// </summary>
            public const string AttrTotalsRowFunction = "totalsRowFunction";

            /// <summary>
            /// Totals Row Label
            /// </summary>
            public const string AttrTotalsRowLabel = "totalsRowLabel";

            /// <summary>
            /// Query Table Field Id
            /// </summary>
            public const string AttrQueryTableFieldId = "queryTableFieldId";

            /// <summary>
            /// Array
            /// </summary>
            public const string AttrArray = "array";

            /// <summary>
            /// XML Map Id
            /// </summary>
            public const string AttrMapId = "mapId";

            /// <summary>
            /// XPath
            /// </summary>
            public const string AttrXpath = "xpath";

            /// <summary>
            /// Denormalized
            /// </summary>
            public const string AttrDenormalized = "denormalized";

            /// <summary>
            /// XML Data Type
            /// </summary>
            public const string AttrXmlDataType = "xmlDataType";

        }
        #endregion

        #region Names defined in sml-volatileDependencies.xsd
        public class VolatileDependencies
        {
            /// <summary>
            /// Volatile Dependency Types
            /// </summary>
            public const string ElVolTypes = "volTypes";

            /// <summary>
            /// Volatile Dependency Type
            /// </summary>
            public const string ElVolType = "volType";

            /// <summary>
            /// 
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Main
            /// </summary>
            public const string ElMain = "main";

            /// <summary>
            /// Topic
            /// </summary>
            public const string ElTp = "tp";

            /// <summary>
            /// Topic Value
            /// </summary>
            public const string ElV = "v";

            /// <summary>
            /// Strings in Subtopic
            /// </summary>
            public const string ElStp = "stp";

            /// <summary>
            /// References
            /// </summary>
            public const string ElTr = "tr";

            /// <summary>
            /// Type
            /// </summary>
            public const string AttrType = "type";

            /// <summary>
            /// First String
            /// </summary>
            public const string AttrFirst = "first";

            /// <summary>
            /// Type
            /// </summary>
            public const string AttrT = "t";

            /// <summary>
            /// Reference
            /// </summary>
            public const string AttrR = "r";

            /// <summary>
            /// Sheet Id
            /// </summary>
            public const string AttrS = "s";

        }
        #endregion

        #region Names defined in sml-workbook.xsd
        public class Workbook
        {
            /// <summary>
            /// Workbook
            /// </summary>
            public const string ElWorkbook = "workbook";

            /// <summary>
            /// File Version
            /// </summary>
            public const string ElFileVersion = "fileVersion";

            /// <summary>
            /// File Sharing
            /// </summary>
            public const string ElFileSharing = "fileSharing";

            /// <summary>
            /// Workbook Properties
            /// </summary>
            public const string ElWorkbookPr = "workbookPr";

            /// <summary>
            /// Workbook Protection
            /// </summary>
            public const string ElWorkbookProtection = "workbookProtection";

            /// <summary>
            /// Workbook Views
            /// </summary>
            public const string ElBookViews = "bookViews";

            /// <summary>
            /// Sheets
            /// </summary>
            public const string ElSheets = "sheets";

            /// <summary>
            /// Function Groups
            /// </summary>
            public const string ElFunctionGroups = "functionGroups";

            /// <summary>
            /// External References
            /// </summary>
            public const string ElExternalReferences = "externalReferences";

            /// <summary>
            /// Defined Names
            /// </summary>
            public const string ElDefinedNames = "definedNames";

            /// <summary>
            /// Calculation Properties
            /// </summary>
            public const string ElCalcPr = "calcPr";

            /// <summary>
            /// OLE Size
            /// </summary>
            public const string ElOleSize = "oleSize";

            /// <summary>
            /// Custom Workbook Views
            /// </summary>
            public const string ElCustomWorkbookViews = "customWorkbookViews";

            /// <summary>
            /// PivotCaches
            /// </summary>
            public const string ElPivotCaches = "pivotCaches";

            /// <summary>
            /// Smart Tag Properties
            /// </summary>
            public const string ElSmartTagPr = "smartTagPr";

            /// <summary>
            /// Smart Tag Types
            /// </summary>
            public const string ElSmartTagTypes = "smartTagTypes";

            /// <summary>
            /// Web Publishing Properties
            /// </summary>
            public const string ElWebPublishing = "webPublishing";

            /// <summary>
            /// File Recovery Properties
            /// </summary>
            public const string ElFileRecoveryPr = "fileRecoveryPr";

            /// <summary>
            /// Web Publish Objects
            /// </summary>
            public const string ElWebPublishObjects = "webPublishObjects";

            /// <summary>
            /// Future Feature Data Storage Area
            /// </summary>
            public const string ElExtLst = "extLst";

            /// <summary>
            /// Workbook View
            /// </summary>
            public const string ElWorkbookView = "workbookView";

            /// <summary>
            /// Custom Workbook View
            /// </summary>
            public const string ElCustomWorkbookView = "customWorkbookView";

            /// <summary>
            /// Sheet Information
            /// </summary>
            public const string ElSheet = "sheet";

            /// <summary>
            /// Smart Tag Type
            /// </summary>
            public const string ElSmartTagType = "smartTagType";

            /// <summary>
            /// Defined Name
            /// </summary>
            public const string ElDefinedName = "definedName";

            /// <summary>
            /// External Reference
            /// </summary>
            public const string ElExternalReference = "externalReference";

            /// <summary>
            /// PivotCache
            /// </summary>
            public const string ElPivotCache = "pivotCache";

            /// <summary>
            /// Function Group
            /// </summary>
            public const string ElFunctionGroup = "functionGroup";

            /// <summary>
            /// Web Publishing Object
            /// </summary>
            public const string ElWebPublishObject = "webPublishObject";

            /// <summary>
            /// Application Name
            /// </summary>
            public const string AttrAppName = "appName";

            /// <summary>
            /// Last Edited Version
            /// </summary>
            public const string AttrLastEdited = "lastEdited";

            /// <summary>
            /// Lowest Edited Version
            /// </summary>
            public const string AttrLowestEdited = "lowestEdited";

            /// <summary>
            /// Build Version
            /// </summary>
            public const string AttrRupBuild = "rupBuild";

            /// <summary>
            /// Code Name
            /// </summary>
            public const string AttrCodeName = "codeName";

            /// <summary>
            /// Visibility
            /// </summary>
            public const string AttrVisibility = "visibility";

            /// <summary>
            /// Minimized
            /// </summary>
            public const string AttrMinimized = "minimized";

            /// <summary>
            /// Show Horizontal Scroll
            /// </summary>
            public const string AttrShowHorizontalScroll = "showHorizontalScroll";

            /// <summary>
            /// Show Vertical Scroll
            /// </summary>
            public const string AttrShowVerticalScroll = "showVerticalScroll";

            /// <summary>
            /// Show Sheet Tabs
            /// </summary>
            public const string AttrShowSheetTabs = "showSheetTabs";

            /// <summary>
            /// Upper Left Corner (X Coordinate)
            /// </summary>
            public const string AttrXWindow = "xWindow";

            /// <summary>
            /// Upper Left Corner (Y Coordinate)
            /// </summary>
            public const string AttrYWindow = "yWindow";

            /// <summary>
            /// Window Width
            /// </summary>
            public const string AttrWindowWidth = "windowWidth";

            /// <summary>
            /// Window Height
            /// </summary>
            public const string AttrWindowHeight = "windowHeight";

            /// <summary>
            /// Sheet Tab Ratio
            /// </summary>
            public const string AttrTabRatio = "tabRatio";

            /// <summary>
            /// First Sheet
            /// </summary>
            public const string AttrFirstSheet = "firstSheet";

            /// <summary>
            /// Active Sheet Index
            /// </summary>
            public const string AttrActiveTab = "activeTab";

            /// <summary>
            /// AutoFilter Date Grouping
            /// </summary>
            public const string AttrAutoFilterDateGrouping = "autoFilterDateGrouping";

            /// <summary>
            /// Custom View Name
            /// </summary>
            public const string AttrName = "name";

            /// <summary>
            /// Custom View GUID
            /// </summary>
            public const string AttrGuid = "guid";

            /// <summary>
            /// Auto Update
            /// </summary>
            public const string AttrAutoUpdate = "autoUpdate";

            /// <summary>
            /// Merge Interval
            /// </summary>
            public const string AttrMergeInterval = "mergeInterval";

            /// <summary>
            /// Changes Saved Win
            /// </summary>
            public const string AttrChangesSavedWin = "changesSavedWin";

            /// <summary>
            /// Only Synch
            /// </summary>
            public const string AttrOnlySync = "onlySync";

            /// <summary>
            /// Personal View
            /// </summary>
            public const string AttrPersonalView = "personalView";

            /// <summary>
            /// Include Print Settings
            /// </summary>
            public const string AttrIncludePrintSettings = "includePrintSettings";

            /// <summary>
            /// Include Hidden Rows & Columns
            /// </summary>
            public const string AttrIncludeHiddenRowCol = "includeHiddenRowCol";

            /// <summary>
            /// Maximized
            /// </summary>
            public const string AttrMaximized = "maximized";

            /// <summary>
            /// Active Sheet in Book View
            /// </summary>
            public const string AttrActiveSheetId = "activeSheetId";

            /// <summary>
            /// Show Formula Bar
            /// </summary>
            public const string AttrShowFormulaBar = "showFormulaBar";

            /// <summary>
            /// Show Status Bar
            /// </summary>
            public const string AttrShowStatusbar = "showStatusbar";

            /// <summary>
            /// Show Comments
            /// </summary>
            public const string AttrShowComments = "showComments";

            /// <summary>
            /// Show Objects
            /// </summary>
            public const string AttrShowObjects = "showObjects";

            /// <summary>
            /// Sheet Tab Id
            /// </summary>
            public const string AttrSheetId = "sheetId";

            /// <summary>
            /// Visible State
            /// </summary>
            public const string AttrState = "state";

            /// <summary>
            /// Date 1904
            /// </summary>
            public const string AttrDate1904 = "date1904";

            /// <summary>
            /// Show Border Unselected Table
            /// </summary>
            public const string AttrShowBorderUnselectedTables = "showBorderUnselectedTables";

            /// <summary>
            /// Filter Privacy
            /// </summary>
            public const string AttrFilterPrivacy = "filterPrivacy";

            /// <summary>
            /// Prompted Solutions
            /// </summary>
            public const string AttrPromptedSolutions = "promptedSolutions";

            /// <summary>
            /// Show Ink Annotations
            /// </summary>
            public const string AttrShowInkAnnotation = "showInkAnnotation";

            /// <summary>
            /// Create Backup File
            /// </summary>
            public const string AttrBackupFile = "backupFile";

            /// <summary>
            /// Save External Link Values
            /// </summary>
            public const string AttrSaveExternalLinkValues = "saveExternalLinkValues";

            /// <summary>
            /// Update Links Behavior
            /// </summary>
            public const string AttrUpdateLinks = "updateLinks";

            /// <summary>
            /// Hide Pivot Field List
            /// </summary>
            public const string AttrHidePivotFieldList = "hidePivotFieldList";

            /// <summary>
            /// Show Pivot Chart Filter
            /// </summary>
            public const string AttrShowPivotChartFilter = "showPivotChartFilter";

            /// <summary>
            /// Allow Refresh Query
            /// </summary>
            public const string AttrAllowRefreshQuery = "allowRefreshQuery";

            /// <summary>
            /// Publish Items
            /// </summary>
            public const string AttrPublishItems = "publishItems";

            /// <summary>
            /// Check Compatibility On Save
            /// </summary>
            public const string AttrCheckCompatibility = "checkCompatibility";

            /// <summary>
            /// Auto Compress Pictures
            /// </summary>
            public const string AttrAutoCompressPictures = "autoCompressPictures";

            /// <summary>
            /// Refresh all Connections on Open
            /// </summary>
            public const string AttrRefreshAllConnections = "refreshAllConnections";

            /// <summary>
            /// Default Theme Version
            /// </summary>
            public const string AttrDefaultThemeVersion = "defaultThemeVersion";

            /// <summary>
            /// Embed SmartTags
            /// </summary>
            public const string AttrEmbed = "embed";

            /// <summary>
            /// Show Smart Tags
            /// </summary>
            public const string AttrShow = "show";

            /// <summary>
            /// SmartTag Namespace URI
            /// </summary>
            public const string AttrNamespaceUri = "namespaceUri";

            /// <summary>
            /// Smart Tag URL
            /// </summary>
            public const string AttrUrl = "url";

            /// <summary>
            /// Auto Recover
            /// </summary>
            public const string AttrAutoRecover = "autoRecover";

            /// <summary>
            /// Crash Save
            /// </summary>
            public const string AttrCrashSave = "crashSave";

            /// <summary>
            /// Data Extract Load
            /// </summary>
            public const string AttrDataExtractLoad = "dataExtractLoad";

            /// <summary>
            /// Repair Load
            /// </summary>
            public const string AttrRepairLoad = "repairLoad";

            /// <summary>
            /// Calculation Id
            /// </summary>
            public const string AttrCalcId = "calcId";

            /// <summary>
            /// Calculation Mode
            /// </summary>
            public const string AttrCalcMode = "calcMode";

            /// <summary>
            /// Full Calculation On Load
            /// </summary>
            public const string AttrFullCalcOnLoad = "fullCalcOnLoad";

            /// <summary>
            /// Reference Mode
            /// </summary>
            public const string AttrRefMode = "refMode";

            /// <summary>
            /// Calculation Iteration
            /// </summary>
            public const string AttrIterate = "iterate";

            /// <summary>
            /// Iteration Count
            /// </summary>
            public const string AttrIterateCount = "iterateCount";

            /// <summary>
            /// Iterative Calculation Delta
            /// </summary>
            public const string AttrIterateDelta = "iterateDelta";

            /// <summary>
            /// Full Precision Calculation
            /// </summary>
            public const string AttrFullPrecision = "fullPrecision";

            /// <summary>
            /// Calc Completed
            /// </summary>
            public const string AttrCalcCompleted = "calcCompleted";

            /// <summary>
            /// Calculate On Save
            /// </summary>
            public const string AttrCalcOnSave = "calcOnSave";

            /// <summary>
            /// Concurrent Calculations
            /// </summary>
            public const string AttrConcurrentCalc = "concurrentCalc";

            /// <summary>
            /// Concurrent Thread Manual Count
            /// </summary>
            public const string AttrConcurrentManualCount = "concurrentManualCount";

            /// <summary>
            /// Force Full Calculation
            /// </summary>
            public const string AttrForceFullCalc = "forceFullCalc";

            /// <summary>
            /// Comment
            /// </summary>
            public const string AttrComment = "comment";

            /// <summary>
            /// Custom Menu Text
            /// </summary>
            public const string AttrCustomMenu = "customMenu";

            /// <summary>
            /// Description
            /// </summary>
            public const string AttrDescription = "description";

            /// <summary>
            /// Help
            /// </summary>
            public const string AttrHelp = "help";

            /// <summary>
            /// Status Bar
            /// </summary>
            public const string AttrStatusBar = "statusBar";

            /// <summary>
            /// Local Name Sheet Id
            /// </summary>
            public const string AttrLocalSheetId = "localSheetId";

            /// <summary>
            /// Hidden Name
            /// </summary>
            public const string AttrHidden = "hidden";

            /// <summary>
            /// Function
            /// </summary>
            public const string AttrFunction = "function";

            /// <summary>
            /// Procedure
            /// </summary>
            public const string AttrVbProcedure = "vbProcedure";

            /// <summary>
            /// External Function
            /// </summary>
            public const string AttrXlm = "xlm";

            /// <summary>
            /// Function Group Id
            /// </summary>
            public const string AttrFunctionGroupId = "functionGroupId";

            /// <summary>
            /// Shortcut Key
            /// </summary>
            public const string AttrShortcutKey = "shortcutKey";

            /// <summary>
            /// Publish To Server
            /// </summary>
            public const string AttrPublishToServer = "publishToServer";

            /// <summary>
            /// Workbook Parameter (Server)
            /// </summary>
            public const string AttrWorkbookParameter = "workbookParameter";

            /// <summary>
            /// PivotCache Id
            /// </summary>
            public const string AttrCacheId = "cacheId";

            /// <summary>
            /// Read Only Recommended
            /// </summary>
            public const string AttrReadOnlyRecommended = "readOnlyRecommended";

            /// <summary>
            /// User Name
            /// </summary>
            public const string AttrUserName = "userName";

            /// <summary>
            /// Write Reservation Password
            /// </summary>
            public const string AttrReservationPassword = "reservationPassword";

            /// <summary>
            /// Reference
            /// </summary>
            public const string AttrRef = "ref";

            /// <summary>
            /// Workbook Password
            /// </summary>
            public const string AttrWorkbookPassword = "workbookPassword";

            /// <summary>
            /// Revisions Password
            /// </summary>
            public const string AttrRevisionsPassword = "revisionsPassword";

            /// <summary>
            /// Lock Structure
            /// </summary>
            public const string AttrLockStructure = "lockStructure";

            /// <summary>
            /// Lock Windows
            /// </summary>
            public const string AttrLockWindows = "lockWindows";

            /// <summary>
            /// Lock Revisions
            /// </summary>
            public const string AttrLockRevision = "lockRevision";

            /// <summary>
            /// Use CSS
            /// </summary>
            public const string AttrCss = "css";

            /// <summary>
            /// Thicket
            /// </summary>
            public const string AttrThicket = "thicket";

            /// <summary>
            /// Enable Long File Names
            /// </summary>
            public const string AttrLongFileNames = "longFileNames";

            /// <summary>
            /// VML in Browsers
            /// </summary>
            public const string AttrVml = "vml";

            /// <summary>
            /// Allow PNG
            /// </summary>
            public const string AttrAllowPng = "allowPng";

            /// <summary>
            /// Target Screen Size
            /// </summary>
            public const string AttrTargetScreenSize = "targetScreenSize";

            /// <summary>
            /// DPI
            /// </summary>
            public const string AttrDpi = "dpi";

            /// <summary>
            /// Code Page
            /// </summary>
            public const string AttrCodePage = "codePage";

            /// <summary>
            /// Built-in Function Group Count
            /// </summary>
            public const string AttrBuiltInGroupCount = "builtInGroupCount";

            /// <summary>
            /// Count
            /// </summary>
            public const string AttrCount = "count";

            /// <summary>
            /// Id
            /// </summary>
            public const string AttrId = "id";

            /// <summary>
            /// Div Id
            /// </summary>
            public const string AttrDivId = "divId";

            /// <summary>
            /// Source Object
            /// </summary>
            public const string AttrSourceObject = "sourceObject";

            /// <summary>
            /// Destination File
            /// </summary>
            public const string AttrDestinationFile = "destinationFile";

            /// <summary>
            /// Title
            /// </summary>
            public const string AttrTitle = "title";

            /// <summary>
            /// Auto Republish
            /// </summary>
            public const string AttrAutoRepublish = "autoRepublish";

        }
        #endregion
    }
}
