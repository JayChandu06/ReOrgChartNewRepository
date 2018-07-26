/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This class defines the application constants used at server side and is complimentary to the constant.js script file.
  

***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/



using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OrgChartModelling.Utils
{
    public class AppConstants
    {

        public static String ORGANIZATION = "";

        public const String MODIFIED_STATUS = "CLINETMODIFIEDSTATUS";
        public const String MODIFIED_STATUS_NEWPERSONASSIGNEDTOPOSITION = "6";
        public const String MODIFIED_STATUS_PERSONUNASSIGNEDFROMPOSITION = "5";
        public const String MODIFIED_STATUS_HASNEWCPM = "4";
        public const String MODIFIED_STATUS_WASOLDCPM = "3";
        public const String MODIFIED_STATUS_CREATED = "2";
        public const String MODIFIED_STATUS_MODIFIED = "1";
        public const String MODIFIED_STATUS_RELINKED = "1";
        public const String MODIFIED_STATUS_DELETED = "-1";

        public const String OUTPUT_MESSAGE_COLUMN_NAME = "MESSAGE";
        public const String OUTPUT_STATUS_COLUMN_NAME = "STATUS";
        public const String OUTPUT_STATUS_TRUE = "SUCCESSFUL";
        public const String OUTPUT_STATUS_FALSE = "FAILED";


        public const string USER_SUCCESS_MESSAGE = "Operation completed successfully!";
        public const string USER_FAILURE_MESSAGE = "Operation failed because of unknown reason!";

        public const string USER_SUCCESS_MESSAGE_VERSION_SAVED = "Version saved successfully!";

        public const string USER_SUCCESS_MESSAGE_INITIATIVE_CREATED = "New initiative created successfully!";
        public const string USER_SUCCESS_MESSAGE_POPULATION_CREATED = "New population created successfully!";
        public const string USER_SUCCESS_MESSAGE_POPULATION_DELETED = "Population deleted successfully!";
        public const string USER_SUCCESS_MESSAGE_VERSION_CREATED = "New version created successfully!";
        public const string USER_SUCCESS_MESSAGE_VERSION_DELETED = "Version deleted successfully!";


        public const string DOWNLOAD_VERSION_OPERATIONAL_SHEET1_NAME = "Transfer List";
        public const string DOWNLOAD_VERSION_OPERATIONAL_SHEET2_NAME = "Audit Trial";

        public const string DOWNLOAD_VERSION_LEGAL_SHEET1_NAME = "Orgunits";
        public const string DOWNLOAD_VERSION_LEGAL_SHEET2_NAME = "Positions";
        public const string DOWNLOAD_VERSION_LEGAL_SHEET3_NAME = "Audit Trial";

        public const string DOWNLOAD_INITIATIVE_OPERATIONAL_SHEET1_NAME = "Transfer List";
        public const string DOWNLOAD_INITIATIVE_OPERATIONAL_SHEET2_NAME = "Audit Trial";
        public const string DOWNLOAD_INITIATIVE_OPERATIONAL_SHEET3_NAME = "Final Version Details";

        public const string DOWNLOAD_INITIATIVE_LEGAL_SHEET1_NAME = "Orgunits";
        public const string DOWNLOAD_INITIATIVE_LEGAL_SHEET2_NAME = "Positions";
        public const string DOWNLOAD_INITIATIVE_LEGAL_SHEET3_NAME = "Audit Trial";
        public const string DOWNLOAD_INITIATIVE_LEGAL_SHEET4_NAME = "Final Version Details";

        public const string DOWNLOAD_PATH = "/download/";
        public const string UPLOAD_PATH = "/upload/";


        public const string DOMAIN = "EUNET";

        public const string ADMIN_GROUP = "EUNET\\REORG_CHART_ADMIN";
        public const string LEAD_GROUP = "EUNET\\REORG_CHART_LEADS";
        public const string USER_GROUP = "EUNET\\REORG_CHART_USERS";

       // public const string CHART_TYPE_LEGAL = "LV";
       // public const string CHART_TYPE_OPERATIONAL = "OV";
        public const string CHART_TYPE_LEGAL = "";
        public const string CHART_TYPE_OPERATIONAL = "";
    }



        public class Node
        {
            public string key { get; set; }
            public string parent { get; set; }
            

            public string PositionID { get; set; }
            public string NextPositionID { get; set; }
            public string FullName { get; set; }
            public string orgunitText { get; set; }
            public string positiontitle { get; set; }
            public string DIVISION { get; set; }

            public string ModifiedStatus { get; set; }

        }

    
}