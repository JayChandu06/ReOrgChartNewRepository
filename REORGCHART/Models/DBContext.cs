using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace REORGCHART.Models
{
    public class DBContext : DbContext
    {
        // You can add custom code to this file. Changes will not be overwritten.
        // 
        // If you want Entity Framework to drop and regenerate your database
        // automatically whenever you change your model schema, please use data migrations.
        // For more information refer to the documentation:
        // http://msdn.microsoft.com/en-us/data/jj591621.aspx

        public DBContext() : base("name=DBContext")
        {
        }

        public System.Data.Entity.DbSet<REORGCHART.Models.UserLastActions> UserLastActions { get; set; }
        public System.Data.Entity.DbSet<REORGCHART.Models.AspNetUsers> AspNetUsers { get; set; }
        public System.Data.Entity.DbSet<REORGCHART.Models.CompanyDetails> Company { get; set; }
        public System.Data.Entity.DbSet<REORGCHART.Models.UserListDetails> UserListDetails { get; set; }
        public System.Data.Entity.DbSet<REORGCHART.Models.UserVersions> UserVersions { get; set; }
        public System.Data.Entity.DbSet<REORGCHART.Models.UserRoles> UserRoles { get; set; }
        public System.Data.Entity.DbSet<REORGCHART.Models.UploadFilesHeaders> UploadFilesHeaders { get; set; }
        public System.Data.Entity.DbSet<REORGCHART.Models.UploadFileDetails> UploadFileDetails { get; set; }
    }

    public class AspNetUsers
    {
        [Key]
        public string Id { get; set; }
        public string Email { get; set; }
        public Boolean EmailConfirmed { get; set; }
        public string PasswordHash { get; set; }
        public string SecurityStamp { get; set; }
        public string PhoneNumber { get; set; }
        public Boolean PhoneNumberConfirmed { get; set; }
        public Boolean TwoFactorEnabled { get; set; }
        public string LockoutEndDateUtc { get; set; }
        public Boolean LockoutEnabled { get; set; }
        public int? AccessFailedCount { get; set; }
        public string UserName { get; set; }
        public int UID { get; set; }
        public string CompanyName { get; set; }
    }

    public class UserLastActions
    {
        public string Role { get; set; }
        public string ShowLevel { get; set; }
        public string Levels { get; set; }
        public string Version { get; set; }
        public string Oper { get; set; }
        public string UsedView { get; set; }
        public string Country { get; set; }
        public string KeyDate { get; set; }
        [Key]
        [Column(Order = 1)]
        public string Company { get; set; }
        [Key]
        [Column(Order = 2)]
        public string UserId { get; set; }
    }

    public class UserVersions
    {
        [Key]
        [Column(Order = 1)]
        public int UserId { get; set; }
        [Key]
        [Column(Order = 2)]
        public string Company { get; set; }
        public int VersionId { get; set; }
        public string UploadedFileName { get; set; }
        public string VersionFileName { get; set; }
        [Key]
        [Column(Order = 3)]
        public DateTime CreatedDate { get; set; }
    }

    public class UserRoles
    {
        [Key]
        [Column(Order = 1)]
        public string UserId { get; set; }
        [Key]
        [Column(Order = 2)]
        public string Role { get; set; }
    }

    public class UploadFilesHeaders
    {
        public string JSONFileName { get; set; }
        public string KeyField { get; set; }
        public string ParentField { get; set; }
        public DateTime CreatedDate { get; set; }
        public string SerialNoFlag { get; set; }
        public string FirstPosition { get; set; }
        public string FirstPositionField { get; set; }
        public string FileType { get; set; }
        public string UseFields { get; set; }
        public string FullNameFields { get; set; }
        public int CurrentVersionNo { get; set; }
        public int VersionNo { get; set; }
        [Key]
        [Column(Order = 3)]
        public string CompanyName { get; set; }
        [Key]
        [Column(Order = 2)]
        public string Role { get; set; }
        [Key]
        [Column(Order = 1)]
        public string UserId { get; set; }
    }

    public class UploadFileDetails
    {
        public string JSONFileName { get; set; }
        public DateTime KeyDate { get; set; }
        public string VersionStatus { get; set; }
        [Key]
        [Column(Order = 4)]
        public int VersionNo { get; set; }
        [Key]
        [Column(Order = 3)]
        public string CompanyName { get; set; }
        [Key]
        [Column(Order = 2)]
        public string Role { get; set; }
        [Key]
        [Column(Order = 1)]
        public string UserId { get; set; }
    }

    public class CompanyDetails
    {
        [Key]
        public int CompanyID { get; set; }
        public string CompanyName { get; set; }
        public string UserName { get; set; }
        public string EmailID { get; set; }
        public string Logo { get; set; }
    }

    public class UserListDetails
    {
        [Key]
        public string UserName { get; set; }
        public string CompanyName { get; set; }
        public string Role { get; set; }
        public string Id { get; set; }
    }

    public class MyModel
    {
        public DateTime UseDate { get; set; }
        public string KeyDate { get; set; }
        public string Country { get; set; }
        public string ShowLevel { get; set; }
        public string Levels { get; set; }
        public string Oper { get; set; }
        public string View { get; set; }
        public string Version { get; set; }
        public string ChartData { get; set; }
        public string Role { get; set; }
        public string UserId { get; set; }
    }

    public class MyLastAction
    {
        public string KeyDate { get; set; }
        public string Country { get; set; }
        public string ShowLevel { get; set; }
        public string Levels { get; set; }
        public string Oper { get; set; }
        public string View { get; set; }
        public string Version { get; set; }
        public string Role { get; set; }
    }

    public class OrgChartData
    {
        public string key { get; set; }
        public string parent { get; set; }
        public string FULL_NAME { get; set; }
        public string GENDER { get; set; }
        public string DEPTID { get; set; }
        public string LOCATION { get; set; }
    }

    public class ChartDataInf
    {
        public string className { get; set; }
        public List<OrgChartData> nodeDataArray { get; set; }
    }

    public class MyData
    {
        public string LEVEL_ID { get; set; }
        public string PARENT_LEVEL_ID { get; set; }
        public string DATE_UPDATED { get; set; }
        public string NAME_PREFIX { get; set; }
        public string FULL_NAME { get; set; }
        public string LAST_NAME { get; set; }
        public string MIDDLE_NAME { get; set; }
        public string FIRST_NAME { get; set; }
        public string LOCATION { get; set; }
        public string LOCATION_DESCR { get; set; }
        public string EMAILID { get; set; }
        public string GENDER { get; set; }
        public string DEPTID { get; set; }
        public string DIVISION { get; set; }
        public string POSITION_TITLE { get; set; }
        public string JOB_TITLE { get; set; }
        public string FUNC_GP_L1_DESCR { get; set; }
        public string POSITION_NBR { get; set; }
        public string FUNC_GP_L2_DESCR { get; set; }
        public string FUNC_GP_L3_DESCR { get; set; }
        public string POSITION_GRADE { get; set; }
        public string FUNDING { get; set; }
        public string EMPLOYEE_TYPE { get; set; }
        public string EMPLOYEE_STATUS { get; set; }
        public string TYPE_OF_CONTRACT { get; set; }
        public string REGIONNAME { get; set; }
        public string SUBREGIONNAME { get; set; }
        public string OPERATIONNAME { get; set; }
        public string FOCUS_OFFICENAME { get; set; }
        public string OFFICEMSRPCODE { get; set; }
        public string AD_DISPLAY_NAME { get; set; }
        public string AD_TITLE { get; set; }
        public string AD_DEPARTMENT { get; set; }
        public string IPPHONE { get; set; }
        public string MSRP_COUNTRY { get; set; }
        public string POSITION_COST { get; set; }
        public string SUP_DISPLAY_NAME { get; set; }
    }
}
