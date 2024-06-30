using System;
using System.Data;

namespace HPQ.Excalibur
{
    public class Security
    {

        private bool _IsSysAdmin = false;
        public bool IsSysAdmin
        {
            get { return _IsSysAdmin; }
        }

        private string _CurrentUser;
        public string CurrentUser
        {
            get { return _CurrentUser; }
            set { _CurrentUser = value; }
        }

        private string _CurrentUserFullName = string.Empty;
        public string CurrentUserFullName
        {
            get { return _CurrentUserFullName; }
        }

        private string _CurrentUserDomain = string.Empty;
        public string CurrentUserDomain
        {
            get { return _CurrentUserDomain; }
        }

        private int _CurrentUserID = 0;
        public int CurrentUserID
        {
            get { return _CurrentUserID; }
        }

        private string _CurrentUserEmail = string.Empty;
        public string CurrentUserEmail
        {
            get { return _CurrentUserEmail; }
        }

        private int _CurrentPartnerID = 0;
        public int CurrentPartnerID
        {
            get { return _CurrentPartnerID; }
        }

        private bool _BpiaApprover = false;
        public bool BpiaApprover
        {
            get { return _BpiaApprover; }
        }


        public Security(string User)
        {
            _CurrentUser = User;

            if (User.IndexOf("\\") > 0)
            {
                _CurrentUserDomain = User.Substring(0, User.IndexOf("\\"));
                _CurrentUser = User.Substring(User.IndexOf("\\") + 1);
            }

            HPQ.Excalibur.Data dw = new HPQ.Excalibur.Data();
            DataTable dt = dw.GetUserInfo(CurrentUser, CurrentUserDomain);

            if (dt.Rows.Count > 0)
            {
                _CurrentUserID = Convert.ToInt32(dt.Rows[0]["ID"]);
                _CurrentUserEmail = Convert.ToString(dt.Rows[0]["email"]);
                _CurrentPartnerID = Convert.ToInt32(dt.Rows[0]["partnerid"]);
                _CurrentUserFullName = Convert.ToString(dt.Rows[0]["name"]);
                _BpiaApprover = Convert.ToBoolean(dt.Rows[0]["BpiaApprover"]);
            }

            _IsSysAdmin = IsExcaliburAdmin();
        }

        public bool IsProgramCoordinator()
        {

            Data dw = new Data();
            DataTable dt = dw.GetProgramCoordinatorStatus(_CurrentUserID.ToString());

            if (dt.Rows.Count > 0)
                return true;
            else
                return false;
        }

        private bool IsExcaliburAdmin()
        {
            Data dw = new Data();
            DataTable dt = dw.SelectEmployees(_CurrentUserID.ToString(), "1", string.Empty, string.Empty, string.Empty);

            if (dt.Rows.Count > 0)
                return true;
            else
                return false;
        }

        public bool UserInRole(string RoleCd)
        {
            return UserInRole(string.Empty, RoleCd);
        }
        public bool UserInRole(string ProductVersionId, string RoleCd)
        {
            Data dw = new Data();
            DataTable dt = null;

            dt = Employee.GetUserInRole(_CurrentUserID.ToString(), ProductVersionId, RoleCd);
            return (Convert.ToInt64(dt.Rows[0]["UserInRole"]) > 0);
        }

         public bool UserInRole(ProgramRoles Role)
        {
            int RoleUserId = 0;
            Data dw = new Data();
            DataTable dt = null;

            dt = dw.GetUserRoles(_CurrentUserID.ToString());
            if (dt.Rows.Count == 0)
                return false;

            switch (Role)
            {
                case ProgramRoles.SystemManager:
                    int.TryParse(dt.Rows[0]["SMID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.POPM:
                    int.TryParse(dt.Rows[0]["PMID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.SEPM:
                    int.TryParse(dt.Rows[0]["SEPMID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.CM:
                    int.TryParse(dt.Rows[0]["PMID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.CommercialMarketing:
                    int.TryParse(dt.Rows[0]["ComMarketingID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.ConsumerMarketing:
                    int.TryParse(dt.Rows[0]["ConsMarketingID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.SmbMarketing:
                    int.TryParse(dt.Rows[0]["SmbMarketingID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.PlatformDevelopment:
                    int.TryParse(dt.Rows[0]["PlatformDevelopmentID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.SupplyChain:
                    int.TryParse(dt.Rows[0]["SupplyChainID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.Service:
                    int.TryParse(dt.Rows[0]["ServiceID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.Finance:
                    int.TryParse(dt.Rows[0]["FinanceID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.CommodityPm:
                    int.TryParse(dt.Rows[0]["PDEID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.AccessoryPm:
                    int.TryParse(dt.Rows[0]["AccessoryPMID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.MarketingOps:
                    int.TryParse(dt.Rows[0]["MarketingOpsID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.PC:
                    int.TryParse(dt.Rows[0]["PCID"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.SEPE:
                    int.TryParse(dt.Rows[0]["SEPE"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.PreinstallPm:
                    int.TryParse(dt.Rows[0]["PINPM"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.SETestLead:
                    int.TryParse(dt.Rows[0]["SETestLead"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.GPLM:
                    int.TryParse(dt.Rows[0]["GPLM"].ToString(), out RoleUserId);
                    break;
                case ProgramRoles.ServiceBomAnalyst:
                    int.TryParse(dt.Rows[0]["SvcBomAnalyst"].ToString(), out RoleUserId);
                    break;
                default:
                    break;
            }

            return (RoleUserId == _CurrentUserID);

        }



        public enum ProgramRoles
        {
            SystemManager,
            POPM,
            SEPM,
            CM,
            CommercialMarketing,
            ConsumerMarketing,
            SmbMarketing,
            PlatformDevelopment,
            SupplyChain,
            Service,
            Finance,
            CommodityPm,
            AccessoryPm,
            MarketingOps,
            PC,
            SEPE,
            PreinstallPm,
            SETestLead,
            GPLM,
            ServiceBomAnalyst
        }

    }

}
