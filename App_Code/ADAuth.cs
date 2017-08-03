using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.DirectoryServices.AccountManagement;

/// <summary>
/// 回傳目前使用者的各種AD資訊
/// 
/// 使用順序：
/// GetDomainUserName->GetUserID->GetGroupLists->SearchInGroups
/// 
/// 註：
/// 1 System.DirectoryServices.AccountManagement 必須用加入參考(Add reference)的方式
/// 光用 using 是不夠的
/// 
/// 2 allowADGroups, domainName 按需求修改
/// 
/// </summary>


namespace ADAuth
{
    public class Auth
    {
        string[] allowADGroups = {
            "Domain Admins",
            "13201採購組",
            "17201會計課"
        };

        string domainName = "Motorpro-sbs";

        public string GetDomainUserName()
        {
            string domainUserName = HttpContext.Current.User.Identity.Name;
            return domainUserName;
        }

        public bool SearchInGroups(List<string> groupList)
        {
            bool inGroup = false;

            foreach (string g in allowADGroups)
            {
                if (groupList.Contains(g))
                {
                    inGroup = true;
                    break; // 只要其中一個群組符合條件即可
                }
            }
            return inGroup;

        }

        public string GetDomainName(string domainUsers)
        {
            string[] usernameArray = domainUsers.Split('\\');
            return usernameArray[0].ToString();
        }

        public string GetUserID(string domainUsers)
        {
            string[] usernameArray = domainUsers.Split('\\');
            return usernameArray[1].ToString();
        }

        public List<string> GetGroupLists(string userName)
        {
            var pc = new PrincipalContext(ContextType.Domain, domainName);
            var src = UserPrincipal.FindByIdentity(pc, userName).GetGroups(pc);
            var result = new List<string>();
            src.ToList().ForEach(sr => result.Add(sr.SamAccountName));
            return result;
        }
    }
}