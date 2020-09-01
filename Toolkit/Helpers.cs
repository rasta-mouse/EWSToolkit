using System.Security;

namespace Toolkit
{
    public static class Helpers
    {
        public static SecureString ToSecureString(this string password)
        {
            var securePassword = new SecureString();

            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }

            securePassword.MakeReadOnly();
            return securePassword;
        }
    }
}