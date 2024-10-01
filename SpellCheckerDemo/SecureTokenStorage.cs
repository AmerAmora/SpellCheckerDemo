using Microsoft.Win32;
using System.Security.Cryptography;
using System.Text;

namespace SpellCheckerDemo;

public static class SecureTokenStorage
{
    private const string RegistryKey = @"SOFTWARE\Qalam\Login";
    private const string TokenValueName = "EncryptedToken";

    public static void StoreToken(string token)
    {
        byte[] encryptedData = ProtectedData.Protect(
            Encoding.UTF8.GetBytes(token) ,
            null ,
            DataProtectionScope.CurrentUser);

        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKey))
        {
            key.SetValue(TokenValueName , Convert.ToBase64String(encryptedData));
        }
    }

    public static string RetrieveToken()
    {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKey))
        {
            if (key == null)
                return null;

            string encryptedBase64 = key.GetValue(TokenValueName) as string;
            if (string.IsNullOrEmpty(encryptedBase64))
                return null;

            byte[] encryptedData = Convert.FromBase64String(encryptedBase64);
            byte[] decryptedData = ProtectedData.Unprotect(
                encryptedData ,
                null ,
                DataProtectionScope.CurrentUser);

            return Encoding.UTF8.GetString(decryptedData);
        }
    }

    public static void ClearToken()
    {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKey , true))
        {
            if (key != null)
            {
                key.DeleteValue(TokenValueName , false);
            }
        }
    }
}
