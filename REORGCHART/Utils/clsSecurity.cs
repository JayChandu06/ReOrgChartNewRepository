/**************************************************************************************************************************************
Initial Version
***************************************************************************************************************************************
Developer                 Date      Change Document No.      Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                 First Release. 

This utility class defines encryption/decryption mechanism for password and other sensitive information which are stored in web.config  
File in encrypted format.

***************************************************************************************************************************************
Change Log
***************************************************************************************************************************************
Developer                 Date      Change Document No.       Version                Description
Amresh Jha (JHAAM2)   25-May-2016    CD 1400013968             1.0                   First Release. 

***************************************************************************************************************************************/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using System.Security.Cryptography;
using System.IO;


namespace OrgChartModelling.Utils
{
    public class clsSecurity
    {
        public static string ConvertStringToHex(String input, System.Text.Encoding encoding)
        {
            Byte[] stringBytes = encoding.GetBytes(input);
            StringBuilder sbBytes = new StringBuilder(stringBytes.Length * 2);
            foreach (byte b in stringBytes)
            {
                sbBytes.AppendFormat("{0:X2}", b);
            }
            return sbBytes.ToString();
        }
        public static string ConvertHexToString(String hexInput, System.Text.Encoding encoding)
        {
            int numberChars = hexInput.Length;
            byte[] bytes = new byte[numberChars / 2];
            for (int i = 0; i <= numberChars - 1; i += 2)
            {
                bytes[i / 2] = Convert.ToByte(hexInput.Substring(i, 2), 16);
            }
            return encoding.GetString(bytes);
        }
        private static string GetSalt()
        {

            byte[] TempByte = { 65, 99, 98, 84, 105, 116, 115, 72, 115, 73 };
            return System.Text.Encoding.UTF8.GetString(TempByte);
        }
        public string EncryptString(string InputText)
        {
            string Password = GetSalt();
            RijndaelManaged RijndaelCipher = new RijndaelManaged();
            byte[] PlainText = System.Text.Encoding.Unicode.GetBytes(InputText);
            byte[] Salt = Encoding.ASCII.GetBytes(Password.ToString());
            // The (Secret Key) will be generated from the specified password and Salt.
            // Dim SecretKey As New PasswordDeriveBytes(Password, Salt) 
            Rfc2898DeriveBytes SecretKey = new Rfc2898DeriveBytes(Password, Salt);
            // Create a encryptor from the existing SecretKey bytes.
            // We use 32 bytes for the secret key
            // (the default Rijndael key length is 256 bit = 32 bytes) and
            // then 16 bytes for the IV (initialization vector),
            // (the default Rijndael IV length is 128 bit = 16 bytes)
            //Dim Encryptor As ICryptoTransform = RijndaelCipher.CreateEncryptor(SecretKey.GetBytes(16), SecretKey.GetBytes(16))
            ICryptoTransform Encryptor = RijndaelCipher.CreateEncryptor(SecretKey.GetBytes(16), SecretKey.GetBytes(16));
            // Create a MemoryStream that is going to hold the encrypted bytes
            MemoryStream memoryStream = new MemoryStream();
            CryptoStream cryptoStream = new CryptoStream(memoryStream, Encryptor, CryptoStreamMode.Write);
            // Start the encryption process.
            cryptoStream.Write(PlainText, 0, PlainText.Length);
            // Finish encrypting.
            cryptoStream.FlushFinalBlock();
            // Convert our encrypted data from a memoryStream into a byte array.
            byte[] CipherBytes = memoryStream.ToArray();
            // Close both streams.
            memoryStream.Close();
            cryptoStream.Close();
            string EncryptedData = Convert.ToBase64String(CipherBytes);
            // Return encrypted string.
            return EncryptedData;
        }

        public string DecryptString(string InputText)
        {
            string Password = GetSalt();
            try
            {
                RijndaelManaged RijndaelCipher = new RijndaelManaged();
                byte[] EncryptedData = Convert.FromBase64String(InputText);
                byte[] Salt = Encoding.ASCII.GetBytes(Password.ToString());
                // Dim SecretKey As New PasswordDeriveBytes(Password, Salt) - changed by athi again due to obsolete 
                Rfc2898DeriveBytes SecretKey = new Rfc2898DeriveBytes(Password, Salt);
                // Create a decryptor from the existing SecretKey bytes.
                ICryptoTransform Decryptor = RijndaelCipher.CreateDecryptor(SecretKey.GetBytes(16), SecretKey.GetBytes(16));
                MemoryStream memoryStream = new MemoryStream(EncryptedData);
                // Create a CryptoStream. (always use Read mode for decryption).
                CryptoStream cryptoStream = new CryptoStream(memoryStream, Decryptor, CryptoStreamMode.Read);
                // Since at this point we don't know what the size of decrypted data
                // will be, allocate the buffer long enough to hold EncryptedData;
                // DecryptedData is never longer than EncryptedData.
                byte[] PlainText = new byte[EncryptedData.Length];
                // Start decrypting.
                int DecryptedCount = cryptoStream.Read(PlainText, 0, PlainText.Length);
                memoryStream.Close();
                cryptoStream.Close();
                // Convert decrypted data into a string.
                string DecryptedData = Encoding.Unicode.GetString(PlainText, 0, DecryptedCount);
                // Return decrypted string.
                return DecryptedData;
            }
            catch (Exception exception)
            {
                return (exception.Message);
            }
        }

    }
}