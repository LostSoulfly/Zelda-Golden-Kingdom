using GTranslate.Translators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TranslationServer
{
    public static class Translator
    {
        public static string CreateMD5(string input, bool noQuotes = false)
        {

            if (noQuotes)
            {
                if (input.Substring(0, 1) == "\"" & input.Substring(input.Length - 1, 1) == "\"")
                    input = input.Substring(1, input.Length - 2);
            }

            // Use input string to calculate MD5 hash
            using (System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create())
            {
                byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                return Convert.ToHexString(hashBytes); // .NET 5 +

                // Convert the byte array to hexadecimal string prior to .NET 5
                // StringBuilder sb = new System.Text.StringBuilder();
                // for (int i = 0; i < hashBytes.Length; i++)
                // {
                //     sb.Append(hashBytes[i].ToString("X2"));
                // }
                // return sb.ToString();
            }
        }

        public static async Task<string> Google(string Text)
        {

            if ((Text == null) || (Text.Length == 0))
            {
                return String.Empty;
            }

            var t = new GoogleTranslator();
            string translation;

            try
            {
                var task = t.TranslateAsync(Text, "en", "es");
                task.Wait();


                translation = task.Result.Translation;
                return translation.Replace(System.Environment.NewLine, "\r\n");
            }
            catch (Exception ex)
            {
                //return "Error: " + ex.Message;
                return String.Empty;
            }
        }

        public static async Task<string> Bing(string Text)
        {

            if ((Text == null) || (Text.Length == 0))
            {
                return String.Empty;
            }

            var t = new BingTranslator();
            string translation;

            try
            {
                var task = t.TranslateAsync(Text, "en", "es");
                task.Wait();


                translation = task.Result.Translation;
                return translation.Replace(System.Environment.NewLine, "\r\n");
            }
            catch (Exception ex)
            {
                //return "Error: " + ex.Message;
                return String.Empty;
            }
        }

        public static async Task<string> Yandex(string Text)
        {

            if ((Text == null) || (Text.Length == 0))
            {
                return String.Empty;
            }

            var t = new YandexTranslator();
            string translation;

            try
            {
                var task = t.TranslateAsync(Text, "en", "es");
                task.Wait();


                translation = task.Result.Translation;
                return translation.Replace(System.Environment.NewLine, "\r\n");
            }
            catch (Exception ex)
            {
                //return "Error: " + ex.Message;
                return String.Empty;
            }
        }


    }
}
