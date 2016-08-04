using System;
using System.Diagnostics;
using System.Net;
using System.Web;

namespace GTranslate
{
    public class GoogleTranslator
    {     

        string GetResultPage(string text, string LanguageFrom, string LanguageTo)
        {
            WebClient client = new WebClient();

            string URI = String.Format("http://www.google.com/translate_t?hl=en&ie=UTF8&text={0}&langpair={1}|{2}",
                HttpUtility.UrlEncode(text), LanguageFrom, LanguageTo);

            try
            {
                return client.DownloadString(URI);
            }
            finally
            {
                client.Dispose();
            }
        }


        public string Translate(string text, string LanguageFrom, string LanguageTo)
        {
            string result = "nothing yet...";
            string resultPage = GetResultPage(text, LanguageFrom, LanguageTo);

            const string anchor = "TRANSLATED_TEXT='";

            try
            {
                result = resultPage.Substring(resultPage.IndexOf(anchor) + anchor.Length);
                result = result.Substring(0, result.IndexOf("';INPUT_TOOL_PATH="));
                result = result.Replace(@"\x26quot;", "\"") // "
                    .Replace(@"\x26#39;", "'") // '
                    .Replace(@"\x26amp;", "&") // &
                    .Replace(@"\x26gt;", ">")  // >
                    .Replace(@"\x26lt;", "<")  // <
                    .Replace(@"\x3d", "=")     // =
                    .Replace(@"\r\x3cbr\x3e", "\r"); // new line
            }
            catch (Exception e)
            {
                //result = "Error: " + e.Message;
                return String.Empty;
            }

            return result;
        }
    }
}
