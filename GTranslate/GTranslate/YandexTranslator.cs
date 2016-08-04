using System;
using System.Text;
using System.Net;
using System.Xml;

namespace GTranslate
{
    public class YandexTranslator
    {

        private string DownloadResult(string text, string LanguageFrom, string LanguageTo, string Key)
        {

            string URI = String.Format("https://translate.yandex.net/api/v1.5/tr/translate?key={0}&lang={1}-{2}&text={3}",
                Key, LanguageFrom, LanguageTo, text);

            try
            {
                return (new WebClient()).DownloadString(URI);
            }
            catch (WebException e)
            {
                //return  "Error: " + e.Message;
                return String.Empty;
            }
        }

        public string Translate(string text, string LanguageFrom, string LanguageTo, string Key)
        {
            string translation = "nothing yet...";
            XmlDocument xmlDocumet = new XmlDocument();

            string y = DownloadResult(text, LanguageFrom, LanguageTo, Key);
            byte[] bytes = Encoding.Default.GetBytes(y);
            translation = Encoding.UTF8.GetString(bytes);
            xmlDocumet.LoadXml(translation);
            translation = xmlDocumet.GetElementsByTagName("text")[0].InnerText;

            return translation;
        }
    }
}
