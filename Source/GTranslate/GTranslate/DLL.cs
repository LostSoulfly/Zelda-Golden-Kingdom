using System;
using System.Text;
using System.Net;
using System.IO;
using System.Threading.Tasks;
//using GTranslate;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using GTranslate.Translators;

namespace GTranslateDLL
{

    [Guid("ce4a8c39-9c28-4ea7-85cd-f75aa27c82d4")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _DLL
    {
        //Define the com objects that VB6 will interface with through the type library
        [DispId(1)]
        string GoogleTranslate(string LangTo, string LangFrom, string Text);

        [DispId(2)]
        string BingTranslate(string to, string from, string text);

        [DispId(3)]
        String YandexTranslate(string LangTo, string LangFrom, string Text);

        [DispId(4)]
        string GetMD5Hash(string Text);

        [DispId(5)]
        string Version();

        [DispId(6)]
        string GetLastTranslation();

        [DispId(7)]
        void ClearLastTranslation();

        /*
        [DispId(5)]
        int getInt();

        [DispId(6)]
        void setInt(int i);
        */
    }

    [Guid("ce4a8c39-9c28-4ea7-85cd-f75aa27c82d2")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [ProgId("GTranslateDLL.Translate")]
    public class DLL : _DLL
    {
        public string LastTranslation;

        public String GoogleTranslate(string LangTo, string LangFrom, string Text)
        {

            if ((Text == null) || (Text.Length == 0))
            {
                return String.Empty;
            }

            var t = new GoogleTranslator();
            string translation;

            try
            {
                var task = t.TranslateAsync(Text, LangTo, LangFrom);
                task.Wait();


                translation = task.Result.Translation;
                LastTranslation = translation;
                return translation.Replace(System.Environment.NewLine, "\r\n");
            }
            catch (Exception ex)
            {
                //return "Error: " + ex.Message;
                return String.Empty;
            }

        }

        public string BingTranslate(string LangTo, string LangFrom, string Text)
        {

            var t = new YandexTranslator();

            string translation;

            try
            {

                var task = t.TranslateAsync(Text, LangTo, LangFrom);
                task.Wait();


                translation = task.Result.Translation.Replace(System.Environment.NewLine, "\r\n");
                LastTranslation = translation;
                return translation;
            }
            catch (Exception ex)
            {
                //return "Error: " + ex.Message;
                return String.Empty;
            }
        }

        public String YandexTranslate(string LangTo, string LangFrom, string Text)
        {

            var t = new YandexTranslator();
            string translation;

            try
            {

                var task = t.TranslateAsync(Text, LangTo, LangFrom);
                task.Wait();


                translation = task.Result.Translation;
                LastTranslation = translation;
                return translation.Replace(System.Environment.NewLine, "\r\n");
            }
            catch (Exception ex)
            {
                //return "Error: " + ex.Message;
                return String.Empty;
            }

        }

        public String GetMD5Hash(String Text)
        {
            //Check wether data was passed
            if ((Text == null) || (Text.Length == 0))
            {
                return String.Empty;
            }

            //Calculate MD5 hash. This requires that the string is splitted into a byte[].
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] textToHash = Encoding.Default.GetBytes(Text);
            byte[] result = md5.ComputeHash(textToHash);

            //Convert result back to string.
            return System.BitConverter.ToString(result);
        }

        [STAThread]
        static void Main()
        {

            DLL test = new DLL();

            Console.WriteLine("Bing: " + test.BingTranslate("en", "es", "A mi padre le haría muy feliz que alguien le diera \"eso\" que brilla tanto. Si tú sabes a qué se refiere, tráeselas y él te recompensará"));
            Console.WriteLine("Google: " + test.GoogleTranslate("EN", "ES", "A mi padre le haría muy feliz que alguien le diera \"eso\" que brilla tanto. Si tú sabes a qué se refiere, tráeselas y él te recompensará"));
            Console.WriteLine("Yandex: " + test.YandexTranslate("en", "es", "A mi padre le haría muy feliz que alguien le diera \"eso\" que brilla tanto. Si tú sabes a qué se refiere, tráeselas y él te recompensará"));
            Console.WriteLine(test.Version());
            Console.Read();
        }

        public string Version()
        {
            return "2023";
        }

        public string GetLastTranslation()
        {
            return LastTranslation;
        }

        public void ClearLastTranslation()
        {
            LastTranslation = "";
        }
    }
}

