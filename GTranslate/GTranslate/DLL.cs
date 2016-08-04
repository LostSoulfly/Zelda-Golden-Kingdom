/*
    *Partial code taken from https://github.com/LuxintN/SelectAndTranslate
    *(Originally had a different Google translation code snippet, but his proved more versatile)
    *Things added in this order -> Google, MD5, Bing, Yandex.
    *So far, Bing seems to have the better translator of the three from my short tests.

    *Yandex = 10,000 calls per day (I've read, not had happen to me.)
    *Bing = you must sign up for (2,000,000 characters per month translated free.)
    *Google = Technically.. they charge for this service but we go around that by exploiting their web translator :x

    *Partial code taken from other projects along the way.
    *partial code by Dragoon/LostSoulFly. And Google.
    *Lots of shoddy modifications and scarce understanding of c# by Dragoon/LostSoulFly.
    *June 5, 2015
    
    * P.S. The Bing/Yandex keys that I have supplied were found randomly on the internet, so there shouldn't be any harm
    * in using them.. but you should get your own. It's free anyway.
*/

using System;
using System.Text;
using System.Net;
using System.IO;
using System.Threading.Tasks;
//using GTranslate;
using System.Runtime.InteropServices;
using System.Security.Cryptography;


namespace GTranslate
{

    [Guid("D5F88E95-8A27-4ae6-B6DE-6542A0FC7159")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _DLL
    {
        //Define the com objects that VB6 will interface with through the type library
        [DispId(1)]
        string GoogleTranslate(string LangTo, string LangFrom, string Text);

        [DispId(2)]
        string BingTranslate(string to, string from, string text, string clientID = "myBTranslate", string clientSecret = "zgQQfksRpj8H60LVHq4afeHtmVTldKrE7PQxRnqxOy4=");

        [DispId(3)]
        String YandexTranslate(string LangTo, string LangFrom, string Text, string key = "trnsl.1.1.20141229T202549Z.5f61901044d9ab3e.4d5c2d268897918f1adbfa15eb58b66d970ecbef");

        [DispId(4)]
        string GetMD5Hash(string Text);

        /*
        [DispId(5)]
        int getInt();

        [DispId(6)]
        void setInt(int i);
        */
    }

    [Guid("14FE32AD-4BF8-495f-AB4D-5C60BD463E59")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [ProgId("GTranslate.Translate")]
    public class DLL : _DLL
    {

        /*Testing the storage of variables while the VB exe is running.
         * The DLL remembers the int that was set even after different function
         * calls, so long as you do not destroy the DLL object in the project!
        private int test = 16;

        public int getInt()
        {
            return test;
        }

        public void setInt(int i)
        {
            test = i;
        }

        */

        public String GoogleTranslate(string LangTo, string LangFrom, string Text)
        {

            if ((Text == null) || (Text.Length == 0))
            {
                return String.Empty;
            }

            GoogleTranslator t = new GoogleTranslator();
            string translation;

            try
            {
                translation = t.Translate(Text, LangFrom, LangTo);
                return translation.Replace(System.Environment.NewLine, "\r\n");
            }
            catch (Exception ex)
            {
                //return "Error: " + ex.Message;
                return String.Empty;
            }

        }

        
        private BingTranslator.AdmAuthentication _auth;

        public string BingTranslate(string to, string from, string Text, string clientID = "myBTranslate", string clientSecret = "zgQQfksRpj8H60LVHq4afeHtmVTldKrE7PQxRnqxOy4=")
        {

            if ((Text == null) || (Text.Length == 0))
            {
                return String.Empty;
            }

            _auth = new BingTranslator.AdmAuthentication(clientID,clientSecret);

            string uri = "http://api.microsofttranslator.com/v2/Http.svc/Translate?text=" + System.Web.HttpUtility.UrlEncode(Text) + "&from=" + from + "&to=" + to;
            string authToken = "Bearer" + " " + _auth.GetAccessToken().access_token;

            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);
            httpWebRequest.Headers.Add("Authorization", authToken);

            WebResponse response = null;
            try
            {
                response = httpWebRequest.GetResponse();
                using (Stream stream = response.GetResponseStream())
                {
                    System.Runtime.Serialization.DataContractSerializer dcs = new System.Runtime.Serialization.DataContractSerializer(Type.GetType("System.String"));
                    string translation = (string)dcs.ReadObject(stream);
                    //Console.WriteLine("Translation for source text '{0}' from {1} to {2} is", text, from, to);
                    return translation.Replace(System.Environment.NewLine, "\r\n");
                }
            }
            catch (Exception ex)
            {
               //Console.WriteLine(ex.Message);
                //return "Error: " + ex.Message;
                return String.Empty;
            }
        }

        public String YandexTranslate(string LangTo, string LangFrom, string Text, string key = "trnsl.1.1.20141229T202549Z.5f61901044d9ab3e.4d5c2d268897918f1adbfa15eb58b66d970ecbef")
        {

            if ((Text == null) || (Text.Length == 0))
            {
                return String.Empty;
            }

            YandexTranslator y = new YandexTranslator();
            string translation;

            try
            {
                translation = y.Translate(Text, LangFrom, LangTo, key);
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

            //DLL test = new DLL();

            //Console.WriteLine ("Bing: " + test.BingTranslate("en", "es", "A mi padre le haría muy feliz que alguien le diera \"eso\" que brilla tanto. Si tú sabes a qué se refiere, tráeselas y él te recompensará"));
            //Console.WriteLine("Google: " + test.GoogleTranslate("EN", "ES", "A mi padre le haría muy feliz que alguien le diera \"eso\" que brilla tanto. Si tú sabes a qué se refiere, tráeselas y él te recompensará"));
            //Console.WriteLine("Yandex: " + test.YandexTranslate("en", "es", "A mi padre le haría muy feliz que alguien le diera \"eso\" que brilla tanto. Si tú sabes a qué se refiere, tráeselas y él te recompensará"));

            //Console.ReadKey();
        }
    }
}
