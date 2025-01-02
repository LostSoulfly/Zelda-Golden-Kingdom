using FileHelpers;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static Client;

namespace TranslationServer
{
    public class ItemTranslator
    {


        public ItemTranslator()
        {
            Database.Initialize();
            Database.useDatabase = true;
        }

        public void TranslateFile(string filePath)
        {

        }

        public void Translate(string l)
        {
            try
            {

                Random rnd = new Random();
                MessageType e = (MessageType)Enum.ToObject(typeof(MessageType), rnd.Next(0, 2));

                string md5 = l.Substring(2, 32);
                l = l.Substring(34).Replace("\r\n", Environment.NewLine);

                string cache = Database.GetTranslation(md5, l);

                if (cache == null)
                {
                    Console.WriteLine($"{md5} - {e}: " + l);

                    string tt = "";
                    switch (e)
                    {
                        case MessageType.Google:
                            tt = Translator.Google(l).GetAwaiter().GetResult();
                            break;
                        case MessageType.Bing:
                            tt = Translator.Bing(l).GetAwaiter().GetResult();
                            break;
                        case MessageType.Yandex:
                            tt = Translator.Yandex(l).GetAwaiter().GetResult();
                            break;
                        default:
                            break;
                    }
                    Console.WriteLine($"{md5} - {e} result: " + tt);
                    Database.AddTranslation(md5, tt, l);
                }
                else
                {
                    Console.WriteLine($"{md5} - CACHE: " + cache);

                }
            } catch (Exception e)
                {
                Console.WriteLine(e.Message);
            }
        }
    }
}
