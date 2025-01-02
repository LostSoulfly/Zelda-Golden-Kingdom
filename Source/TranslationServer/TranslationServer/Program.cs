// See https://aka.ms/new-console-template for more information
using GTranslate.Translators;
using System.Net;
using System.Net.Sockets;
using TranslationServer;
using static System.Net.Mime.MediaTypeNames;



//var itemTrans = new ItemTranslator();

//itemTrans.TranslateFile(@"C:\Users\Dragoon\Desktop\Zelda\Source\Server\data\items\item800.dat");

//System.Environment.Exit(0);

var svr = new Server("127.0.0.1");
await svr.Run();
Database.SaveDb();

class Server
{
    IPEndPoint ipep;
    TcpListener listener;
    bool Running;
    List<Client> clients;
    private CancellationTokenSource cts;

    public Server(string host)
    {
        Database.Initialize();
        Database.useDatabase = true;
        IPAddress ip = IPAddress.Parse(host);
        ipep = new(ip, 6969);
        Running = false;
        clients = new();

        this.cts = new CancellationTokenSource();
    }

    public void Stop()
    {
        Running = false;
        cts.Cancel();
    }
    public async Task Run()
    {
        listener = new(ipep);
        listener.Start();
        Running = true;
        Console.WriteLine("Translation server running on port " + ipep.Port);
        while (Running)
        {
            var c = await listener.AcceptTcpClientAsync(cts.Token);
            var client = new Client(c);
            clients.Add(client);
            var clientTask = client.Run(); //don't await
            clientTask.ContinueWith(t => clients.Remove(client));
            clientTask.ContinueWith(t => Database.SaveDb());
        }

        Database.SaveDb();
    }
}

class Client
{
    TcpClient client;
    NetworkStream stream;

    public enum MessageType
    {
        Google = 10,
        Bing = 11,
        Yandex = 12,
        Random = 13,
        DisableCache = 14,
        EnableCache = 15

    }

    public Client(TcpClient client)
    {
        this.client = client;
        stream = client.GetStream();

    }

    public async Task Run()
    {
        var r = new StreamReader(stream);
        var w = new StreamWriter(stream);
        while (true)
        {

            var l = await r.ReadLineAsync();

            if (l == null)
                break;

            if (l.Length > 0)
            {
                Thread.Sleep(50);
                try
                {
                    int t = int.Parse(l.Substring(0, 2));
                    MessageType e;

                    switch ((MessageType)t)
                    {
                        case MessageType.Google:
                        case MessageType.Bing:
                        case MessageType.Yandex:
                            e = (MessageType)Enum.ToObject(typeof(MessageType), t);
                            break;
                        case MessageType.Random:
                            Random rnd = new Random();
                            e = (MessageType)Enum.ToObject(typeof(MessageType), rnd.Next(0, 2));
                            break;
                        case MessageType.DisableCache:
                            Database.useDatabase = false;
                            await w.WriteLineAsync("99Cache Disabled");
                            Console.WriteLine("Cache Disabled.");
                            await w.FlushAsync();
                            continue;
                            break;
                        case MessageType.EnableCache:
                            Database.useDatabase = true;
                            await w.WriteLineAsync("99Cache Enabled");
                            Console.WriteLine("Cache Enabled.");
                            await w.FlushAsync();
                            continue;
                            break;
                        default:

                            await w.WriteLineAsync("99Unknown command.");
                            Console.WriteLine("Unknown command:" + l);
                            await w.FlushAsync();
                            continue;
                            break;
                    }


                    if (t < 13 && t >= 10)
                    {
                        e = (MessageType)Enum.ToObject(typeof(MessageType), t);
                    } else if (t == 14)
                    {
                        Random rnd = new Random();
                        e = (MessageType)Enum.ToObject(typeof(MessageType), rnd.Next(0, 2));
                    }

                    string md5 = l.Substring(2, 32);
                    l = l.Substring(34).Replace("\r\n", Environment.NewLine);

                    string cache = Database.GetTranslation(md5, l);

                    if (cache == null)
                    {
                        Console.WriteLine($"{this.client.Client.RemoteEndPoint} - {md5} - {e}: " + l);

                        string tt = "";
                        switch (e)
                        {
                            case MessageType.Google:
                                tt = await Translator.Google(l);
                                break;
                            case MessageType.Bing:
                                tt = await Translator.Bing(l);
                                break;
                            case MessageType.Yandex:
                                tt = await Translator.Yandex(l);
                                break;
                            default:
                                break;
                        }
                        Console.WriteLine($"{this.client.Client.RemoteEndPoint} - {md5} - {e} result: " + tt);
                        Database.AddTranslation(md5, tt, l);
                        await w.WriteLineAsync(md5 + tt);
                    } else
                    {
                        Console.WriteLine($"{this.client.Client.RemoteEndPoint} - {md5} - CACHE: " + cache);
                        await w.WriteLineAsync(md5 + cache);

                    }
                } catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    this.client.Close();
                }
            }
            else
            {

                await w.WriteLineAsync("");
            }

            await w.FlushAsync();

        }
    }

}