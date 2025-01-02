using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace TranslationServer
{

    public class Schema
    {
        public string md5 { get; set; }
        public string Translation { get; set; }
        public string Original { get; set; }

        public Schema(string md5, string translation, string original)
        {
            this.md5 = md5;
            this.Translation = translation;
            this.Original = original;
        }
    }

    public static class Database
    {
        private static readonly object _lock = new object();

        public static List<Schema> db;

        public static bool useDatabase = true;

        public static void Initialize()
        {
            LoadDb();
            //SaveDb();
        }

        public static string GetTranslation(string md5, string original)
        {
            if (!useDatabase) return null;

            var exists = Database.db.Find(e => e.md5 == md5);
            if (exists != null)
            {
                if (exists.Original == null)
                {
                    exists.Original = original;
                    RemoveTranslation(md5);
                    AddTranslation(md5, exists.Translation, original);
                }
                return exists.Translation;
            }
            return null;
        }

        public static void AddTranslation(string md5, string translation, string original)
        {
            if (string.IsNullOrEmpty(translation)) return;

            lock (_lock)
            {
                Database.db.Add(new Schema(md5, translation, original));
            }
            Console.WriteLine("Added translation to database: " + md5);
        }

        public static void RemoveTranslation(string md5)
        {
            lock (_lock)
            {
                var exists = Database.db.Find(e => e.md5 == md5);
                Database.db.Remove(exists);
            }
            Console.WriteLine("Removed translation from database: " + md5);
        }

        public static void SaveDb()
        {
            lock (_lock)
            {
                string jsonString = JsonConvert.SerializeObject(db, formatting: Formatting.Indented);
                File.WriteAllText("db.json", jsonString);
                Console.WriteLine("Saved " + Database.db.Count + " translations");
            }
        }

        private static void LoadDb()
        {
            lock (_lock)
            {
                try
                {
                    string json = File.ReadAllText("db.json");
                    if (json != null)
                    {
                        Database.db = JsonConvert.DeserializeObject<List<Schema>>(json);
                        Console.WriteLine("Loaded " + Database.db.Count + " translations");
                    }
                } catch
                {
                    Database.db = new List<Schema>();
                }
            }
        }
    }
}
