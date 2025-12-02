using SuperOffice.ErpSync.ConnectorWS;
using SuperOffice.CRM;

namespace ConnectorService.Extensions
{
    /// <summary>
    /// New implementation of IConfigDataStore which stores config data in a specified folder on server.
    /// <para/>
    /// Replaces the default Isolated storage implementation (which tends to not work on NETWORK SERVICE accounts).
    /// </summary>
    /// <remarks>
    /// Folder is specified in appsetings.json
    /// </remarks>
    [ConfigDataStore("FileStorage", int.MaxValue / 2)]
    public class FileConfigDataStore : IConfigDataStore
    {
        readonly string _baseDirectory;

        public FileConfigDataStore()
        {
            _baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        }

        /// <summary>
        /// Remove all trace of configuration from data store.
        /// </summary>
        /// <param name="key">Connection id</param>
        public void DeleteData(string key)
        {
            CheckConfigFolder();
            var path = Path.Combine(_baseDirectory, key);
            if (Directory.Exists(path))
                Directory.Delete(path, true);
        }


        /// <summary>
        /// Persist data to storage, taking care to only store data listed in members collection.  Passwords should be stored in encrypted format.
        /// </summary>
        /// <param name="key">Connection id</param>
        /// <param name="members">Collection of field names and their types.</param>
        /// <param name="data">Collection of field names and CultureFormatted values to be saved to storage.</param>
        public void SaveData(string key, Dictionary<string, FieldMetadataTypeInfo> members, Dictionary<string, string> data)
        {
            CheckConfigFolder();

            var path = Path.Combine(_baseDirectory, key);
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            foreach (var oldFile in Directory.GetFiles(path))
                File.Delete(oldFile);

            foreach (var item in data)
            {
                if (!members.TryGetValue(item.Key, out var fieldType))
                    throw new Exception(string.Format("Cannot store undeclared data value '{0}'", item.Key));
                var itemPath = Path.Combine(path, item.Key + ".txt");
                var dataWriter = new StreamWriter(new FileStream(itemPath, FileMode.CreateNew));
                dataWriter.Write(fieldType == FieldMetadataTypeInfo.Password ? ReversibleEncryptedString.Encrypt(item.Value) : item.Value);
                dataWriter.Close();
            }
        }


        /// <summary>
        /// Retrieve data from storage, taking care to only fill the data collection with fields listed in the members collection.  Passwords should be decrypted before being placed in data.
        /// </summary>
        /// <param name="key">Connection id</param>
        /// <param name="members">Collection of field names and their types.</param>
        /// <param name="data">Collection of field names and CultureFormatted values - to be loaded from storage.</param>
        public void LoadData(string key, Dictionary<string, FieldMetadataTypeInfo> members, Dictionary<string, string> data)
        {
            CheckConfigFolder();

            var path = Path.Combine(_baseDirectory, key);
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            if (Directory.GetFiles(path).Length == 0)
                throw new Exception(string.Format("No data found for connection key '{0}'", key));

            foreach (var member in members)
            {
                var itemPath = Path.Combine(path, member.Key + ".txt");
                var dataReader = new StreamReader(new FileStream(itemPath, FileMode.Open));
                var raw = dataReader.ReadToEnd();
                dataReader.Close();

                if (member.Value == FieldMetadataTypeInfo.Password)
                    raw = ReversibleEncryptedString.Decrypt(raw);

                data[member.Key] = raw;
            }
        }

        private void CheckConfigFolder()
        {
            if (!Directory.Exists(_baseDirectory))
                Directory.CreateDirectory(_baseDirectory);
        }

    }
}
