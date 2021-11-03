using Invim.Restxcel.Settings;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace Invim.Restxcel.Models
{
    public class RestxcelTemplateCollection
    {
        private readonly Dictionary<string, RestxcelTemplate> _templates;
        private readonly RestxcelSettings _settings;

        public RestxcelTemplateCollection(RestxcelSettings settings)
        {
            _templates = new();
            _settings = settings;

            ValidateSettings();
        }

        private void ValidateSettings()
        {
            if(_settings == null)
            {
                throw new InvalidDataException("no settings provided");
            }
            if(string.IsNullOrEmpty(_settings.TemplatesDirectory))
            {
                throw new InvalidDataException("no template directory specified");
            }
            if(!Directory.Exists(_settings.TemplatesDirectory))
            {
                throw new DirectoryNotFoundException($"template directory \"{_settings.TemplatesDirectory}\" not found");
            }
            if(!IsDirectoryWritable(_settings.TemplatesDirectory))
            {
                throw new InvalidDataException($"template directory \"{_settings.TemplatesDirectory}\" is not writeable");
            }
        }

        private bool IsDirectoryWritable(string dirPath, bool throwIfFails = false)
        {
            try
            {
                using (FileStream fs = File.Create(
                    Path.Combine(
                        dirPath,
                        Path.GetRandomFileName()
                    ),
                    1,
                    FileOptions.DeleteOnClose)
                )
                { }
                return true;
            }
            catch
            {
                if (throwIfFails)
                    throw;
                else
                    return false;
            }
        }

        public RestxcelTemplate this[string id] => FindById(id);

        private RestxcelTemplate FindById(string id) => _templates.ContainsKey(id) ? _templates[id] : LoadFromFileSystem(id);

        private RestxcelTemplate LoadFromFileSystem(string id)
        {
            var data = File.ReadAllBytes(GetTemplateDataFilePath(id));
            string metadata = File.ReadAllText(GetTemplateMetadataFilePath(id));
            var template = JsonConvert.DeserializeObject<RestxcelTemplate>(metadata);
            template.SetData(data);
            return template;
        }

        private string GetTemplateDataFilePath(string id) => Path.Combine(_settings.TemplatesDirectory, $"{id}.dat");
        private string GetTemplateMetadataFilePath(string id) => Path.Combine(_settings.TemplatesDirectory, $"{id}.metadata.json");

        private void SaveToFileSystem(RestxcelTemplate template)
        {
            File.WriteAllBytes(GetTemplateDataFilePath(template.Id), template.GetData());
            File.WriteAllText(GetTemplateMetadataFilePath(template.Id), JsonConvert.SerializeObject(template));
        }

        public RestxcelTemplate NewTemplate(byte[] data, bool permanent)
        {
            CheckData(data);
            RestxcelTemplate template = new();
            template.Permanent = permanent;
            template.SetData(data);
            if(permanent)
            {
                SaveToFileSystem(template);
            }
            _templates.Add(template.Id, template);
            return template;
        }

        private void CheckData(byte[] data)
        {
            if(data == null || data.Length < 1)
            {
                throw new InvalidDataException("no bytes received");
            }
            if(data.Length > 1048576)
            {
                throw new InvalidDataException("template size exceeded maximum of 1024 KB");
            }
            try
            {
                var tempPackage = new ExcelPackage(new MemoryStream(), new MemoryStream(data));
            }
            catch
            {
                throw new InvalidDataException("the file provided is not a valid Excel template");
            }
        }
    }
}
