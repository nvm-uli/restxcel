using System.Collections.Generic;

namespace Invim.Restxcel.Models
{
    public class RestxcelDocumentCollection
    {
        private readonly Dictionary<string, RestxcelDocument> _documents;

        public RestxcelDocumentCollection()
        {
            _documents = new Dictionary<string, RestxcelDocument>();
        }

        private RestxcelDocument FindById(string id) => _documents[id];

        public RestxcelDocument this[string id] => FindById(id);

        public RestxcelDocument NewDocument(out string id, RestxcelTemplate template = null)
        {
            RestxcelDocument doc = new(template);
            id = doc.Id;
            _documents.Add(id, doc);
            return doc;
        }
    }
}
