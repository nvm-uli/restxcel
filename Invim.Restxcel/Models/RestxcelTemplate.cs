using System;

namespace Invim.Restxcel.Models
{
    public class RestxcelTemplate
    {
        private byte[] _data = null;

        public string Id { get; set; }
        public bool Permanent { get; set; }
        public DateTime Created { get; set; }

        public RestxcelTemplate()
        {
            Id = Guid.NewGuid().ToString();
            Created = DateTime.Now;
        }

        public void SetData(byte[] data) => _data = data;
        public byte[] GetData() => _data;
    }
}
