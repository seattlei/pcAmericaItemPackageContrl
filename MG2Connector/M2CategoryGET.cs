using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MG2Connector
{
    public class ChildrenData
    {

        [JsonProperty("id")]
        public int Id { get; set; }

        [JsonProperty("parent_id")]
        public int ParentId { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("is_active")]
        public bool IsActive { get; set; }

        [JsonProperty("position")]
        public int Position { get; set; }

        [JsonProperty("level")]
        public int Level { get; set; }

        [JsonProperty("product_count")]
        public int ProductCount { get; set; }

        [JsonProperty("children_data")]
        public IList<ChildrenData> ChildrenData1 { get; set; }
    }

    public class ChildrenData1
    {

        [JsonProperty("id")]
        public int Id { get; set; }

        [JsonProperty("parent_id")]
        public int ParentId { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("is_active")]
        public bool IsActive { get; set; }

        [JsonProperty("position")]
        public int Position { get; set; }

        [JsonProperty("level")]
        public int Level { get; set; }

        [JsonProperty("product_count")]
        public int ProductCount { get; set; }

        [JsonProperty("children_data")]
        public IList<ChildrenData> ChildrenData { get; set; }
    }

    public class M2CategoryGET
    {

        [JsonProperty("id")]
        public int Id { get; set; }

        [JsonProperty("parent_id")]
        public int ParentId { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("is_active")]
        public bool IsActive { get; set; }

        [JsonProperty("position")]
        public int Position { get; set; }

        [JsonProperty("level")]
        public int Level { get; set; }

        [JsonProperty("product_count")]
        public int ProductCount { get; set; }

        [JsonProperty("children_data")]
        public IList<ChildrenData> ChildrenData { get; set; }
    }



}
