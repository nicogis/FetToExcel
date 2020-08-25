//-----------------------------------------------------------------------
// <copyright file="Config.cs" company="Studio A&T s.r.l.">
//     Author: nicogis
//     Copyright (c) Studio A&T s.r.l. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace FetToExcel
{
    using Newtonsoft.Json;

    public class Config
    {
        [JsonProperty("pathTeachersXml")]
        public string PathTeachersXml { get; set; }

        [JsonProperty("pathTemplateExcel")]
        public string PathTemplateExcel { get; set; }

        [JsonProperty("pathOuputExcel")]
        public string PathOuputExcel { get; set; }

        [JsonProperty("openExcel")]
        public bool OpenExcel { get; set; }

        [JsonProperty("cellStartTeachers")]
        public string CellStartTeachers { get; set; }
    }
}