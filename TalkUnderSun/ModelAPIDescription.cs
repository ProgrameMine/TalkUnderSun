using System.ComponentModel;

namespace TalkUnderSun
{
    public class ModelAPIDescription
    {
        [Description("API")]
        public string API { get; set; }

        [Description("Description")]
        public string Description { get; set; }

        public string Type { get; set; }

        public string FunctionAddr { get; set; }

        public string Function { get; set; }

        public string SystemName { get; set; }
    }
}
