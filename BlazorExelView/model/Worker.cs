using MudBlazor;
using MudBlazor.Utilities;

namespace BlazorExelView.model
{
    public class Worker
    {
        public int ID { get; set; }
        public string IDView { get; set; }
        public string Name { get; set; }
        public string State { get; set; }
        public string Description { get; set; }
        public Dictionary<string, object> classBG { get; set; }

        public Worker(string name, int iD)
        {
            Name = name;
            ID = iD;
        }

        public void Init()
        {
            var color = string.IsNullOrEmpty(State) ? "rgba(244,67,54,0.5)" : "rgba(255,255,255,1)";
            classBG = new Dictionary<string, object>()
            {
                { "style",  $"background-color:{color}"},
            };
        }
    }
}
