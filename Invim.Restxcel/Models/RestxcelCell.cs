namespace Invim.Restxcel.Models
{
    public record RestxcelCell
    {
        public int? Row { get; set; }
        public int? Col { get; set; }
        public string Address { get; set; }
        public string Value { get; set; }
        public string FillColorRgb { get; set; }
        public string FontColorRgb { get; set; }
        public float? FontSize { get; set; }
        public string Font { get; set; }
        public bool? Bold { get; set; }
        public int? TextRotation { get; set; }
        public bool? ShrinkToFit { get; set; }
    }
}
