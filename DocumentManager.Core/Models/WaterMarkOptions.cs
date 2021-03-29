namespace DocumentManager.Core.Models
{
    public class WaterMarkOptions: DocumentOptions
    {
        public WaterMarkOptions()
        {
            Position = "bottomRight";
            ElementColor = "silver";
            ElementFontFamily = "font-family:\"Calibri\";font-size:1pt";
            Text = "CONFIDENTIAL";
            Opacity = ".5";
            ElementStyle = "position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:315;z-index:-251657216;" +
                           "mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;" +
                           "mso-position-vertical-relative:margin";
        }

        public string ElementStyle { get; set; }

        public string ElementColor { get; set; }

        public string ElementFontFamily { get; set; }

        public string Text { get; set; }

        public string Position { get; set; }

        public string Opacity { get; set; }
    }
}
