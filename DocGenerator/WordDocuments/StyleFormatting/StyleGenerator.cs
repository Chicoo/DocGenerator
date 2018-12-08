using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using StyleValues = DocumentFormat.OpenXml.Wordprocessing.StyleValues;

namespace DocumentGenerator.WordDocuments.StyleFormatting
{
    internal class StyleGenerator
    {
        internal static Styles DefaultStyle()
        {
            var element =
               new Styles(
                //Create the normal style
                    new Style(
                       new StyleName() { Val = "Normal" },
                       new PrimaryStyle()
                   ) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = OnOffValue.FromBoolean(true) },
                //Create the style for heading 1
                   new Style(
                        new StyleName() { Val = "heading 1" },
                        new BasedOn() { Val = "Normal" },
                        new NextParagraphStyle() { Val = "Normal" },
                        new UIPriority() { Val = 9 },
                        new PrimaryStyle(),
                        new Rsid() { Val = new HexBinaryValue() { Value = "00596592" } },
                        new StyleParagraphProperties(
                            new KeepNext(),
                            new KeepLines(),
                            new SpacingBetweenLines() {Before = "480", After = "0" },
                            new OutlineLevel() { Val = 0 }),
                        new StyleRunProperties(
                            new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi },
                            new Bold(),
                            new BoldComplexScript(),
                            new Color() { Val = "365F91", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" },
                            new FontSize() { Val = "28" },
                            new FontSizeComplexScript() { Val = "28" })
                    ) { Type = StyleValues.Paragraph, StyleId = "Heading1" },
                //Create the style for heading 2
                    new Style(
                        new StyleName() { Val = "heading 2" },
                        new BasedOn() { Val = "Normal" },
                        new NextParagraphStyle() { Val = "Normal" },
                        new LinkedStyle() { Val = "Heading2Char" },
                        new UIPriority() { Val = 9 },
                        new UnhideWhenUsed(),
                        new PrimaryStyle(),
                        new Rsid() { Val = new HexBinaryValue() { Value = "00CC42C0" } },
                        new StyleParagraphProperties(
                            new KeepNext(),
                            new KeepLines(),
                            new SpacingBetweenLines() { Before = "200", After = "0" },
                            new OutlineLevel() { Val = 1 }),
                        new StyleRunProperties(
                            new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi },
                            new Bold(),
                            new BoldComplexScript(),
                            new Color() { Val = "4F81BD", ThemeColor = ThemeColorValues.Accent1 },
                            new FontSize() { Val = "26" },
                            new FontSizeComplexScript() { Val = "26" })
                    ) { Type = StyleValues.Paragraph, StyleId = "Heading2" },
                //Create the style for heading 3
                    new Style(
                        new StyleName() { Val = "heading 3" },
                        new BasedOn() { Val = "Normal" },
                        new NextParagraphStyle() { Val = "Normal" },
                        new LinkedStyle() { Val = "Heading3Char" },
                        new UIPriority() { Val = 9 },
                        new UnhideWhenUsed(),
                        new PrimaryStyle(),
                        new Rsid() { Val = new HexBinaryValue() { Value = "00CC42C0" } },
                        new StyleParagraphProperties(
                            new KeepNext(),
                            new KeepLines(),
                            new SpacingBetweenLines() { Before = "200", After = "0" },
                            new OutlineLevel() { Val = 2 }),
                        new StyleRunProperties(
                            new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi },
                            new Bold(),
                            new BoldComplexScript(),
                            new Color() { Val = "4F81BD", ThemeColor = ThemeColorValues.Accent1 })
                    ) { Type = StyleValues.Paragraph, StyleId = "Heading3" },
                   //Create a bullet list style
                   new Style(
                       new StyleName() { Val = "List Paragraph" },
                       new BasedOn() { Val = "Normal" },
                       new UIPriority() { Val = 34 },
                       new PrimaryStyle(),
                       new Rsid() { Val = new HexBinaryValue() { Value = "00C87658" } },
                       new StyleParagraphProperties(
                           new Indentation() { Left = "720" },
                           new ContextualSpacing())
                   ) { Type = StyleValues.Paragraph, StyleId = "ListParagraph" }
               );

            return element;
        }
    }
}
