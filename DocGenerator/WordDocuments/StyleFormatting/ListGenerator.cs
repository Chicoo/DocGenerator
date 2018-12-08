using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace DocumentGenerator.WordDocuments.StyleFormatting
{
    /// <summary>
    /// This class creates styles for different lists in OOXML.
    /// </summary>
    internal class ListGenerator
    {
        internal static void AddOOXMLBulletList(TextDocument document )
        {
            //document.Filename
        }

        /// <summary>
        /// Gets a random hexdecimal value as string.
        /// </summary>
        /// <returns></returns>
        public static string RandomIdGenerator()
        {
            int i;
            using (var randomNumber = new System.Security.Cryptography.RNGCryptoServiceProvider())
            {
                var randomId = new byte[8];
                randomNumber.GetNonZeroBytes(randomId);
                // If the system architecture is little-endian (that is, little end first), 
                // reverse the byte array. 
                if (BitConverter.IsLittleEndian) Array.Reverse(randomId);

                i = BitConverter.ToInt32(randomId, 0);
            }

            return i.ToString("X");
        }

        /// <summary>
        /// Creates the definition (abstract numbering) list for a bullet list.
        /// </summary>
        /// <returns>The definition for a bullet list.</returns>
        internal static AbstractNum CreateBulletList(int abstractNumberId)
        {
            //Create the id's
            var nsid = RandomIdGenerator();
            var templateCode = RandomIdGenerator();
            var templateCode1 = RandomIdGenerator();
            var templateCode2 = RandomIdGenerator();
            var templateCode3 = RandomIdGenerator();

            var bulletList = new AbstractNum(
                         new Nsid { Val = new HexBinaryValue { Value = nsid } },
                         new MultiLevelType { Val = MultiLevelValues.HybridMultilevel },
                         new TemplateCode { Val = new HexBinaryValue { Value = templateCode } },
                         new Level(
                             new StartNumberingValue { Val = 1 },
                             new NumberingFormat { Val = NumberFormatValues.Bullet },
                             new LevelText { Val = "·" },
                             new LevelJustification { Val = LevelJustificationValues.Left },
                             new PreviousParagraphProperties(
                                 new Indentation { Left = "720", Hanging = "360" }),
                             new NumberingSymbolRunProperties(
                                 new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" })
                         ) { LevelIndex = 0, TemplateCode = new HexBinaryValue { Value = templateCode1 } },
                         new Level(
                             new StartNumberingValue { Val = 1 },
                             new NumberingFormat { Val = NumberFormatValues.Bullet },
                             new LevelText { Val = "o" },
                             new LevelJustification { Val = LevelJustificationValues.Left },
                             new PreviousParagraphProperties(
                                 new Indentation { Left = "1440", Hanging = "360" }),
                             new NumberingSymbolRunProperties(
                                 new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" })
                         ) { LevelIndex = 1, TemplateCode = new HexBinaryValue { Value = templateCode2 } },
                         new Level(
                             new StartNumberingValue { Val = 1 },
                             new NumberingFormat { Val = NumberFormatValues.Bullet },
                             new LevelText { Val = "§" },
                             new LevelJustification { Val = LevelJustificationValues.Left },
                             new PreviousParagraphProperties(
                                 new Indentation { Left = "2160", Hanging = "360" }),
                             new NumberingSymbolRunProperties(
                                 new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" })
                         ) { LevelIndex = 2, TemplateCode = new HexBinaryValue { Value = templateCode3 }, Tentative = new OnOffValue(true) },
                         new Level(
                             new StartNumberingValue { Val = 1 },
                             new NumberingFormat { Val = NumberFormatValues.Bullet },
                             new LevelText { Val = "·" },
                             new LevelJustification { Val = LevelJustificationValues.Left },
                             new PreviousParagraphProperties(
                                 new Indentation { Left = "2880", Hanging = "360" }),
                             new NumberingSymbolRunProperties(
                                 new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" })
                         ) { LevelIndex = 3, TemplateCode = new HexBinaryValue { Value = templateCode1 }, Tentative = new OnOffValue(true) },
                         new Level(
                             new StartNumberingValue { Val = 1 },
                             new NumberingFormat { Val = NumberFormatValues.Bullet },
                             new LevelText { Val = "o" },
                             new LevelJustification { Val = LevelJustificationValues.Left },
                             new PreviousParagraphProperties(
                                 new Indentation { Left = "3600", Hanging = "360" }),
                             new NumberingSymbolRunProperties(
                                 new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" })
                         ) { LevelIndex = 4, TemplateCode = new HexBinaryValue { Value = templateCode2 }, Tentative = new OnOffValue(true) },
                         new Level(
                             new StartNumberingValue { Val = 1 },
                             new NumberingFormat { Val = NumberFormatValues.Bullet },
                             new LevelText { Val = "§" },
                             new LevelJustification { Val = LevelJustificationValues.Left },
                             new PreviousParagraphProperties(
                                 new Indentation { Left = "4320", Hanging = "360" }),
                             new NumberingSymbolRunProperties(
                                 new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" })
                         ) { LevelIndex = 5, TemplateCode = new HexBinaryValue { Value = templateCode3 }, Tentative = new OnOffValue(true) },
                         new Level(
                             new StartNumberingValue { Val = 1 },
                             new NumberingFormat { Val = NumberFormatValues.Bullet },
                             new LevelText { Val = "·" },
                             new LevelJustification { Val = LevelJustificationValues.Left },
                             new PreviousParagraphProperties(
                                 new Indentation { Left = "5040", Hanging = "360" }),
                             new NumberingSymbolRunProperties(
                                 new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" })
                         ) { LevelIndex = 6, TemplateCode = new HexBinaryValue { Value = templateCode1 }, Tentative = new OnOffValue(true) },
                         new Level(
                             new StartNumberingValue { Val = 1 },
                             new NumberingFormat { Val = NumberFormatValues.Bullet },
                             new LevelText { Val = "o" },
                             new LevelJustification { Val = LevelJustificationValues.Left },
                             new PreviousParagraphProperties(
                                 new Indentation { Left = "5760", Hanging = "360" }),
                             new NumberingSymbolRunProperties(
                                 new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" })
                         ) { LevelIndex = 7, TemplateCode = new HexBinaryValue { Value = templateCode2 }, Tentative = new OnOffValue(true) },
                         new Level(
                             new StartNumberingValue { Val = 1 },
                             new NumberingFormat { Val = NumberFormatValues.Bullet },
                             new LevelText { Val = "§" },
                             new LevelJustification { Val = LevelJustificationValues.Left },
                             new PreviousParagraphProperties(
                                 new Indentation { Left = "6480", Hanging = "360" }),
                             new NumberingSymbolRunProperties(
                                 new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" })
                         ) { LevelIndex = 8, TemplateCode = new HexBinaryValue { Value = templateCode3 }, Tentative = new OnOffValue(true) }
                     ) { AbstractNumberId = abstractNumberId };


            return bulletList;
        }

        internal static AbstractNum CreateNumberedList(int abstractNumberId)
        {
            //Create the id's
            var nsid = RandomIdGenerator();
            var templateCode = RandomIdGenerator();
            var templateCode1 = RandomIdGenerator();
            var templateCode2 = RandomIdGenerator();
            var templateCode3 = RandomIdGenerator();

            var numberedList =
                new AbstractNum(
                        new Nsid { Val = new HexBinaryValue { Value = nsid } },
                        new MultiLevelType { Val = MultiLevelValues.HybridMultilevel },
                        new TemplateCode { Val = new HexBinaryValue { Value = templateCode } },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "%1." },
                            new LevelJustification { Val = LevelJustificationValues.Left },
                            new PreviousParagraphProperties(
                                new Indentation { Left = "720", Hanging = "360" })
                        ) { LevelIndex = 0, TemplateCode = new HexBinaryValue { Value = templateCode1 } },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.LowerLetter },
                            new LevelText { Val = "%2." },
                            new LevelJustification { Val = LevelJustificationValues.Left },
                            new PreviousParagraphProperties(
                                new Indentation { Left = "1440", Hanging = "360" })
                        ) { LevelIndex = 1, TemplateCode = new HexBinaryValue { Value = templateCode2 } },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.LowerRoman },
                            new LevelText { Val = "%3." },
                            new LevelJustification { Val = LevelJustificationValues.Right },
                            new PreviousParagraphProperties(
                                new Indentation { Left = "2160", Hanging = "180" })
                        ) { LevelIndex = 2, TemplateCode = new HexBinaryValue { Value = templateCode3 } },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "%4." },
                            new LevelJustification { Val = LevelJustificationValues.Left },
                            new PreviousParagraphProperties(
                                new Indentation { Left = "2880", Hanging = "360" })
                        ) { LevelIndex = 3, TemplateCode = new HexBinaryValue { Value = templateCode1 } },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.LowerLetter },
                            new LevelText { Val = "%5." },
                            new LevelJustification { Val = LevelJustificationValues.Left },
                            new PreviousParagraphProperties(
                                new Indentation { Left = "3600", Hanging = "360" })
                        ) { LevelIndex = 4, TemplateCode = new HexBinaryValue { Value = templateCode2 }, Tentative = new OnOffValue(true) },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.LowerRoman },
                            new LevelText { Val = "%6." },
                            new LevelJustification { Val = LevelJustificationValues.Right },
                            new PreviousParagraphProperties(
                                new Indentation { Left = "4320", Hanging = "180" })
                        ) { LevelIndex = 5, TemplateCode = new HexBinaryValue { Value = templateCode3 }, Tentative = new OnOffValue(true) },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "%7." },
                            new LevelJustification { Val = LevelJustificationValues.Left },
                            new PreviousParagraphProperties(
                                new Indentation { Left = "5040", Hanging = "360" })
                        ) { LevelIndex = 6, TemplateCode = new HexBinaryValue { Value = templateCode1 }, Tentative = new OnOffValue(true) },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.LowerLetter },
                            new LevelText { Val = "%8." },
                            new LevelJustification { Val = LevelJustificationValues.Left },
                            new PreviousParagraphProperties(
                                new Indentation { Left = "5760", Hanging = "360" })
                        ) { LevelIndex = 7, TemplateCode = new HexBinaryValue { Value = templateCode2 }, Tentative = new OnOffValue(true) },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.LowerRoman },
                            new LevelText { Val = "%9." },
                            new LevelJustification { Val = LevelJustificationValues.Right },
                            new PreviousParagraphProperties(
                                new Indentation { Left = "6480", Hanging = "180" })
                        ) { LevelIndex = 8, TemplateCode = new HexBinaryValue { Value = templateCode3 }, Tentative = new OnOffValue(true) }
                    ) { AbstractNumberId = abstractNumberId };
                    

            return numberedList;
        }


        internal static Numbering CreateNumbering(NumberingDefinitionsPart numberingDefinitionsPart)
        {
            //For now do not add bullet and numbered list when there are already numberings defined.
            //Otherwise the orgiginal numbering will not work.
            var element = new Numbering();
            if (numberingDefinitionsPart != null)
            {
                element.Load(numberingDefinitionsPart);
            }
            else
            {
                var abstractNumberId = element.ChildElements.OfType<AbstractNum>().Count();
                element.Append(CreateBulletList(abstractNumberId));
                element.Append(CreateNumberedList(abstractNumberId + 1));    
            }
            
            return element;
        }
    }
}
