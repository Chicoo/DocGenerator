using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OOXMLParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using AODL.Document.Content;
using AODL.Document;
using AODL.Document.Styles;
using AODL.Document.Content.Text;
using System.Collections;

namespace DocumentGenerator.WordDocuments
{
    /// <summary>
    /// Class for creating a bullet list in a document.
    /// </summary>
    public class BulletList : List
    {
        #region Constructors
        /// <summary>
        /// Creates a bullet list.
        /// </summary>
        public BulletList() : base()
        {
            _type = ListType.Bullet;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Gets a list in ODF format.
        /// </summary>
        /// <param name="document">The document to which the list belongs.</param>
        /// <returns>The list as an IContent document</returns>
        public override IContent GetODFList(IDocument document)
        {
            //Create a list
            AODL.Document.Content.Text.List list = new AODL.Document.Content.Text.List(document, "BL1", ListStyles.Bullet, "BL1P1");
            
            //Set the list level to 0 as this is the top of the list
            int currentLevel = 0;
            //Create a new list item
            AODL.Document.Content.Text.ListItem listItem = null;
            //Loop through the items
            for(int currentIndex =0; currentIndex < Items.Count; currentIndex++)
            {
                if (Items[currentIndex].Level > currentLevel)
                {
                    //Start a new list and append items.
                    if (listItem == null) listItem = new AODL.Document.Content.Text.ListItem(document);
                    listItem.Content.Add(GetODFSublist(Items[currentIndex], document, list, ref currentIndex));
                    currentIndex--;
                    //Skip the rest of the routine
                    continue;
                }
                if (Items[currentIndex].Level == currentLevel)
                {
                    //Create a new list item
                     listItem = new AODL.Document.Content.Text.ListItem(document);
                    //Create a paragraph	
                    var paragraph = ParagraphBuilder.CreateStandardTextParagraph(document);
                    //Add the text
                    foreach (var formatedText in CommonDocumentFunctions.ParseParagraphForODF(document, Items[currentIndex].Text))
                    {
                        paragraph.TextContent.Add(formatedText);
                    }

                    //Add paragraph to the list item
                    listItem.Content.Add(paragraph);

                    //Add the list item
                    list.Content.Add(listItem);
                }
            }

            return list;
        }
        #endregion

    }
}
