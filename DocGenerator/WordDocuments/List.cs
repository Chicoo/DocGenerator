using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OOXMLParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using AODL.Document.Content;
using AODL.Document;
using AODL.Document.Content.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentGenerator.WordDocuments
{
    #region ListItem
    /// <summary>
    /// This class creates a list item used in a bullet or numbered list.
    /// </summary>
    public struct ListItem
    {
        private readonly int _level;
        /// <summary>
        /// The level of the list item. This determines the number of bullet type for the item.
        /// </summary>
        public int Level { get { return _level; } }
        private readonly string _text;
        /// <summary>
        /// The text of the list item.
        /// </summary>
        public string Text { get { return _text; } }

        /// <summary>
        /// A list item used by a bullet or numbered list.
        /// </summary>
        /// <param name="level">The level of the item.</param>
        /// <param name="text">The test of the item.</param>
        public ListItem(int level, string text)
        {
            _level = level;
            _text = text;
        }
    }
    #endregion


    #region ListType
    /// <summary>
    /// The types for a list in a document
    /// </summary>
    public enum ListType
    {
        /// <summary>
        /// A bullet list
        /// </summary>
        Bullet,
        /// <summary>
        /// A numbered list.
        /// </summary>
        Number
    }
    #endregion ListType


    /// <summary>
    /// This is the toplevel class for a list.
    /// </summary>
    public abstract class List : Paragraph
    {
        #region Properties
        private List<ListItem> listItems;
        /// <summary>
        /// List of items in the bullet list
        /// </summary>
        public IList<ListItem> Items
        {
            get { return listItems.AsReadOnly(); }
        }

        protected ListType _type;

        /// <summary>
        /// The type of the list.
        /// </summary>
        public ListType Type => _type;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a bullet list.
        /// </summary>
        protected List()
        {
            listItems = new List<ListItem>();
        }
        #endregion

        #region Protected Methods
        /// <summary>
        /// Creates sublists for nested lists.
        /// </summary>
        /// <param name="firstItem">The item to start the list with.</param>
        /// <param name="document">The document to which the sublist belongs.</param>
        /// <param name="currentIndex">The index of the item.</param>
        /// <returns>An ODF list</returns>
        protected AODL.Document.Content.Text.List GetODFSublist(ListItem firstItem, IDocument document, AODL.Document.Content.Text.List parentList, ref int currentIndex)
        {
            //Create a list
            AODL.Document.Content.Text.List list = new AODL.Document.Content.Text.List(document, parentList);

            //Set the level of the list
            int currentLevel = firstItem.Level;
            //Create a new list item
            AODL.Document.Content.Text.ListItem listItem = null;
            //Loop through the items
            while (currentIndex < Items.Count)
            {
                if (Items[currentIndex].Level > currentLevel)
                {
                    //Start a new list and append items.
                    listItem.Content.Add(GetODFSublist(Items[currentIndex], document, list, ref currentIndex));
                    continue;
                }
                if (Items[currentIndex].Level < currentLevel)
                {
                    //Stop this list.
                    break;
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

                    currentIndex++;
                }
            }

            return list;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Adds a list item to use in the list.
        /// </summary>
        /// <param name="level">The level of the item.</param>
        /// <param name="text">The text of the item.</param>
        public void AddItem(int level, string text)
        {
            ListItem item = new ListItem(level, text);
            listItems.Add(item);
        }

        /// <summary>
        /// Gets a list in OOXMLFormat.
        /// This is returned as a list as an object may have multiple paragraphs.
        /// </summary>
        /// <param name="listId">The id of the listvin the numbering.</param>
        /// <returns>List of lisr in OOXML Format.</returns>
        public virtual List<OOXMLParagraph> GetOOXMLList(int listId)
        {
            List<OOXMLParagraph> list = new List<OOXMLParagraph>();
            

            foreach (var item in Items)
            {
                var element = new OOXMLParagraph();

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "ListParagraph" };

                NumberingProperties numberingProperties1 = new NumberingProperties();
                NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = item.Level };
                NumberingId numberingId1 = new NumberingId() { Val = listId };

                numberingProperties1.Append(numberingLevelReference1);
                numberingProperties1.Append(numberingId1);

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(numberingProperties1);

                Run run1 = new Run();
                Text text = new Text
                {
                    Text = item.Text
                };

                run1.Append(text);

                element.Append(paragraphProperties1);
                element.Append(run1);
                list.Add(element);
            }
            return list;
        }

        /// <summary>
        /// Gets a list in ODF format.
        /// </summary>
        /// <param name="document">The document to which the list belongs.</param>
        /// <returns>The list as an IContent document</returns>
        public abstract IContent GetODFList(IDocument document);
        #endregion

    }

}
