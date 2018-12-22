/*************************************************************************
 *
 * DO NOT ALTER OR REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER
 * 
 * Copyright 2008 Sun Microsystems, Inc. All rights reserved.
 * 
 * Use is subject to license terms.
 * 
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0. You can also
 * obtain a copy of the License at http://odftoolkit.org/docs/license.txt
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * 
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 ************************************************************************/

using System;
using System.IO;
using System.Drawing;
using System.Xml;
using AODL.Document.Styles;
using AODL.Document.Content.Text;
using AODL.Document.Content;
using AODL.Document;
using AODL.Document.Content.EmbedObjects ;
using AODL.Document.Content.Charts ;
using AODL.Document.Exceptions;

namespace AODL.Document.Content.Draw
{
	public class AODLGraphicException : AODLException
	{
		public AODLGraphicException(string message, Exception e)
			:base (message, e)
		{
			
		}
	}
	
	
	/// <summary>
	/// Frame represent graphic resp. a draw container.
	/// </summary>
	public class Frame : IContent, IContentContainer
	{
		/// <summary>
		/// Gets or sets the frame style.
		/// </summary>
		/// <value>The frame style.</value>
		public FrameStyle FrameStyle
		{
			get { return (FrameStyle)Style; }
			set { Style = (IStyle) value; }
		}

		private string _realgraphicname;
		/// <summary>
		/// Gets the name of the real graphic.
		/// </summary>
		/// <value>The name of the real graphic.</value>
		public string RealGraphicName
		{
			get { return _realgraphicname; }
			set { _realgraphicname = value; }
		}

		private string _graphicSourcePath;
		/// <summary>
		/// Gets or sets the graphic source path.
		/// </summary>
		/// <value>The graphic source path.</value>
		public string GraphicSourcePath
		{
			get { return _graphicSourcePath; }
			set { _graphicSourcePath = value; }
		}

		private int _heightInPixel;
		/// <summary>
		/// Gets the image height in pixel.
		/// </summary>
		/// <value>The height in pixel.</value>
		public int HeightInPixel
		{
			get { return _heightInPixel; }
		}

		private int _widthInPixel;
		/// <summary>
		/// Gets the image width in pixel.
		/// </summary>
		/// <value>The width in pixel.</value>
		public int WidthInPixel
		{
			get { return _widthInPixel; }
		}

		private double _height;
		/// <summary>
		/// Gets the frame height in cm or inch depending on which format is used in the current document.
		/// </summary>
		/// <value>The height.</value>
		public double Height
		{
			get { return _height; }
		}

		private double _width;
		/// <summary>
		/// Gets the frame width in cm or inch depending on which format is used in the current document.
		/// </summary>
		/// <value>The height.</value>
		public double Width
		{
			get { return _width; }
		}

		private string _measurementFormat;
		/// <summary>
		/// Gets the measurement format. This will depent on the current document.
		/// Possible values are cm or in (inch).
		/// </summary>
		/// <value>The measurement format.</value>
		public string MeasurementFormat
		{
			get { return _measurementFormat; }
		}

		private int _dPI_Y;
		/// <summary>
		/// Gets the image vertical resulotion.
		/// </summary>
		/// <value>The vertical resulotion.</value>
		public int DPI_Y
		{
			get { return _dPI_Y; }
		}

		private int _dPI_X;
		/// <summary>
		/// Gets the image horizontal resulotion.
		/// </summary>
		/// <value>The horízontal resulotion.</value>
		public int DPI_X
		{
			get { return _dPI_X; }
		}

		public string AlternateText 
		{
			get 
			{
				return _node.InnerText;
			}
			set
			{
				_node.InnerText = value;
			}
		}
		
		/// <summary>
		/// Gets or sets the draw name.
		/// </summary>
		/// <value>The name of the graphic.</value>
		public string DrawName
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:name",
				                                         Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:name",
				                                         Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("name", value, "draw");
				_node.SelectSingleNode("@draw:name",
				                            Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the type of the anchor. e.g paragraph
		/// </summary>
		/// <value>The type of the anchor.</value>
		public string AnchorType
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@text:anchor-type",
				                                         Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@text:anchor-type",
				                                         Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("anchor-type", value, "text");
				_node.SelectSingleNode("@text:anchor-type",
				                            Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the z index.
		/// </summary>
		/// <value>The index of the Z.</value>
		public string ZIndex
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:z-index",
				                                         Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:z-index",
				                                         Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("z-index", value, "draw");
				_node.SelectSingleNode("@draw:z-index",
				                            Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the width of the frame. e.g 2.98cm
		/// </summary>
		/// <value>The width of the graphic.</value>
		public string SvgWidth
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@svg:width",
				                                         Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:width",
				                                         Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("width", value, "svg");
				_node.SelectSingleNode("@svg:width",
				                            Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the height of the frame. e.g 3.00cm
		/// </summary>
		/// <value>The height of the graphic.</value>
		public string SvgHeight
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@svg:height",
				                                         Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:height",
				                                         Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("height", value, "svg");
				_node.SelectSingleNode("@svg:height",
				                            Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the horizontal position where
		/// the hosted drawing e.g Graphic should be
		/// anchored.
		/// </summary>
		/// <example>myFrame.SvgX = "1.5cm"</example>
		/// <value>The SVG X.</value>
		public string SvgX
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@svg:x",
				                                         Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:x",
				                                         Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("x", value, "svg");
				_node.SelectSingleNode("@svg:x",
				                            Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the vertical position where
		/// the hosted drawing e.g Graphic should be
		/// anchored.
		/// </summary>
		/// <example>myFrame.SvgY = "1.5cm"</example>
		/// <value>The SVG Y.</value>
		public string SvgY
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@svg:y",
				                                         Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@svg:y",
				                                         Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("y", value, "svg");
				_node.SelectSingleNode("@svg:y",
				                            Document.NamespaceManager).InnerText = value;
			}
		}

		private Guid _graphicIdentifier;
		/// <summary>
		/// Gets or sets the graphic identifier.
		/// </summary>
		/// <value>The graphic identifier.</value>
		public Guid GraphicIdentifier
		{
			get { return _graphicIdentifier; }
			set { _graphicIdentifier = value; }
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Frame"/> class.
		/// </summary>
		/// <param name="document">The textdocument.</param>
		/// <param name="stylename">The stylename.</param>
		public Frame(IDocument document, string stylename)
		{
			Document			= document;

			NewXmlNode();
			InitStandards();

			if (stylename != null)
			{
				Style				= (IStyle)new FrameStyle(Document, stylename);
				StyleName			= stylename;
				Document.Styles.Add(Style);
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Frame"/> class.
		/// </summary>
		/// <param name="document">The textdocument.</param>
		/// <param name="stylename">The stylename.</param>
		/// <param name="drawName">The  draw name.</param>
		/// <param name="graphicfile">The graphicfile.</param>
		public Frame(IDocument document, string stylename, string drawName, string graphicfile)
		{
			Document			= document;
			NewXmlNode();
			InitStandards();

			StyleName			= stylename;
			//			this.AnchorType			= "paragraph";

			DrawName			= drawName;
			GraphicSourcePath	= graphicfile;

			_graphicIdentifier = Guid.NewGuid();
			_realgraphicname	= _graphicIdentifier.ToString()
				+ LoadImageFromFile(graphicfile);
            Graphic graphic = new Graphic(Document, this, _realgraphicname)
            {
                GraphicRealPath = GraphicSourcePath
            };
            Content.Add(graphic);
			Style				= (IStyle)new FrameStyle(Document, stylename);
			Document.Styles.Add(Style);
		}

		/// <summary>
		/// Inits the standards.
		/// </summary>
		private void InitStandards()
		{
			//Todo: FrameBuilder
			AnchorType				= "paragraph";

			Content				= new ContentCollection();
			Content.Inserted		+= Content_Inserted;
			Content.Removed		+= Content_Removed;
		}

		/// <summary>
		/// Create a new XmlNode.
		/// </summary>
		private void NewXmlNode()
		{
			Node		= Document.CreateNode("frame", "draw");
		}

		/// <summary>
		/// Create a XmlAttribute for propertie XmlNode.
		/// </summary>
		/// <param name="name">The attribute name.</param>
		/// <param name="text">The attribute value.</param>
		/// <param name="prefix">The namespace prefix.</param>
		private void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			Node.Attributes.Append(xa);
		}

		/// <summary>
		/// Loads the image from file.
		/// </summary>
		/// <param name="graphicfilename">The graphicfilename.</param>
		public string LoadImageFromFile(string graphicfilename)
		{
			Image image	= null;
			try
			{
				image = Image.FromFile(graphicfilename);
			}
			catch (OutOfMemoryException e)
			{
				throw new AODLGraphicException(string.Format(
					"file '{0}' contains an unrecognized image", graphicfilename), e);
			}
			
			Graphics graphics   = null;
			
			try
			{
				graphics = Graphics.FromImage(image);
			}
			catch (Exception e)
			{
				throw new AODLGraphicException(string.Format(
					"file '{0}' contains an unrecognized image. I could not" +
					" create graphics object from image {1}", graphicfilename, image), e);
			}
			
			_widthInPixel  = image.Width;
			_heightInPixel = image.Height;
			_dPI_X			= (int)graphics.DpiX;
			_dPI_Y			= (int)graphics.DpiY;
			_measurementFormat = "cm";
			_width		= AODL.Document.Helper.SizeConverter.GetWidthInCm(image.Width, _dPI_X);
			_height	= AODL.Document.Helper.SizeConverter.GetHeightInCm(image.Height, _dPI_Y);
			if (SvgHeight == null && SvgWidth == null)
			{
				SvgHeight	= _height.ToString("F3")+ _measurementFormat;
				if (SvgHeight.IndexOf(",", 0)  > -1)
					SvgHeight = SvgHeight.Replace(",",".");
				SvgWidth	= _width.ToString("F3")+ _measurementFormat;
				if (SvgWidth.IndexOf(",", 0)  > -1)
					SvgWidth = SvgWidth.Replace(",",".");
			}
			else
			{
				// This should only reached, if the file is loading by a importer.
				if (AODL.Document.Helper.SizeConverter.IsCm(SvgWidth))
				{
					_measurementFormat = "cm";
				}
				else
				{
					_measurementFormat = "in";
				}
			}
			graphics.Dispose();
			image.Dispose();

			return new FileInfo(graphicfilename).Name;
		}

		/// <summary>
		/// Content_s the inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Inserted(int index, object value)
		{
			if (value is Graphic)
			{
				if (((Graphic)value).Frame == null)
				{
					((Graphic)value).Frame = this;
					if (((Graphic)value).GraphicRealPath != null
					   && GraphicSourcePath == null)
						GraphicSourcePath = ((Graphic)value).GraphicRealPath;
					Node.AppendChild(((IContent)value).Node);
				}
				if (((IContent)value).Node != null)
				{
					if (((IContent)value).Node.ParentNode == null
					   || !((IContent)value).Node.ParentNode.Equals(Node))
					{
						Node.AppendChild(((IContent)value).Node);
					}
				}
			}
			else if (value is EmbedObject )
			{
				if (((EmbedObject)value).ObjectType =="chart")
				{
					if (Document.IsLoadedFile&&!((Chart)value).IsNewed )
					{
						Node .AppendChild (((Chart)value).ParentNode );
					}
					else
					{
						string  objectLink = "."+@"/"+((EmbedObject)value).ObjectName;
						
						Node .AppendChild (((Chart)value).CreateParentNode (objectLink) );
					}
				}
			}
			else
			{
				if (((IContent)value).Node != null)
				{
					if (((IContent)value).Node.ParentNode == null
					   || !((IContent)value).Node.ParentNode.Equals(Node))
					{
						Node.AppendChild(((IContent)value).Node);
					}
				}
			}
		}

		/// <summary>
		/// Content_s the removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Removed(int index, object value)
		{
			Node.RemoveChild(((IContent)value).Node);
			//if graphic remove it
			if (value is Graphic)
				if (Document.Graphics.Contains(value as Graphic))
				Document.Graphics.Remove(value as Graphic);
		}

		#region IContent Member
		/// <summary>
		/// Gets or sets the name of the style.
		/// </summary>
		/// <value>The name of the style.</value>
		public string StyleName
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@draw:style-name",
				                                         Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@draw:style-name",
				                                         Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("style-name", value, "draw");
				_node.SelectSingleNode("@draw:style-name",
				                            Document.NamespaceManager).InnerText = value;
			}
		}

		private IDocument _document;
		/// <summary>
		/// Every object (typeof(IContent)) have to know his document.
		/// </summary>
		/// <value></value>
		public IDocument Document
		{
			get
			{
				return _document;
			}
			set
			{
				_document = value;
			}
		}

		private IStyle _style;
		/// <summary>
		/// A Style class wich is referenced with the content object.
		/// If no style is available this is null.
		/// </summary>
		/// <value></value>
		public IStyle Style
		{
			get
			{
				return _style;
			}
			set
			{
				StyleName	= value.StyleName;
				_style = value;
			}
		}

		private XmlNode _node;
		/// <summary>
		/// Gets or sets the node.
		/// </summary>
		/// <value>The node.</value>
		public XmlNode Node
		{
			get { return _node; }
			set { _node = value; }
		}

		#endregion

		#region IHtml Member
		//
		//		/// <summary>
		//		/// Return the content as Html string
		//		/// </summary>
		//		/// <returns>The html string</returns>
		//		public string GetHtml()
		//		{
		//			string html			= "";
		//			if (this.Content.Count > 0)
		//				if (this.Content[0] is Graphic)
		//					html			= ((Graphic)this.Content[0]).GetHtml();
		//
		//			foreach(IContent content in this.Content)
		//				if (content is IHtml)
		//					html		+= ((IHtml)content).GetHtml();
		//
		//			return html;
		//		}

		#endregion


		#region IContentContainer Member

		private ContentCollection _content;
		/// <summary>
		/// Gets or sets the content.
		/// </summary>
		/// <value>The content.</value>
		public ContentCollection Content
		{
			get
			{
				return _content;
			}
			set
			{
				_content = value;
			}
		}

		#endregion

		public void CreatePubAttr(string name, string text, string prefix)
		{
			CreateAttribute( name,text, prefix);
		}
	}
}

/*
 * $Log: Frame.cs,v $
 * Revision 1.4  2008/04/29 15:39:44  mt
 * new copyright header
 *
 * Revision 1.3  2008/02/08 07:12:19  larsbehr
 * - added initial chart support
 * - several bug fixes
 *
 * Revision 1.1  2007/02/25 08:58:32  larsbehr
 * initial checkin, import from Sourceforge.net to OpenOffice.org
 *
 * Revision 1.6  2007/02/13 17:58:47  larsbm
 * - add first part of implementation of master style pages
 * - pdf exporter conversations for tables and images and added measurement helper
 *
 * Revision 1.5  2006/05/02 17:37:16  larsbm
 * - Flag added graphics with guid
 * - Set guid based read and write directories
 *
 * Revision 1.4  2006/02/16 18:35:41  larsbm
 * - Add FrameBuilder class
 * - TextSequence implementation (Todo loading!)
 * - Free draing postioning via x and y coordinates
 * - Graphic will give access to it's full qualified path
 *   via the GraphicRealPath property
 * - Fixed Bug with CellSpan in Spreadsheetdocuments
 * - Fixed bug graphic of loaded files won't be deleted if they
 *   are removed from the content.
 * - Break-Before property for Paragraph properties for Page Break
 *
 * Revision 1.3  2006/02/05 20:02:25  larsbm
 * - Fixed several bugs
 * - clean up some messy code
 *
 * Revision 1.2  2006/01/29 18:52:14  larsbm
 * - Added support for common styles (style templates in OpenOffice)
 * - Draw TextBox import and export
 * - DrawTextBox html export
 *
 * Revision 1.1  2006/01/29 11:29:46  larsbm
 * *** empty log message ***
 *
 * Revision 1.5  2005/12/18 18:29:46  larsbm
 * - AODC Gui redesign
 * - AODC HTML exporter refecatored
 * - Full Meta Data Support
 * - Increase textprocessing performance
 *
 * Revision 1.4  2005/12/12 19:39:17  larsbm
 * - Added Paragraph Header
 * - Added Table Row Header
 * - Fixed some bugs
 * - better whitespace handling
 * - Implmemenation of HTML Exporter
 *
 * Revision 1.3  2005/11/20 17:31:20  larsbm
 * - added suport for XLinks, TabStopStyles
 * - First experimental of loading dcuments
 * - load and save via importer and exporter interfaces
 *
 * Revision 1.2  2005/10/22 10:47:41  larsbm
 * - add graphic support
 *
 * Revision 1.1  2005/10/17 19:32:47  larsbm
 * - start vers. 1.0.3.0
 * - add frame, framestyle, graphic, graphicproperties
 *
 */