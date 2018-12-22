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
using System.Drawing;
using System.Xml;
using AODL.Document.Styles;
using AODL.Document.Content;
using AODL.Document;

namespace AODL.Document.Content.Draw
{
	/// <summary>
	/// Graphic represent a graphic resp. image.
	/// </summary>
	public class Graphic : IContent, IContentContainer
	{
		/// <summary>
		/// Gets or sets the H ref.
		/// </summary>
		/// <value>The H ref.</value>
		public string HRef
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:href",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@xlink:href",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("href", value, "xlink");
				_node.SelectSingleNode("@xlink:href",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the actuate.
		/// e.g. onLoad
		/// </summary>
		/// <value>The actuate.</value>
		public string Actuate
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:actuate",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@xlink:actuate",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("actuate", value, "xlink");
				_node.SelectSingleNode("@xlink:actuate",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the type of the Xlink.
		/// e.g. simple, standard, ..
		/// </summary>
		/// <value>The type of the X link.</value>
		public string XLinkType
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:type",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@xlink:type",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("type", value, "xlink");
				_node.SelectSingleNode("@xlink:type",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// Gets or sets the show.
		/// e.g. embed
		/// </summary>
		/// <value>The show.</value>
		public string Show
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@xlink:show",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@xlink:show",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("show", value, "xlink");
				_node.SelectSingleNode("@xlink:show",
					Document.NamespaceManager).InnerText = value;
			}
		}

		private string _graphicRealPath;
		/// <summary>
		/// Gets or sets the graphic real path.
		/// </summary>
		/// <value>The graphic real path.</value>
		public string GraphicRealPath
		{
			get { return _graphicRealPath; }
			set { _graphicRealPath = value; }
		}

		private string _graphicFileName;
		/// <summary>
		/// Gets or sets the name of the graphic file.
		/// </summary>
		/// <value>The name of the graphic file.</value>
		public string GraphicFileName
		{
			get { return _graphicFileName; }
			set { _graphicFileName = value; }
		}

		private Frame _frame;
		/// <summary>
		/// Gets or sets the frame.
		/// </summary>
		/// <value>The frame.</value>
		public Frame Frame
		{
			get { return _frame; }
			set { _frame = value; }
		}
		
		/// <summary>
		/// Initializes a new instance of the <see cref="Graphic"/> class.
		/// </summary>
		/// <param name="document">The document.</param>
		/// <param name="frame">The frame.</param>
		/// <param name="graphiclink">The graphiclink.</param>
		public Graphic(IDocument document, Frame frame, string graphiclink)
		{
			Frame			= frame;
			Document		= document;
			GraphicFileName = graphiclink;
			NewXmlNode("Pictures/"+graphiclink);
			InitStandards();
			Document.Graphics.Add(this);
			Document.DocumentMetadata.ImageCount	+= 1;
		}

		/// <summary>
		/// Inits the standards.
		/// </summary>
		private void InitStandards()
		{
			Content				= new ContentCollection();
			Content.Inserted		+= Content_Inserted;
			Content.Removed		+= Content_Removed;
		}

		/// <summary>
		/// Create a new XmlNode.
		/// </summary>
		/// <param name="graphiclink">The stylename which should be referenced with this frame.</param>
		private void NewXmlNode(string graphiclink)
		{			
			Node		= Document.CreateNode("image", "draw");

			XmlAttribute xa = Document.CreateAttribute("href", "xlink");
			xa.Value		= graphiclink;

			Node.Attributes.Append(xa);

			xa				= Document.CreateAttribute("type", "xlink");
			xa.Value		= "standard"; 

			Node.Attributes.Append(xa);

			xa				= Document.CreateAttribute("show", "xlink");
			xa.Value		= "embed"; 

			Node.Attributes.Append(xa);

			xa				= Document.CreateAttribute("actuate", "xlink");
			xa.Value		= "onLoad"; 

			Node.Attributes.Append(xa);
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
		/// Content_s the inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Inserted(int index, object value)
		{
			Node.AppendChild(((IContent)value).Node);
		}

		/// <summary>
		/// Content_s the removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void Content_Removed(int index, object value)
		{
			Node.RemoveChild(((IContent)value).Node);
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
				XmlNode xn = _node.SelectSingleNode("@text:style-name",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@text:style-name",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("style-name", value, "text");
				_node.SelectSingleNode("@text:style-name",
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

//		#region IHtml Member
//
//		/// <summary>
//		/// Return the content as Html string
//		/// </summary>
//		/// <returns>The html string</returns>
//		public string GetHtml()
//		{
//			string align	= "<span align=\"#align#\">\n";
//			string html		= "<img src=\""+this.GetHtmlImgFolder()+"\" hspace=\"14\" vspace=\"14\""; //>\n";
//
//			Size size		= this.GetSizeInPix();
//			html			+= " width=\""+size.Width+"\" height=\""+size.Height+"\">\n";
//
//			if (this.Frame.FrameStyle.GraphicProperties != null)
//				if (this.Frame.FrameStyle.GraphicProperties.HorizontalPosition != null)
//				{
//					align	= align.Replace("#align#", 
//						this.Frame.FrameStyle.GraphicProperties.HorizontalPosition);
//					html	= align + html + "</span>\n";
//				}
//
//			return html;
//		}
//
//		/// <summary>
//		/// Gets the HTML img folder.
//		/// </summary>
//		/// <returns></returns>
//		private string GetHtmlImgFolder()
//		{
//			try
//			{
//				if (this.Frame.RealGraphicName != null)
//				{
//					return "temphtmlimg/Pictures/"+this.Frame.RealGraphicName;
//				}
//			}
//			catch(Exception ex)
//			{
//			}
//			return "";
//		}
//
//		/// <summary>
//		/// Gets the size in pix. As it is set in the frame.
//		/// This is needed, because the size of the graphic
//		/// could be another.
//		/// </summary>
//		/// <returns>The size in pixel</returns>
//		private Size GetSizeInPix()
//		{
//			try
//			{
//				double pxtocm		= 37.7928;
//				double intocm		= 2.41;
//				double height		= 0.0;
//				double width		= 0.0;
//
//				if (this.Frame.GraphicHeight.IndexOf("cm") > 0)
//				{
//					height		= Convert.ToDouble(this.Frame.GraphicHeight.Replace("cm",""), 
//						System.Globalization.NumberFormatInfo.InvariantInfo);
//					width		= Convert.ToDouble(this.Frame.GraphicWidth.Replace("cm",""),
//						System.Globalization.NumberFormatInfo.InvariantInfo);
//
//					height				*= pxtocm;
//					width				*= pxtocm;
//				}
//				else if (this.Frame.GraphicHeight.IndexOf("in") > 0)
//				{
//					height		= Convert.ToDouble(this.Frame.GraphicHeight.Replace("in",""), 
//						System.Globalization.NumberFormatInfo.InvariantInfo);
//					width		= Convert.ToDouble(this.Frame.GraphicWidth.Replace("in",""),
//						System.Globalization.NumberFormatInfo.InvariantInfo);
//
//					height				*= intocm*pxtocm;
//					width				*= intocm*pxtocm;
//				}
//				else if (this.Frame.GraphicHeight.IndexOf("px") > 0)
//				{
//					height		= Convert.ToDouble(this.Frame.GraphicHeight.Replace("px",""), 
//						System.Globalization.NumberFormatInfo.InvariantInfo);
//					width		= Convert.ToDouble(this.Frame.GraphicWidth.Replace("px",""),
//						System.Globalization.NumberFormatInfo.InvariantInfo);
//				}
//
//
//				Size size			= new Size((int)width, (int)height);
//
//				return size;
//			}
//			catch(Exception ex)
//			{
//				throw;
//			}
//		}
//
//		#endregion

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
	}
}
