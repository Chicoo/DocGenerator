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
using System.Xml ;
using AODL.Document ;
using AODL.Document .Content ;
using AODL.Document.Content .Draw  ;
using AODL.Document .Styles; 

namespace AODL.Document.Content.EmbedObjects
{
	/// <summary>
	/// Summary description for EmbedObject.
	/// </summary>
	public class EmbedObject : IContent  
	{


		
		/// <summary>
		/// Gets or sets the H ref.
		/// </summary>
		/// <value>The H ref.</value>
		public string HRef
		{
			get 
			{ 
				XmlNode xn = _parentnode.SelectSingleNode("@xlink:href",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _parentnode.SelectSingleNode("@xlink:href",
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
				XmlNode xn = _parentnode.SelectSingleNode("@xlink:actuate",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _parentnode.SelectSingleNode("@xlink:actuate",
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
				XmlNode xn = _parentnode.SelectSingleNode("@xlink:type",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _parentnode.SelectSingleNode("@xlink:type",
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
				XmlNode xn = _parentnode.SelectSingleNode("@xlink:show",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _parentnode.SelectSingleNode("@xlink:show",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("show", value, "xlink");
				_node.SelectSingleNode("@xlink:show",
					Document.NamespaceManager).InnerText = value;
			}
		}

		private string _objectrealpath;
		/// <summary>
		/// Gets or sets the embed object real path.
		/// </summary>
		/// <value>The object real path.</value>
		public string ObjectRealPath
		{
			get { return _objectrealpath; }
			set { _objectrealpath = value; }
		}

		private string _objectfilename;
		/// <summary>
		/// Gets or sets the name of the object file.
		/// </summary>
		/// <value>The name of the object file.</value>
		public string ObjectFileName
		{
			get { return _objectfilename; }
			set { _objectfilename = value; }
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
		/// gets and sets the object name
		/// </summary>

		private string _objectname;

		public string ObjectName
		{
			get {return _objectname ;}

			set {_objectname =value;}
		}

		/// <summary>
		/// gets and sets the object type
		/// </summary>

		private string objecttype;

		public string ObjectType
		{
			get
			{
				return objecttype ;
			}

			set
			{
				objecttype =value;
			}
		}


		/// <summary>
		/// gets and sets the parent node 
		/// </summary>

		private XmlNode _parentnode;

		public XmlNode ParentNode
		{
			get
			{
				return _parentnode ;
			}

			set
			{
				_parentnode =value;
			}

		}

		#region IContent Member
		/// <summary>
		/// Gets or sets the name of the style.
		/// </summary>
		/// <value>The name of the style.</value>
		public virtual string StyleName
		{
			get 
			{ 
				return null;
			}
			set
			{
				
			}
		}

		protected IDocument _document;
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

		protected IStyle _style;
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

		protected XmlNode _node;
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
		
		public EmbedObject(XmlNode ParentNode,IDocument document)
		{
			this.ParentNode    = ParentNode;
			Document      = document;
			//this.NewXmlNode ();
			
			//this.Frame         =frame;  
		}

		public EmbedObject(IDocument document)
		{
			Document      = document;
			NewXmlNode();
		}

		private void NewXmlNode()
		{
			ParentNode 		= Document.CreateNode("object", "draw");       

			XmlAttribute xa = Document.CreateAttribute("href", "xlink");

			xa				= Document.CreateAttribute("type", "xlink");
			xa.Value		= "simple"; 

			ParentNode.Attributes.Append(xa);

			xa				= Document.CreateAttribute("show", "xlink");
			xa.Value		= "embed"; 

			ParentNode.Attributes.Append(xa);

			xa				= Document.CreateAttribute("actuate", "xlink");
			xa.Value		= "onLoad"; 

			ParentNode.Attributes.Append(xa);
		}

		protected  virtual void CreateAttribute(string name, string text, string prefix)
		{
			XmlAttribute xa = Document.CreateAttribute(name, prefix);
			xa.Value		= text;
			ParentNode .Attributes.Append(xa);
		}


	}

}
