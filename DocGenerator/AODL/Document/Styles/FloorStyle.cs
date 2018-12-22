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
using AODL.Document .Styles .Properties ;
using AODL.Document ;
using AODL.Document .Styles ;
using AODL.Document .Collections ;

namespace AODL.Document.Styles
{
	/// <summary>
	/// Summary description for FloorStyle.
	/// </summary>
	public class FloorStyle : AbstractStyle
	{

		/// <summary>
		/// gets and sets the chart graphic properties
		/// </summary>
		public ChartGraphicProperties ChartGraphicProperties
		{
			get
			{
				foreach(IProperty property in PropertyCollection)
					if (property is ChartGraphicProperties)
						return (ChartGraphicProperties)property;
				ChartGraphicProperties chartGraphicProperties	= new ChartGraphicProperties(this);
				PropertyCollection.Add((IProperty)chartGraphicProperties);
				return chartGraphicProperties;
			}
			set
			{
				if (PropertyCollection.Contains((IProperty)value))
					PropertyCollection.Remove((IProperty)value);
				PropertyCollection.Add(value);
			}
		}

		/// <summary>
		/// gets and sets family style
		/// </summary>

		public string FamilyStyle
		{
			get 
			{ 
				XmlNode xn = _node.SelectSingleNode("@style:family",
					Document.NamespaceManager);
				if (xn != null)
					return xn.InnerText;
				return null;
			}
			set
			{
				XmlNode xn = _node.SelectSingleNode("@style:family",
					Document.NamespaceManager);
				if (xn == null)
					CreateAttribute("family", value, "style");
				_node.SelectSingleNode("@style:family",
					Document.NamespaceManager).InnerText = value;
			}
		}

		/// <summary>
		/// the constructor of floor style
		/// </summary>
		/// <param name="document"></param>
		/// <param name="node"></param>

		public FloorStyle(IDocument document)
		{
			Document			= document;
			InitStandards();
			//this.Node               = node;
			NewXmlNode ();
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="TableStyle"/> class.
		/// </summary>
		/// <param name="document">The spreadsheet document.</param>
		/// <param name="styleName">Name of the style.</param>
		public FloorStyle(IDocument document, string styleName)
		{
			Document					= document;
			InitStandards();
			StyleName					= styleName;
		}

		/// <summary>
		/// Inits the standards.
		/// </summary>
		private void InitStandards()
		{
			NewXmlNode();
			PropertyCollection				= new IPropertyCollection();
			PropertyCollection.Inserted	+= PropertyCollection_Inserted;
			PropertyCollection.Removed		+= PropertyCollection_Removed;
			FamilyStyle					= "chart";
			//			this.Document.Styles.Add(this);
		}

		/// <summary>
		/// Create a new Xml node.
		/// </summary>
		private void NewXmlNode()
		{
			Node		= Document.CreateNode("style", "style");
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
		/// Properties the collection_ inserted.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void PropertyCollection_Inserted(int index, object value)
		{
			Node.AppendChild(((IProperty)value).Node);
		}

		/// <summary>
		/// Properties the collection_ removed.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		private void PropertyCollection_Removed(int index, object value)
		{
			Node.RemoveChild(((IProperty)value).Node);
		}
	}
}

