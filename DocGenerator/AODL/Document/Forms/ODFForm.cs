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
using AODL.Document;
using AODL.Document.Collections;
using AODL.Document.TextDocuments;
using AODL.Document.Forms.Controls;
using AODL.Document.Content;
using System.Xml;
using System.Collections;

namespace AODL.Document.Forms
{
	/// <summary>
	/// Summary description for ODFForm.
	/// </summary>

	public class ODFForm
	{
		
		private XmlNode _node;

		/// <summary>
		/// Represents the IRI of the processing agent for the form
		/// </summary>
		public string Href
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@xlink:href",
				                                         Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@xlink:href",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("href", "xlink"));
				nd.InnerText = value;
			}
		}

		/// <summary>
		/// Do not change it unless it is necessary
		/// </summary>
		public string ControlImplementation
		{
			get
			{
				XmlNode xn = Node.SelectSingleNode("@form:control-implementation",
				                                        Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = Node.SelectSingleNode("@form:control-implementation",
				                                        Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("control-implementation", "form"));
				nd.InnerText = value;
			}
		}

		/// <summary>
		/// Form name
		/// </summary>
		public string Name
		{
			get
			{
				XmlNode xn = Node.SelectSingleNode("@form:name",
				                                        Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = Node.SelectSingleNode("@form:name",
				                                        Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("name", "form"));
				nd.InnerText = value;
			}
		}

		/// <summary>
		/// Specifies the target frame of the form
		/// </summary>
		public TargetFrame TargetFrame
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@office:target-frame",
				                                         Document.NamespaceManager);
				if (xn == null) return TargetFrame.NotSet;
				
				string s = xn.InnerText;
				TargetFrame tf;
				switch (s)
				{
						case "_self": tf = TargetFrame.Self; break;
						case "_blank": tf = TargetFrame.Blank; break;
						case "_parent": tf = TargetFrame.Parent; break;
						case "_top": tf = TargetFrame.Top; break;
						default: tf = TargetFrame.Blank; break;
				}
				return tf;
			}
			set
			{
				string s;
				switch (value)
				{
						case TargetFrame.Self: s = "_self"; break;
						case TargetFrame.Blank: s = "_blank"; break;
						case TargetFrame.Parent: s = "_parent";  break;
						case TargetFrame.Top: s = "_top"; break;
						default: s = "_blank"; break;
				}
				XmlNode nd = _node.SelectSingleNode("@office:target-frame",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("target-frame", "office"));
				nd.InnerText = s;
			}
		}
		
		/// <summary>
		/// Specifies the HTTP method to use to submit the data in the form to
		/// the server
		/// </summary>
		public Method Method
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:method",
				                                         Document.NamespaceManager);
				if (xn == null) return Method.NotSet;

				string s = xn.InnerText;
				Method mt;
				switch (s)
				{
					case "get":
						mt = Method.Get;
						break;
					case "post":
						mt = Method.Post;
						break;
					default:
						mt = Method.NotSet;
						break;
				}
				return mt;
			}
			set
			{
				string s;
				switch (value)
				{
					case Method.Get:
						s = "get";
						break;
					case Method.Post:
						s = "post";
						break;
					default:
						return;
				}
				XmlNode nd = _node.SelectSingleNode("@form:method",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("method", "form"));
				nd.InnerText = s;
			}
		}

		/// <summary>
		///  Specifies the content type used to submit the form to the server
		/// </summary>
		public string Enctype
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:enctype",
				                                         Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@form:enctype",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("enctype", "form"));
				nd.InnerText = value;
			}
		}

		/// <summary>
		/// Specifies whether or not data records can be deleted
		/// </summary>
		public XmlBoolean AllowDeletes
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:allow-deletes",
				                                         Document.NamespaceManager);
				if (xn == null) return XmlBoolean.NotSet;

				string s = xn.InnerText;
				XmlBoolean at;
				switch (s)
				{
						case "true": at = XmlBoolean.True; break;
						case "false": at = XmlBoolean.False; break;
						default: at = XmlBoolean.NotSet; break;
				}
				return at;
			}
			set
			{
				string s;
				switch (value)
				{
						case XmlBoolean.True: s = "true"; break;
						case XmlBoolean.False: s = "false"; break;
						default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@form:allow-deletes",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("allow-deletes", "form"));
				nd.InnerText = s;
			}
		}

		/// <summary>
		/// Specifies whether or not new data records can be inserted
		/// </summary>
		public XmlBoolean AllowInserts
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:allow-inserts",
				                                         Document.NamespaceManager);
				if (xn == null) return XmlBoolean.NotSet;

				string s = xn.InnerText;
				XmlBoolean at;
				switch (s)
				{
						case "true": at = XmlBoolean.True; break;
						case "false": at = XmlBoolean.False; break;
						default: at = XmlBoolean.NotSet; break;
				}
				return at;
			}
			set
			{
				string s;
				switch (value)
				{
						case XmlBoolean.True: s = "true"; break;
						case XmlBoolean.False: s = "false"; break;
						default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@form:allow-inserts",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("allow-inserts", "form"));
				nd.InnerText = s;
			}
		}

		/// <summary>
		/// Specifies whether or not data records can be updated
		/// </summary>
		public XmlBoolean AllowUpdates
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:allow-updates",
				                                         Document.NamespaceManager);
				if (xn == null) return XmlBoolean.NotSet;

				string s = xn.InnerText;
				XmlBoolean at;
				switch (s)
				{
						case "true": at = XmlBoolean.True; break;
						case "false": at = XmlBoolean.False; break;
						default: at = XmlBoolean.NotSet; break;
				}
				return at;
			}
			set
			{
				string s;
				switch (value)
				{
						case XmlBoolean.True: s = "true"; break;
						case XmlBoolean.False: s = "false"; break;
						default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@form:allow-updates",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("allow-updates", "form"));
				nd.InnerText = s;
			}
		}

		/// <summary>
		/// Specifies whether or not filters should be applied to the form
		/// </summary>
		public XmlBoolean ApplyFilter
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:apply-filter",
				                                         Document.NamespaceManager);
				if (xn == null) return XmlBoolean.NotSet;

				string s = xn.InnerText;
				XmlBoolean at;
				switch (s)
				{
						case "true": at = XmlBoolean.True; break;
						case "false": at = XmlBoolean.False; break;
						default: at = XmlBoolean.NotSet; break;
				}
				return at;
			}
			set
			{
				string s;
				switch (value)
				{
						case XmlBoolean.True: s = "true"; break;
						case XmlBoolean.False: s = "false"; break;
						default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@form:apply-filter",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("apply-filter", "form"));
				nd.InnerText = s;
			}
		}
		

		/// <summary>
		/// Specifies the type of command to execute on the data source
		/// </summary>
		public CommandType CommandType
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:command-type",
				                                         Document.NamespaceManager);
				if (xn==null) return CommandType.NotSet;

				string s = xn.InnerText;
				CommandType ct;
				switch (s)
				{
						case "table": ct = CommandType.Table; break;
						case "query": ct = CommandType.Query; break;
						case "command": ct = CommandType.Command; break;
						default: ct = CommandType.NotSet; break;
				}
				return ct;
			}
			set
			{
				string s;
				switch (value)
				{
						case CommandType.Table: s = "table"; break;
						case CommandType.Query: s = "query"; break;
						case CommandType.Command: s = "command"; break;
						default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@form:command-type",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("command-type", "form"));
				nd.InnerText = s;
			}
		}
		
		/// <summary>
		/// Specifies the command to execute on the data source
		/// </summary>
		public string Command
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:command",
				                                         Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@form:command",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("command", "form"));
				nd.InnerText = value;
			}
		}
		
		/// <summary>
		/// Specifies the name of a data source to use for the form
		/// </summary>
		public string DataSource
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:datasource",
				                                         Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@form:datasource",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("datasource", "form"));
				nd.InnerText = value;
			}
		}
		
		/// <summary>
		/// Specifies the names of the columns in the result set represented by the parent form
		/// </summary>
		public string MasterFields
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:master-fields",
				                                         Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@form:master-fields",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("master-fields", "form"));
				nd.InnerText = value;
			}
		}

		/// <summary>
		/// Specifies the names of the columns in detail forms that are related to columns in the parent form
		/// </summary>
		public string DetailFields
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:detail-fields",
				                                         Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@form:detail-fields",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("detail-fields", "form"));
				nd.InnerText = value;
			}
		}

		/// <summary>
		/// Specifies whether or not the application processes the command before passing it to the
		/// database driver
		/// </summary>
		public XmlBoolean EscapeProcessing
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:escape-processing",
				                                         Document.NamespaceManager);
				if (xn == null) return XmlBoolean.NotSet;

				string s = xn.InnerText;
				XmlBoolean at;
				switch (s)
				{
						case "true": at = XmlBoolean.True; break;
						case "false": at = XmlBoolean.False; break;
						default: at = XmlBoolean.NotSet; break;
				}
				return at;
			}
			set
			{
				string s;
				switch (value)
				{
						case XmlBoolean.True: s = "true"; break;
						case XmlBoolean.False: s = "false"; break;
						default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@form:escape-processing",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("escape-processing", "form"));
				nd.InnerText = s;
			}
		}

		/// <summary>
		/// Specifies a filter for the command to base the form on
		/// </summary>
		public string Filter
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:filter",
				                                         Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@form:filter",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("filter", "form"));
				nd.InnerText = value;
			}
		}

		/// <summary>
		/// Specifies how the records in a database form are navigated
		/// </summary>
		public NavigationMode NavigationMode
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:navigation-mode",
				                                         Document.NamespaceManager);
				if (xn == null) return NavigationMode.NotSet;
				
				string s = xn.InnerText;
				NavigationMode nm;
				switch (s)
				{
						case "current": nm = NavigationMode.Current; break;
						case "parent": nm = NavigationMode.Parent; break;
						case "none": nm = NavigationMode.None; break;
						default: nm = NavigationMode.NotSet; break;
				}
				return nm;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@form:navigation-mode",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("navigation-mode", "form"));
				nd.InnerText = value.ToString().ToLower();
			}
		}

		
		/// <summary>
		/// Specifies whether or not to discard all results that are
		/// retrieved from the underlying data source
		/// </summary>
		public XmlBoolean IgnoreResult
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:ignore-result",
				                                         Document.NamespaceManager);
				if (xn == null) return XmlBoolean.NotSet;

				string s = xn.InnerText;
				XmlBoolean at;
				switch (s)
				{
						case "true": at = XmlBoolean.True; break;
						case "false": at = XmlBoolean.False; break;
						default: at = XmlBoolean.NotSet; break;
				}
				return at;
			}
			set
			{
				string s;
				switch (value)
				{
						case XmlBoolean.True: s = "true"; break;
						case XmlBoolean.False: s = "false"; break;
						default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@form:ignore-result",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("ignore-result", "form"));
				nd.InnerText = s;
			}
		}


		/// <summary>
		/// Specifies a sort criteria for the command.
		/// </summary>
		public string Order
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:order",
				                                         Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@form:order",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("order", "form"));
				nd.InnerText = value;
			}
		}

		/// <summary>
		/// Specifies how the application responds when the user presses
		/// the TAB key in the controls in a for
		/// </summary>
		public TabCycle TabCycle
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:tab-cycle",
				                                         Document.NamespaceManager);
				if (xn == null) return TabCycle.NotSet;
				
				string s = xn.InnerText;
				TabCycle tc;
				switch (s)
				{
						case "current": tc = TabCycle.Current; break;
						case "records": tc = TabCycle.Records; break;
						case "page": tc = TabCycle.Page; break;
						default: tc = TabCycle.NotSet; break;
				}
				return tc;
			}
			set
			{
				string s;
				switch (value)
				{
						case TabCycle.Current: s = "current"; break;
						case TabCycle.Records: s = "records"; break;
						case TabCycle.Page: s = "page"; break;
						default: return;
				}
				XmlNode nd = _node.SelectSingleNode("@form:tab-cycle",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("tab-cycle", "form"));
				nd.InnerText = s;
			}
		}
		
		/// <summary>
		/// Specifies the source database by an [XLink].
		/// </summary>
		public string ConnectionResource
		{
			get
			{
				XmlNode xn = _node.SelectSingleNode("@form:connection-resource",
				                                         Document.NamespaceManager);
				if (xn == null) return null;
				return xn.InnerText;
			}
			set
			{
				XmlNode nd = _node.SelectSingleNode("@form:connection-resource",
				                                         Document.NamespaceManager);
				if (nd == null)
					nd = Node.Attributes.Append(Document.CreateAttribute("connection-resource", "form"));
				nd.InnerText = value;
			}
		}


		/// <summary>
		/// The XML node that represents the form and its content
		/// </summary>
		public XmlNode Node
		{
			get { return _node; }
			set { _node = value; }
		}
		
		private void CreateBasicNode()
		{
			Node = Document.CreateNode("form", "form");
		}

		
		/// <summary>
		/// Creates an ODFForm
		/// </summary>
		/// <param name="document">Parent document</param>
		/// <param name="name">Form name</param>
		public ODFForm(IDocument document, string name)
		{
			Document = document;
			CreateBasicNode();
			ControlImplementation = "ooo:com.sun.star.form.component.Form";
			ApplyFilter = XmlBoolean.True;
			CommandType = CommandType.Table;
			Name = name;

			_controls = new ODFControlsCollection();
			_controls.Inserted += ControlsCollection_Inserted;
			_controls.Removed += ControlsCollection_Removed;
			_controls.Clearing += ControlsCollection_Clearing;

			_properties = new FormPropertyCollection();
			_properties.Inserted += PropertyCollection_Inserted;
			_properties.Removed += PropertyCollection_Removed;

			_formCollection = new ODFFormCollection();
			_formCollection.Inserted += FormCollection_Inserted;
			_formCollection.Removed += FormCollection_Removed;
			
		}

		public ODFForm(XmlNode node, IDocument document)
		{
			Document = document;
			Node = node;
			
			_controls = new ODFControlsCollection();
			_controls.Inserted +=ControlsCollection_Inserted;
			_controls.Removed +=ControlsCollection_Removed;
			_controls.Clearing += ControlsCollection_Clearing;

			_properties = new FormPropertyCollection();
			_properties.Inserted += PropertyCollection_Inserted;
			_properties.Removed += PropertyCollection_Removed;

			_formCollection = new ODFFormCollection();
			_formCollection.Inserted += FormCollection_Inserted;
			_formCollection.Removed += FormCollection_Removed;
		}

		public void SuppressControlEvents()
		{
			_controls.Inserted -=ControlsCollection_Inserted;
			_controls.Removed -=ControlsCollection_Removed;
			_controls.Clearing -= ControlsCollection_Clearing;
		}

		public void RestoreControlEvents()
		{
			_controls.Inserted +=ControlsCollection_Inserted;
			_controls.Removed +=ControlsCollection_Removed;
			_controls.Clearing += ControlsCollection_Clearing;
		}

		public void SuppressPropertyEvents()
		{
			_properties.Inserted -= PropertyCollection_Inserted;
			_properties.Removed -= PropertyCollection_Removed;
		}

		private void RestorePropertyEvents()
		{
			_properties.Inserted += PropertyCollection_Inserted;
			_properties.Removed += PropertyCollection_Removed;
		}

		private IDocument _document;

		/// <summary>
		/// Parent document
		/// </summary>
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

		private void ControlsCollection_Inserted(int index, object value)
		{
			ODFFormControl ctrl = value as ODFFormControl;
			Node.AppendChild(ctrl.Node);
			
			ctrl.AddToContentCollection();
		}

		private void ControlsCollection_Removed(int index, object value)
		{
			ODFFormControl ctrl = value as ODFFormControl;
			Node.RemoveChild(ctrl.Node);
			
			ctrl.RemoveFromContentCollection();
		}

		private void ControlsCollection_Clearing()
		{
			for (int i=0; i< _controls.Count; i++)
			{
                if (_controls[i] is ODFFormControl ctrl)
                {
                    Node.RemoveChild(ctrl.Node);
                    ctrl.RemoveFromContentCollection();
                }
            }
		}

		private ODFControlsCollection _controls;

		/// <summary>
		/// List of the child controls
		/// </summary>
		public ODFControlsCollection Controls
		{
			get
			{
				return _controls;
			}
			set
			{
				_controls = value;
			}
		}

		private void PropertyCollection_Inserted(int index, object value)
		{
			XmlNode form_prop;
			form_prop = Node.SelectSingleNode("form:properties", Document.NamespaceManager);
			
			if (form_prop == null)
			{
				form_prop = Document.CreateNode("properties", "form");
			}
			
			FormProperty prop = value as FormProperty;
			form_prop.AppendChild(prop.Node);
			Node.AppendChild(form_prop);
		}

		private void PropertyCollection_Removed(int index, object value)
		{
			XmlNode form_prop;
			form_prop = Node.SelectSingleNode("form:properties", Document.NamespaceManager);
			
			if (form_prop != null)
			{
				FormProperty prop = value as FormProperty;
				form_prop.RemoveChild(prop.Node);
				if (index == 0)
				{
					Node.RemoveChild(form_prop);
				}
			}
		}

		private FormPropertyCollection _properties;

		/// <summary>
		/// Generic form:property collection
		/// </summary>
		public FormPropertyCollection Properties
		{
			get
			{
				return _properties;
			}
			set
			{
				_properties = value;
			}
		}

		private ODFFormCollection _formCollection;

		/// <summary>
		/// Child forms collection
		/// </summary>
		public ODFFormCollection ChildForms
		{
			get { return _formCollection; }
			set { _formCollection = value; }
		}

		private void FormCollection_Inserted(int index, object value)
		{
			ODFForm child = (value as ODFForm);
			if (child != null)
				Node.AppendChild(child.Node);
		}

		private void FormCollection_Removed(int index, object value)
		{
			ODFForm child = (value as ODFForm);
			if (child !=null)
			{
				child.Controls.Clear();
				if (child.Node !=null)
					Node.RemoveChild(child.Node);
			}
		}

		/// <summary>
		/// Looks up a control by its ID
		/// </summary>
		/// <param name="id">Control ID</param>
		/// <param name="searchInSubforms">Specifies whether to look in the subforms</param>
		/// <returns></returns>
		public ODFFormControl FindControlById (string id, bool searchInSubforms)
		{
			if (searchInSubforms)
			{
				foreach (ODFForm f in ChildForms)
				{
					ODFFormControl ctrl;
					ctrl = f.FindControlById(id, searchInSubforms);
					if (ctrl != null) return ctrl;
				}
			}
			foreach (ODFFormControl c in _controls)
			{
				if ( c != null)
				{
					if (c.ID == id)
						return c;
				}
			}
			return null;
		}

		/// <summary>
		/// Looks up a control by its name
		/// </summary>
		/// <param name="id">Control name</param>
		/// <param name="searchInSubforms">Specifies whether to look in the subforms</param>
		/// <returns></returns>
		public ODFFormControl FindControlByName (string name, bool searchInSubforms)
		{
			if (searchInSubforms)
			{
				foreach (ODFForm f in ChildForms)
				{
					ODFFormControl ctrl;
					ctrl = f.FindControlByName(name, true);
					if (ctrl != null) return ctrl;
				}
			}
			foreach (ODFFormControl c in _controls)
			{
				if ( c != null)
				{
					if (c.Name == name)
						return c;
				}
			}
			return null;
		}

		/// <summary>
		/// Gets the generic form property by its name
		/// </summary>
		/// <param name="name">Generic form property</param>
		/// <returns></returns>
		public FormProperty GetFormProperty(string name)
		{
			foreach (FormProperty fp in _properties)
			{
				if (fp.Name == name)
				{
					return fp;
				}
			}
			return null;
		}

		public void FixPropertyCollection()
		{
			_properties.Clear();
			SuppressPropertyEvents();
			XmlNode form_prop = Node.SelectSingleNode("form:properties", Document.NamespaceManager);
			if (form_prop == null) return;

			foreach (XmlNode nodeChild in form_prop)
			{
				if (nodeChild.Name == "form:property" && nodeChild.ParentNode == form_prop)
				{
					SingleFormProperty sp = new SingleFormProperty(Document, nodeChild);
					_properties.Add(sp);
				}
				if (nodeChild.Name == "form:list-property" && nodeChild.ParentNode == form_prop)
				{
					ListFormProperty lp = new ListFormProperty(Document, nodeChild);
					_properties.Add(lp);
				}
			}
			RestorePropertyEvents();
		}
	}
}
