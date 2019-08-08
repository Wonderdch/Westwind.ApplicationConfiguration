#region License
/*
 **************************************************************
 *  Author: Rick Strahl 
 *          ?West Wind Technologies, 2009-2013
 *          http://www.west-wind.com/
 * 
 * Created: 09/12/2009
 *
 * Permission is hereby granted, free of charge, to any person
 * obtaining a copy of this software and associated documentation
 * files (the "Software"), to deal in the Software without
 * restriction, including without limitation the rights to use,
 * copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the
 * Software is furnished to do so, subject to the following
 * conditions:
 * 
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
 * OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
 * HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
 * WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
 * FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
 * OTHER DEALINGS IN THE SOFTWARE.
 **************************************************************  
*/
#endregion

using System;
using System.Collections;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Reflection;
using System.Xml;
using System.Globalization;

namespace Westwind.Utilities.Configuration
{

	/// <summary>
	/// Reads and Writes configuration settings in .NET config files and 
	/// sections. Allows reading and writing to default or external files 
	/// and specification of the configuration section that settings are
	/// applied to.
	/// </summary>
	public class ConfigurationFileConfigurationProvider<TAppConfiguration> :
		ConfigurationProviderBase<TAppConfiguration>
		where TAppConfiguration : AppConfiguration, new()
	{

		/// <summary>
		/// Optional - the Configuration file where configuration settings are
		/// stored in. If not specified uses the default Configuration Manager
		/// and its default store.
		/// </summary>
		public string ConfigurationFile { get; set; }

		/// <summary>
		/// Optional The Configuration section where settings are stored.
		/// If not specified the appSettings section is used.
		/// </summary>
		//public new string ConfigurationSection {get; set; }


		/// <summary>
		/// internal property used to ensure there are no multiple write
		/// operations at the same time
		/// </summary>
		private object syncWriteLock = new object();

		/// <summary>
		/// Internally used reference to the Namespace Manager object
		/// used to make sure we're searching the proper Namespace
		/// for the appSettings section when reading and writing manually
		/// </summary>
		private XmlNamespaceManager XmlNamespaces = null;

		//Internally used namespace prefix for the default namespace
		private string XmlNamespacePrefix = "ww:";


		/// <summary>
		/// Reads configuration settings into a new instance of the configuration object.
		/// </summary>
		/// <typeparam name="TAppConfig"></typeparam>
		/// <returns></returns>
		public override TAppConfig Read<TAppConfig>()
		{
			var config = Activator.CreateInstance(typeof(TAppConfig), true) as TAppConfig;

			if (!Read(config)) return null;

			return config;
		}

		/// <summary>
		/// Reads configuration settings from the current configuration manager. 
		/// Uses the internal APIs to write these values.
		/// </summary>
		/// <typeparam name="TAppConfiguration"></typeparam>
		/// <param name="config"></param>
		/// <returns></returns>
		public override bool Read(AppConfiguration config)
		{
			// Config reading from external files works a bit differently 
			// so use a separate method to handle it
			if (!string.IsNullOrEmpty(ConfigurationFile)) return Read(config, ConfigurationFile);

			Type typeWebConfig = config.GetType();
			MemberInfo[] fields =
				typeWebConfig.GetMembers(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty |
										 BindingFlags.GetField);

			// Set a flag for missing fields
			// If we have any we'll need to write them out into .config
			bool missingFields = false;

			// Refresh the sections - re-read after write operations
			// sometimes sections don't want to re-read            
			ConfigurationManager.RefreshSection(string.IsNullOrEmpty(ConfigurationSection)
				? "appSettings"
				: ConfigurationSection);

			NameValueCollection configManager = string.IsNullOrEmpty(ConfigurationSection)
				? ConfigurationManager.AppSettings
				: ConfigurationManager.GetSection(ConfigurationSection) as NameValueCollection;

			if (configManager == null)
			{
				Write(config);
				return true;
			}

			// Loop through all fields and properties                 
			foreach (MemberInfo member in fields)
			{
				Type fieldType = null;

				if (member.MemberType == MemberTypes.Field)
				{
					var field = (FieldInfo)member;
					fieldType = field.FieldType;
				}
				else if (member.MemberType == MemberTypes.Property)
				{
					var property = (PropertyInfo)member;
					fieldType = property.PropertyType;
				}
				else
				{
					continue;
				}

				string fieldName = member.Name.ToLower();

				// Error Message is an internal public property
				if (fieldName == "errormessage" || fieldName == "provider") continue;

				if (!IsIList(fieldType))
				{
					// Single value
					string value = configManager[fieldName];

					if (value == null)
					{
						missingFields = true;
						continue;
					}

					try
					{
						// Assign the value to the property
						ReflectionUtils.SetPropertyEx(config, member.Name,
							StringToTypedValue(value, fieldType, CultureInfo.InvariantCulture));
					}
					catch
					{
					}
				}
				else
				{
					// List Value
					var list = Activator.CreateInstance(fieldType) as IList;

					var elementType = fieldType.GetElementType();
					if (elementType == null)
					{
						var generic = fieldType.GetGenericArguments();
						if (generic != null && generic.Length > 0) elementType = generic[0];
					}

					int count = 1;
					string value = string.Empty;

					while (value != null)
					{
						value = configManager[fieldName + count];
						if (value == null) break;

						list.Add(StringToTypedValue(value, elementType, CultureInfo.InvariantCulture));
						count++;
					}

					try
					{
						ReflectionUtils.SetPropertyEx(config, member.Name, list);
					}
					catch { }
				}
			}

			DecryptFields(config);

			// We have to write any missing keys
			if (missingFields) Write(config);

			return true;
		}

		private static bool IsIList(Type type)
		{
			// Enumerable types explicitly supported as 'simple values'
			if (type == typeof(string) || type == typeof(byte[])) return false;

			return type.GetInterface("IList") != null;
		}

		/// <summary>
		/// Reads Configuration settings from an external file or explicitly from a file.
		/// Uses XML DOM to read values instead of using the native APIs.
		/// </summary>
		/// <typeparam name="TAppConfiguration"></typeparam>
		/// <param name="config">Configuration instance</param>
		/// <param name="filename">Filename to read from</param>
		/// <returns></returns>
		public override bool Read(AppConfiguration config, string filename)
		{
			Type typeWebConfig = config.GetType();
			MemberInfo[] fields = typeWebConfig.GetMembers(BindingFlags.Public |
														   BindingFlags.Instance);

			// Set a flag for missing fields
			// If we have any we'll need to write them out 
			bool missingFields = false;

			XmlDocument dom = new XmlDocument();

			try
			{
				dom.Load(filename);
			}
			catch
			{
				// Can't open or doesn't exist - so try to create it
				if (!Write(config)) return false;

				// Now load again
				dom.Load(filename);
			}

			// Retrieve XML Namespace information to assign default 
			// Namespace explicitly.
			GetXmlNamespaceInfo(dom);

			string configSection = ConfigurationSection;
			if (configSection == string.Empty) configSection = "appSettings";

			foreach (MemberInfo member in fields)
			{
				Type fieldType = null;

				if (member.MemberType == MemberTypes.Field)
				{
					var field = (FieldInfo)member;
					fieldType = field.FieldType;
				}
				else if (member.MemberType == MemberTypes.Property)
				{
					var property = (PropertyInfo)member;
					fieldType = property.PropertyType;
				}
				else
				{
					continue;
				}

				string fieldname = member.Name;
				if (fieldname == "Provider" || fieldname == "ErrorMessage") continue;

				XmlNode sectionNode = dom.DocumentElement.SelectSingleNode(XmlNamespacePrefix + configSection, XmlNamespaces);
				if (sectionNode == null)
				{
					sectionNode = CreateConfigSection(dom, ConfigurationSection);
					dom.DocumentElement.AppendChild(sectionNode);
				}

				string value = GetNamedValueFromXml(dom, fieldname, configSection);
				if (value == null)
				{
					missingFields = true;
					continue;
				}

				// Assign the Property
				ReflectionUtils.SetPropertyEx(config, fieldname,
					StringToTypedValue(value, fieldType, CultureInfo.InvariantCulture));
			}

			DecryptFields(config);

			// We have to write any missing keys
			if (missingFields) Write(config);

			return true;
		}

		public override bool Write(AppConfiguration config)
		{
			EncryptFields(config);

			lock (syncWriteLock)
			{
				// Load the config file into DOM parser
				XmlDocument dom = new XmlDocument();

				string configFile = ConfigurationFile;

				if (string.IsNullOrEmpty(configFile)) configFile = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;

				try
				{
					dom.Load(configFile);
				}
				catch
				{
					// Can't load the file - create an empty document
					dom.LoadXml(@"<?xml version='1.0'?>
		<configuration>
		</configuration>");
				}

				// Load up the Namespaces object so we can 
				// reference the appropriate default namespace
				GetXmlNamespaceInfo(dom);

				// Parse through each of hte properties of the properties
				Type typeWebConfig = config.GetType();
				MemberInfo[] fields = typeWebConfig.GetMembers(BindingFlags.Instance | BindingFlags.GetField | BindingFlags.GetProperty | BindingFlags.Public);


				string configSection = "appSettings";

				if (!string.IsNullOrEmpty(ConfigurationSection)) configSection = ConfigurationSection;

				// make sure we're getting the latest values before we write
				ConfigurationManager.RefreshSection(configSection);

				foreach (MemberInfo field in fields)
				{
					// Don't persist ErrorMessage property
					if (field.Name == "ErrorMessage" || field.Name == "Provider") continue;

					object rawValue;
					switch (field.MemberType)
					{
						case MemberTypes.Field:
							rawValue = ((FieldInfo)field).GetValue(config);
							break;
						case MemberTypes.Property:
							rawValue = ((PropertyInfo)field).GetValue(config, null);
							break;
						default:
							continue;
					}

					string value = TypedValueToString(rawValue, CultureInfo.InvariantCulture);

					if (value == "ILIST_TYPE")
					{
						var count = 0;
						foreach (var item in rawValue as IList)
						{
							value = TypedValueToString(item, CultureInfo.InvariantCulture);
							WriteConfigurationValue(field.Name + ++count, value, field, dom, configSection);
						}
					}
					else
					{
						WriteConfigurationValue(field.Name, value, field, dom, configSection);
					}
				}

				try
				{
					// this will fail if permissions are not there
					dom.Save(configFile);

					ConfigurationManager.RefreshSection(configSection);
				}
				catch
				{
					return false;
				}
				finally
				{
					DecryptFields(config);
				}
			}

			return true;
		}

		private void WriteConfigurationValue(string keyName, string value, MemberInfo field, XmlDocument dom, string configSection)
		{
			XmlNode node = dom.DocumentElement.SelectSingleNode(
				XmlNamespacePrefix + configSection + "/" +
				XmlNamespacePrefix + "add[@key='" + keyName + "']", XmlNamespaces);

			if (node != null)
			{
				// just write the value into the attribute
				node.Attributes.GetNamedItem("value").Value = value;
				return;
			}

			// Create the node and attributes and write it
			node = dom.CreateNode(XmlNodeType.Element, "add", dom.DocumentElement.NamespaceURI);

			XmlAttribute keyAttribute = dom.CreateAttribute("key");
			keyAttribute.Value = keyName;
			XmlAttribute valueAttribute = dom.CreateAttribute("value");
			valueAttribute.Value = value;

			node.Attributes.Append(keyAttribute);
			node.Attributes.Append(valueAttribute);

			XmlNode parent = dom.DocumentElement.SelectSingleNode(
				XmlNamespacePrefix + configSection, XmlNamespaces);

			if (parent == null) parent = CreateConfigSection(dom, configSection);

			parent.AppendChild(node);
		}

		/// <summary>
		/// Returns a single value from the XML in a configuration file.
		/// </summary>
		/// <param name="dom"></param>
		/// <param name="key"></param>
		/// <param name="configSection"></param>
		/// <returns></returns>
		protected string GetNamedValueFromXml(XmlDocument dom, string key, string configSection)
		{
			XmlNode node = dom.DocumentElement.SelectSingleNode(
				   XmlNamespacePrefix + configSection + @"/" +
				   XmlNamespacePrefix + "add[@key='" + key + "']", XmlNamespaces);

			if (node == null) return null;

			return node.Attributes["value"].Value;
		}

		/// <summary>
		/// Used to load up the default namespace reference and prefix
		/// information. This is required so that SelectSingleNode can
		/// find info in 2.0 or later config files that include a namespace
		/// on the root element definition.
		/// </summary>
		/// <param name="dom"></param>
		protected void GetXmlNamespaceInfo(XmlDocument dom)
		{
			// Load up the Namespaces object so we can 
			// reference the appropriate default namespace
			if (dom.DocumentElement.NamespaceURI == null || dom.DocumentElement.NamespaceURI == string.Empty)
			{
				XmlNamespaces = null;
				XmlNamespacePrefix = string.Empty;
			}
			else
			{
				XmlNamespacePrefix = string.IsNullOrEmpty(dom.DocumentElement.Prefix) ? "ww" : dom.DocumentElement.Prefix;

				XmlNamespaces = new XmlNamespaceManager(dom.NameTable);
				XmlNamespaces.AddNamespace(XmlNamespacePrefix, dom.DocumentElement.NamespaceURI);

				XmlNamespacePrefix += ":";
			}
		}

		/// <summary>
		/// Creates a Configuration section and also creates a ConfigSections section for new 
		/// non appSettings sections.
		/// </summary>
		/// <param name="dom"></param>
		/// <param name="configSection"></param>
		/// <returns></returns>
		private XmlNode CreateConfigSection(XmlDocument dom, string configSection)
		{
			// Create the actual section first and attach to document
			XmlNode appSettingsNode = dom.CreateNode(XmlNodeType.Element,
				configSection, dom.DocumentElement.NamespaceURI);

			XmlNode parent = dom.DocumentElement.AppendChild(appSettingsNode);

			// Now check and make sure that the section header exists
			if (configSection != "appSettings")
			{
				XmlNode configSectionHeader = dom.DocumentElement.SelectSingleNode(XmlNamespacePrefix + "configSections",
								XmlNamespaces);
				if (configSectionHeader == null)
				{
					// Create the node and attributes and write it
					XmlNode configSectionNode = dom.CreateNode(XmlNodeType.Element,
							 "configSections", dom.DocumentElement.NamespaceURI);

					// Insert as first element in DOM
					configSectionHeader = dom.DocumentElement.InsertBefore(configSectionNode,
							 dom.DocumentElement.ChildNodes[0]);
				}

				// Check for the Section
				XmlNode section = configSectionHeader.SelectSingleNode(XmlNamespacePrefix + "section[@name='" + configSection + "']",
						XmlNamespaces);

				if (section == null)
				{
					section = dom.CreateNode(XmlNodeType.Element, "section", null);

					XmlAttribute nameAttribute = dom.CreateAttribute("name");
					nameAttribute.Value = configSection;

					XmlAttribute typeAttribute = dom.CreateAttribute("type");
					typeAttribute.Value = "System.Configuration.NameValueSectionHandler,System,Version=1.0.3300.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";

					XmlAttribute requireAttribute = dom.CreateAttribute("requirePermission");
					requireAttribute.Value = "false";

					section.Attributes.Append(nameAttribute);
					section.Attributes.Append(requireAttribute);
					section.Attributes.Append(typeAttribute);

					configSectionHeader.AppendChild(section);
				}
			}

			return parent;
		}


		/// <summary>
		/// Converts a type to string if possible. This method supports an optional culture generically on any value.
		/// It calls the ToString() method on common types and uses a type converter on all other objects
		/// if available
		/// </summary>
		/// <param name="rawValue">The Value or Object to convert to a string</param>
		/// <param name="culture">Culture for numeric and DateTime values</param>
		/// <param name="unsupportedReturn">Return string for unsupported types</param>
		/// <returns>string</returns>
		private static string TypedValueToString(object rawValue, CultureInfo culture = null, string unsupportedReturn = null)
		{
			if (rawValue == null) return string.Empty;

			if (culture == null) culture = CultureInfo.CurrentCulture;

			Type valueType = rawValue.GetType();

			string returnValue;

			if (valueType == typeof(string))
			{
				returnValue = rawValue as string;
			}
			else if (valueType == typeof(int) || valueType == typeof(decimal) ||
					 valueType == typeof(double) || valueType == typeof(float) || valueType == typeof(Single))
			{
				returnValue = string.Format(culture.NumberFormat, "{0}", rawValue);
			}
			else if (valueType == typeof(DateTime))
			{
				returnValue = string.Format(culture.DateTimeFormat, "{0}", rawValue);
			}
			else if (valueType == typeof(bool) || valueType == typeof(Byte) || valueType.IsEnum)
			{
				returnValue = rawValue.ToString();
			}
			else if (valueType == typeof(byte[]))
			{
				returnValue = Convert.ToBase64String(rawValue as byte[]);
			}
			else if (valueType == typeof(Guid?))
			{
				if (rawValue == null)
				{
					returnValue = string.Empty;
				}
				else
				{
					return rawValue.ToString();
				}
			}
			else if (rawValue is IList)
			{
				return "ILIST_TYPE";
			}
			else
			{
				// Any type that supports a type converter
				TypeConverter converter = TypeDescriptor.GetConverter(valueType);
				if (converter != null && converter.CanConvertTo(typeof(string)) && converter.CanConvertFrom(typeof(string)))
				{
					returnValue = converter.ConvertToString(null, culture, rawValue);
				}
				else
				{
					// Last resort - just call ToString() on unknown type
					returnValue = !string.IsNullOrEmpty(unsupportedReturn) ? unsupportedReturn : rawValue.ToString();
				}
			}

			return returnValue;
		}

		/// <summary>
		/// Turns a string into a typed value generically.
		/// Explicitly assigns common types and falls back
		/// on using type converters for unhandled types.         
		/// 
		/// Common uses: 
		/// * UI -&gt; to data conversions
		/// * Parsers
		/// <seealso>Class ReflectionUtils</seealso>
		/// </summary>
		/// <param name="sourceString"> The string to convert from </param>
		/// <param name="targetType"> The type to convert to </param>
		/// <param name="culture"> Culture used for numeric and datetime values. </param>
		/// <returns>object. Throws exception if it cannot be converted.</returns>
		private static object StringToTypedValue(string sourceString, Type targetType, CultureInfo culture = null)
		{
			object result = null;

			bool isEmpty = string.IsNullOrEmpty(sourceString);

			if (culture == null) culture = CultureInfo.CurrentCulture;

			if (targetType == typeof(string))
			{
				result = sourceString;
			}
			else if (targetType == typeof(Int32) || targetType == typeof(int))
			{
				result = isEmpty ? 0 : Int32.Parse(sourceString, NumberStyles.Any, culture.NumberFormat);
			}
			else if (targetType == typeof(Int64))
			{
				result = isEmpty ? (Int64)0 : Int64.Parse(sourceString, NumberStyles.Any, culture.NumberFormat);
			}
			else if (targetType == typeof(Int16))
			{
				result = isEmpty ? (Int16)0 : Int16.Parse(sourceString, NumberStyles.Any, culture.NumberFormat);
			}
			else if (targetType == typeof(decimal))
			{
				result = isEmpty ? 0M : decimal.Parse(sourceString, NumberStyles.Any, culture.NumberFormat);
			}
			else if (targetType == typeof(DateTime))
			{
				result = isEmpty ? DateTime.MinValue : Convert.ToDateTime(sourceString, culture.DateTimeFormat);
			}
			else if (targetType == typeof(byte))
			{
				result = isEmpty ? 0 : Convert.ToByte(sourceString);
			}
			else if (targetType == typeof(double))
			{
				result = isEmpty ? 0F : Double.Parse(sourceString, NumberStyles.Any, culture.NumberFormat);
			}
			else if (targetType == typeof(Single))
			{
				result = isEmpty ? 0F : Single.Parse(sourceString, NumberStyles.Any, culture.NumberFormat);
			}
			else if (targetType == typeof(bool))
			{
				if (!isEmpty &&
					sourceString.ToLower() == "true" || sourceString.ToLower() == "on" || sourceString == "1")
				{
					result = true;
				}
				else
				{
					result = false;
				}
			}
			else if (targetType == typeof(Guid))
			{
				result = isEmpty ? Guid.Empty : new Guid(sourceString);
			}
			else if (targetType.IsEnum)
			{
				result = Enum.Parse(targetType, sourceString);
			}
			else if (targetType == typeof(byte[]))
			{
				result = Convert.FromBase64String(sourceString);
			}
			else if (targetType.Name.StartsWith("Nullable`"))
			{
				if (sourceString.ToLower() == "null" || sourceString == string.Empty)
				{
					result = null;
				}
				else
				{
					targetType = Nullable.GetUnderlyingType(targetType);
					result = StringToTypedValue(sourceString, targetType);
				}
			}
			else
			{
				// Check for TypeConverters or FromString static method
				TypeConverter converter = TypeDescriptor.GetConverter(targetType);

				if (converter != null && converter.CanConvertFrom(typeof(string)))
				{
					result = converter.ConvertFromString(null, culture, sourceString);
				}
				else
				{
					// Try to invoke a static FromString method if it exists
					try
					{
						var mi = targetType.GetMethod("FromString");
						if (mi != null)
						{
							return mi.Invoke(null, new object[] { sourceString });
						}
					}
					catch
					{
						// ignore error and assume not supported 
					}

					Debug.Assert(false, $"Type Conversion not handled in StringToTypedValue for {targetType.Name} {sourceString}");

					//throw (new InvalidCastException(Resources.StringToTypedValueValueTypeConversionFailed + targetType.Name));
				}
			}

			return result;
		}
	}
}
