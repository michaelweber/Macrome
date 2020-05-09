using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Binding;
using System.CommandLine.Builder;
using System.CommandLine.Invocation;
using System.CommandLine.Parsing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace Macrome
{
    /// <summary>
    /// Almost entirely taken from CommandLine.DragonFruit (https://github.com/dotnet/command-line-api/tree/master/src/System.CommandLine.DragonFruit)
    /// so we can use the CLI generation functionality without affecting the EntryPoint function
    /// </summary>
    public static class CommandLineExtensions
    {
        public class CommandHelpMetadata
        {
            public string Description { get; set; }

            public string Name { get; set; }

            public Dictionary<string, string> ParameterDescriptions { get; } = new Dictionary<string, string>();
        }

        public class XmlDocReader
        {
            private IEnumerable<XElement> _members { get; }

            private XmlDocReader(XDocument document)
            {
                _members = document.Descendants("members");
            }

            public static bool TryLoad(string filePath, out XmlDocReader xmlDocReader)
            {
                try
                {
                    return TryLoad(File.OpenText(filePath), out xmlDocReader);
                }
                catch
                {
                    xmlDocReader = null;
                    return false;
                }
            }

            public static bool TryLoad(TextReader reader, out XmlDocReader xmlDocReader)
            {
                try
                {
                    xmlDocReader = new XmlDocReader(XDocument.Load(reader));
                    return true;
                }
                catch
                {
                    xmlDocReader = null;
                    return false;
                }
            }

            public bool TryGetMethodDescription(MethodInfo info, out CommandHelpMetadata commandHelpMetadata)
            {
                commandHelpMetadata = null;

                var sb = new StringBuilder();
                sb.Append("M:");
                AppendTypeName(sb, info.DeclaringType);
                sb.Append(".")
                  .Append(info.Name)
                  .Append("(");

                bool first = true;
                foreach (ParameterInfo param in info.GetParameters())
                {
                    if (first)
                    {
                        first = false;
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    AppendTypeName(sb, param.ParameterType);
                }

                sb.Append(")");

                string name = sb.ToString();

                XElement member = _members.Elements("member")
                                         .FirstOrDefault(m => string.Equals(m.Attribute("name")?.Value, name));

                if (member == null)
                {
                    return false;
                }

                commandHelpMetadata = new CommandHelpMetadata();

                foreach (XElement element in member.Elements())
                {
                    switch (element.Name.ToString())
                    {
                        case "summary":
                            commandHelpMetadata.Description = element.Value?.Trim();
                            break;
                        case "param":
                            commandHelpMetadata.ParameterDescriptions.Add(element.Attribute("name")?.Value, element.Value?.Trim());
                            break;
                    }
                }

                return true;
            }

            private static void AppendTypeName(StringBuilder sb, Type type)
            {
                if (type.IsNested)
                {
                    AppendTypeName(sb, type.DeclaringType);
                    sb.Append(".").Append(type.Name);
                }
                else if (type.IsGenericType)
                {
                    var typeDefName = type.GetGenericTypeDefinition().FullName;

                    sb.Append(typeDefName.Substring(0, typeDefName.IndexOf("`")));

                    sb.Append("{");

                    foreach (var genericArgument in type.GetGenericArguments())
                    {
                        AppendTypeName(sb, genericArgument);
                    }

                    sb.Append("}");
                }
                else
                {
                    sb.Append(type.FullName);
                }
            }
        }

        private static readonly string[] _argumentParameterNames =
        {
            "arguments",
            "argument",
            "args"
        };

        public static void ConfigureFromMethod(
            this Command command,
            MethodInfo method,
            object target = null)
        {
            if (command == null)
            {
                throw new ArgumentNullException(nameof(command));
            }

            if (method == null)
            {
                throw new ArgumentNullException(nameof(method));
            }

            foreach (var option in method.BuildOptions())
            {
                command.AddOption(option);
            }

            if (method.GetParameters()
                .FirstOrDefault(p => _argumentParameterNames.Contains(p.Name)) is ParameterInfo argsParam)
            {
                var argument = new Argument
                {
                    ArgumentType = argsParam.ParameterType,
                    Name = argsParam.Name
                };

                if (argsParam.HasDefaultValue)
                {
                    if (argsParam.DefaultValue != null)
                    {
                        argument.SetDefaultValue(argsParam.DefaultValue);
                    }
                    else
                    {
                        argument.SetDefaultValueFactory(() => null);
                    }
                }

                command.AddArgument(argument);
            }

            if (XmlDocReader.TryLoad(GetDefaultXmlDocsFileLocation(method.DeclaringType.Assembly), out var xmlDocs))
            {
                if (xmlDocs.TryGetMethodDescription(method, out CommandHelpMetadata metadata) &&
                    metadata.Description != null)
                {
                    command.Description = metadata.Description;
                    var options = command.Options.ToArray();

                    foreach (var parameterDescription in metadata.ParameterDescriptions)
                    {
                        var kebabCasedParameterName = parameterDescription.Key.ToKebabCase();

                        var option = options.FirstOrDefault(o => o.HasAlias(kebabCasedParameterName));

                        if (option != null)
                        {
                            option.Description = parameterDescription.Value;
                        }
                        else
                        {
                            foreach (var argument in command.Arguments)
                            {
                                if (string.Equals(
                                    argument.Name,
                                    kebabCasedParameterName,
                                    StringComparison.OrdinalIgnoreCase))
                                {
                                    argument.Description = parameterDescription.Value;
                                }
                            }
                        }
                    }

                    metadata.Name = method.DeclaringType.Assembly.GetName().Name;
                }
            }

            command.Handler = CommandHandler.Create(method, target);
        }

        public static IEnumerable<Option> BuildOptions(this MethodInfo method)
        {
            var descriptor = HandlerDescriptor.FromMethodInfo(method);

            var omittedTypes = new[]
            {
                typeof(IConsole),
                typeof(InvocationContext),
                typeof(BindingContext),
                typeof(ParseResult),
                typeof(CancellationToken),
            };

            foreach (var option in descriptor.ParameterDescriptors
                .Where(d => !omittedTypes.Contains(d.ValueType))
                .Where(d => !_argumentParameterNames.Contains(d.ValueName))
                .Select(p => p.BuildOption()))
            {
                yield return option;
            }
        }

        public static Option BuildOption(this ParameterDescriptor parameter)
        {
            var argument = new Argument
            {
                ArgumentType = parameter.ValueType
            };

            if (parameter.HasDefaultValue)
            {
                argument.SetDefaultValueFactory(parameter.GetDefaultValue);
            }

            var option = new Option(
                parameter.BuildAlias(),
                parameter.ValueName)
            {
                Argument = argument
            };

            return option;
        }

        public static string BuildAlias(this IValueDescriptor descriptor)
        {
            if (descriptor == null)
            {
                throw new ArgumentNullException(nameof(descriptor));
            }

            return BuildAlias(descriptor.ValueName);
        }

        internal static string BuildAlias(string parameterName)
        {
            if (String.IsNullOrWhiteSpace(parameterName))
            {
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(parameterName));
            }

            return parameterName.Length > 1
                ? $"--{parameterName.ToKebabCase()}"
                : $"-{parameterName.ToLowerInvariant()}";
        }

        public static CommandLineBuilder ConfigureHelpFromXmlComments(
            this CommandLineBuilder builder,
            MethodInfo method,
            string xmlDocsFilePath)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            if (method == null)
            {
                throw new ArgumentNullException(nameof(method));
            }

            if (XmlDocReader.TryLoad(xmlDocsFilePath ?? GetDefaultXmlDocsFileLocation(method.DeclaringType.Assembly), out var xmlDocs))
            {
                if (xmlDocs.TryGetMethodDescription(method, out CommandHelpMetadata metadata) &&
                    metadata.Description != null)
                {
                    builder.Command.Description = metadata.Description;
                    var options = builder.Options.ToArray();

                    foreach (var parameterDescription in metadata.ParameterDescriptions)
                    {
                        var kebabCasedParameterName = parameterDescription.Key.ToKebabCase();

                        var option = options.FirstOrDefault(o => o.HasAlias(kebabCasedParameterName));

                        if (option != null)
                        {
                            option.Description = parameterDescription.Value;
                        }
                        else
                        {
                            foreach (var argument in builder.Command.Arguments)
                            {
                                if (string.Equals(
                                        argument.Name,
                                        kebabCasedParameterName,
                                        StringComparison.OrdinalIgnoreCase))
                                {
                                    argument.Description = parameterDescription.Value;
                                }
                            }
                        }
                    }

                    metadata.Name = method.DeclaringType.Assembly.GetName().Name;
                }
            }

            return builder;
        }

        private static string GetDefaultXmlDocsFileLocation(Assembly assembly)
        {
            return Path.Combine(
                Path.GetDirectoryName(assembly.Location),
                Path.GetFileNameWithoutExtension(assembly.Location) + ".xml");
        }
    }
}
