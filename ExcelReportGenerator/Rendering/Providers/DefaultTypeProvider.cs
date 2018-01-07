﻿using ExcelReportGenerator.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelReportGenerator.Rendering.Providers
{
    public class DefaultTypeProvider : ITypeProvider
    {
        private const char NamespaceSeparator = ':';

        private readonly IDictionary<string, Type> _typesCache = new Dictionary<string, Type>();

        /// <param name="assemblies">Collection of assemblies where types will be searched. If null or empty than current execution assembly will be used</param>
        /// <param name="defaultType">Type which will be returned if the template is not specified explicitly</param>
        public DefaultTypeProvider(ICollection<Assembly> assemblies = null, Type defaultType = null)
        {
            if (assemblies == null || !assemblies.Any())
            {
                Assembly entryAssembly = Assembly.GetEntryAssembly();
                if (entryAssembly == null)
                {
                    throw new InvalidOperationException("Assemblies are not provided but entry assembly is null. Provide assemblies and try again.");
                }
                Assemblies = new[] { entryAssembly };
            }
            else
            {
                Assemblies = assemblies;
            }

            DefaultType = defaultType;
        }

        protected ICollection<Assembly> Assemblies { get; }

        protected Type DefaultType { get; }

        /// <summary>
        /// Provides type based on template
        /// </summary>
        public virtual Type GetType(string typeTemplate)
        {
            if (string.IsNullOrWhiteSpace(typeTemplate))
            {
                return DefaultType ?? throw new InvalidOperationException("Template is not specified but defaultType is null");
            }

            if (_typesCache.TryGetValue(typeTemplate, out Type type))
            {
                return type;
            }

            string[] typeNameParts = typeTemplate.Split(NamespaceSeparator);
            bool isNamespaceSpecified = false;
            string @namespace = null;
            string name;
            if (typeNameParts.Length == 1)
            {
                name = typeNameParts[0];
            }
            else if (typeNameParts.Length == 2)
            {
                isNamespaceSpecified = true;
                @namespace = typeNameParts[0];
                @namespace = string.IsNullOrWhiteSpace(@namespace) ? null : @namespace.Trim();
                name = typeNameParts[1];
            }
            else
            {
                throw new IncorrectTemplateException($"Type name template \"{typeTemplate}\" is incorrect");
            }

            name = name.Trim();
            IEnumerable<Type> allAssembliesTypes = Assemblies.SelectMany(a => a.GetTypes());
            IList<Type> foundTypes = (isNamespaceSpecified
                    ? allAssembliesTypes.Where(t => t.Namespace == @namespace && t.Name == name)
                    : allAssembliesTypes.Where(t => t.Name == name))
                .ToList();

            if (foundTypes.Count == 1)
            {
                _typesCache[typeTemplate] = foundTypes.First();
                return _typesCache[typeTemplate];
            }
            if (!foundTypes.Any())
            {
                throw new TypeNotFoundException($"Cannot find type by template \"{typeTemplate}\"");
            }

            throw new IncorrectTemplateException($"More than one type found by template \"{typeTemplate}\"");
        }
    }
}