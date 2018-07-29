using PluginContracts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace excelscanner
{
    class PluginLoader
    {
        public ICollection<IExcelProcess> Load(string directory)
        {
            // Find plugins in source directory
            string[] dllFileNames = GetDLLs(directory);
            // Load the assemblies into memory
            ICollection<Assembly> assemblies = LoadAssemblies(dllFileNames);
            // Load the plugins from the assemblies
            ICollection<Type> pluginTypes = GetPluginTypes(assemblies);

            // Instantiate
            ICollection<IExcelProcess> plugins = new List<IExcelProcess>(pluginTypes.Count);
            foreach (Type type in pluginTypes)
            {
                IExcelProcess plugin = (IExcelProcess)Activator.CreateInstance(type);
                plugins.Add(plugin);
            }

            return plugins;
        }

        private string[] GetDLLs(string directory)
        {
            string[] dlls = null;
            if (Directory.Exists(directory))
            {
                dlls = Directory.GetFiles(directory, "*.dll");
            }

            return dlls;
        }

        private ICollection<Assembly> LoadAssemblies(string[] dllFileNames)
        {
            if (dllFileNames == null) return new List<Assembly>(0);

            ICollection<Assembly> assemblies = new List<Assembly>(dllFileNames.Length);
            foreach (string dllFile in dllFileNames)
            {
                AssemblyName an = AssemblyName.GetAssemblyName(dllFile);
                Assembly assembly = Assembly.Load(an);
                assemblies.Add(assembly);
            }

            return assemblies;
        }

        private ICollection<Type> GetPluginTypes(ICollection<Assembly> assemblies)
        {
            Type pluginType = typeof(IExcelProcess);
            ICollection<Type> pluginTypes = new List<Type>();

            foreach (Assembly assembly in assemblies)
            {
                if (assembly != null)
                {
                    Type[] types = assembly.GetTypes();
                    foreach (Type type in types)
                    {
                        if (type.IsInterface || type.IsAbstract)
                        {
                            continue;
                        }
                        else
                        {
                            if (type.GetInterface(pluginType.FullName) != null)
                            {
                                pluginTypes.Add(type);
                            }
                        }
                    }
                }
            }

            return pluginTypes;
        }
    }
}