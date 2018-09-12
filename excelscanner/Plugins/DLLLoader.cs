using NLog;
using EBPPluginContracts;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using System.Reflection;

namespace ExcelBatchProcessor.Plugins
{
    /// <summary>
    /// Loads assemblies containing plugins that can be used to process files.
    /// </summary>
    /// <remarks>
    /// Plugins are .dll assemblies that implement the <c>IExcelProcess</c> interface.
    /// All plugins need to be stored in the <c>./plugins</c> directory relative to
    /// the working directory of the application (Note: Customizable plugin locations
    /// are planned for a future update).
    /// </remarks>
    public class DLLLoader
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        readonly IFileSystem FileSystem;

        /// <summary>
        /// Constructor for the DLLoader class with injectable FileSystem.
        /// Use of this constructor should generally be limited to testing during development.
        /// Use the overloaded constructor instead.
        /// </summary>
        /// <param name="FileSystem"></param>
        public DLLLoader(IFileSystem FileSystem)
        {
            this.FileSystem = FileSystem;
        }
        /// <summary>
        /// Constructor for the DLLoader class. Consumers of the class will generally use
        /// this constructor.
        /// </summary>
        public DLLLoader() : this(FileSystem: new FileSystem())
        {

        }
        /// <summary>
        /// Loads assemblies from .dll files in the directory provided, instantiates
        /// the classes, and returns the collection of plugins ready to be used.
        /// </summary>
        /// <param name="directory"></param>
        /// <returns></returns>
        public virtual ICollection<IExcelProcess> Load(string directory)
        {
            // Find plugins in source directory
            logger.Trace("Searching for .dlls in {dir}", directory);
            string[] dllFileNames = GetDLLs(directory);
            logger.Trace("Found {num} .dlls", dllFileNames.Length);

            // Load the assemblies into memory
            ICollection<Assembly> assemblies = LoadAssemblies(dllFileNames);

            // Load the plugins from the assemblies
            ICollection<Type> pluginTypes = GetPluginTypes(assemblies);
            logger.Trace("Found {num} plugins that implement the plugin interface.", pluginTypes.Count);

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
            string[] dlls = new string[] {};
            if (FileSystem.Directory.Exists(directory))
            {
                dlls = FileSystem.Directory.GetFiles(directory, "*.dll");
            }

            return dlls;
        }

        private ICollection<Assembly> LoadAssemblies(string[] dllFileNames)
        {
            if (dllFileNames == null) return new List<Assembly>(0);

            ICollection<Assembly> assemblies = new List<Assembly>(dllFileNames.Length);
            foreach (string dllFile in dllFileNames)
            {
                try
                {
                    AssemblyName an = AssemblyName.GetAssemblyName(dllFile);
                    Assembly assembly = Assembly.Load(an);
                    assemblies.Add(assembly);
                }
                catch (FileLoadException ex)
                {
                    logger.Error(ex, "Could not load assembly {f}. Is another process using it?", dllFile);
                    continue;
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Could not load assembly {file}.", dllFile);
                    continue;
                }
                
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