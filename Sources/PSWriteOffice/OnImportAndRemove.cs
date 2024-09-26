using System;
using System.IO;
using System.Management.Automation;
using System.Reflection;
using System.Collections.Generic;

public class OnModuleImportAndRemove : IModuleAssemblyInitializer, IModuleAssemblyCleanup {
    public void OnImport() {
        if (IsNetFramework()) {
            AppDomain.CurrentDomain.AssemblyResolve += MyResolveEventHandler;
        }
    }

    public void OnRemove(PSModuleInfo module) {
        if (IsNetFramework()) {
            AppDomain.CurrentDomain.AssemblyResolve -= MyResolveEventHandler;
        }
    }

    private static Assembly MyResolveEventHandler(object sender, ResolveEventArgs args) {
        var libDirectory = Path.GetDirectoryName(typeof(OnModuleImportAndRemove).Assembly.Location);
        var directoriesToSearch = new List<string> { libDirectory };

        if (Directory.Exists(libDirectory)) {
            directoriesToSearch.AddRange(Directory.GetDirectories(libDirectory, "*", SearchOption.AllDirectories));
        }

        var requestedAssemblyName = new AssemblyName(args.Name).Name + ".dll";

        foreach (var directory in directoriesToSearch) {
            var assemblyPath = Path.Combine(directory, requestedAssemblyName);

            if (File.Exists(assemblyPath)) {
                try {
                    return Assembly.LoadFrom(assemblyPath);
                } catch (Exception ex) {
                    Console.WriteLine($"Failed to load assembly from {assemblyPath}: {ex.Message}");
                }
            }
        }

        return null;
    }

    private bool IsNetFramework() {
        return System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription.StartsWith(".NET Framework", StringComparison.OrdinalIgnoreCase);
    }
    private bool IsNetCore() {
        return System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription.StartsWith(".NET Core", StringComparison.OrdinalIgnoreCase);
    }
    private bool IsNet5OrHigher() {
        return System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription.StartsWith(".NET 5", StringComparison.OrdinalIgnoreCase) ||
               System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription.StartsWith(".NET 6", StringComparison.OrdinalIgnoreCase) ||
               System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription.StartsWith(".NET 7", StringComparison.OrdinalIgnoreCase) ||
               System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription.StartsWith(".NET 8", StringComparison.OrdinalIgnoreCase) ||
               System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription.StartsWith(".NET 9", StringComparison.OrdinalIgnoreCase);
    }
}