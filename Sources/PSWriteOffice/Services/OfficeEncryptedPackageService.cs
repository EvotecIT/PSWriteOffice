using System;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace PSWriteOffice.Services;

internal static class OfficeEncryptedPackageService
{
    public static ExcelDocument LoadExcel(string path, string password, bool readOnly, bool autoSave)
    {
        return InvokeStatic<ExcelDocument>(
            typeof(ExcelDocument),
            nameof(LoadExcel),
            "LoadEncrypted",
            new[] { typeof(string), typeof(string), typeof(bool), typeof(bool) },
            path,
            password,
            readOnly,
            autoSave);
    }

    public static void SaveExcel(ExcelDocument document, string path, string password, bool openExcel, ExcelSaveOptions? saveOptions)
    {
        InvokeInstance(
            document,
            nameof(SaveExcel),
            "SaveEncrypted",
            new[] { typeof(string), typeof(string), typeof(bool), typeof(ExcelSaveOptions) },
            path,
            password,
            openExcel,
            saveOptions);
    }

    public static WordDocument LoadWord(string path, string password, bool readOnly, bool autoSave)
    {
        return InvokeStatic<WordDocument>(
            typeof(WordDocument),
            nameof(LoadWord),
            "LoadEncrypted",
            new[] { typeof(string), typeof(string), typeof(bool), typeof(bool) },
            path,
            password,
            readOnly,
            autoSave);
    }

    public static void SaveWord(WordDocument document, string path, string password, bool openWord)
    {
        InvokeInstance(
            document,
            nameof(SaveWord),
            "SaveEncrypted",
            new[] { typeof(string), typeof(string), typeof(bool) },
            path,
            password,
            openWord);
    }

    public static PowerPointPresentation OpenPowerPoint(string path, string password)
    {
        return InvokeStatic<PowerPointPresentation>(
            typeof(PowerPointPresentation),
            nameof(OpenPowerPoint),
            "OpenEncrypted",
            new[] { typeof(string), typeof(string) },
            path,
            password);
    }

    public static void SavePowerPoint(PowerPointPresentation presentation, Stream stream, string password)
    {
        InvokeInstance(
            presentation,
            nameof(SavePowerPoint),
            "SaveEncrypted",
            new[] { typeof(Stream), typeof(string) },
            stream,
            password);
    }

    private static T InvokeStatic<T>(Type declaringType, string featureName, string methodName, Type[] parameterTypes, params object?[] arguments)
    {
        var method = GetRequiredMethod(declaringType, featureName, methodName, parameterTypes);
        return (T)Invoke(method, null, arguments)!;
    }

    private static void InvokeInstance(object target, string featureName, string methodName, Type[] parameterTypes, params object?[] arguments)
    {
        var method = GetRequiredMethod(target.GetType(), featureName, methodName, parameterTypes);
        Invoke(method, target, arguments);
    }

    private static MethodInfo GetRequiredMethod(Type declaringType, string featureName, string methodName, Type[] parameterTypes)
    {
        var method = declaringType
            .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static)
            .FirstOrDefault(candidate => IsCompatibleMethod(candidate, methodName, parameterTypes));
        if (method != null)
        {
            return method;
        }

        throw new NotSupportedException($"OfficeIMO {declaringType.FullName}.{methodName} is not available in the loaded OfficeIMO package. Update OfficeIMO packages before using encrypted package support for {featureName}.");
    }

    private static object? Invoke(MethodInfo method, object? target, object?[] arguments)
    {
        var invokeArguments = ExpandArguments(method, arguments);
        try
        {
            return method.Invoke(target, invokeArguments);
        }
        catch (TargetInvocationException ex) when (ex.InnerException != null)
        {
            throw ex.InnerException;
        }
    }

    private static bool IsCompatibleMethod(MethodInfo method, string methodName, Type[] parameterTypes)
    {
        if (!string.Equals(method.Name, methodName, StringComparison.Ordinal))
        {
            return false;
        }

        var parameters = method.GetParameters();
        if (parameters.Length < parameterTypes.Length)
        {
            return false;
        }

        for (var i = 0; i < parameterTypes.Length; i++)
        {
            if (parameters[i].ParameterType != parameterTypes[i])
            {
                return false;
            }
        }

        return parameters.Skip(parameterTypes.Length).All(parameter => parameter.IsOptional || parameter.HasDefaultValue);
    }

    private static object?[] ExpandArguments(MethodInfo method, object?[] arguments)
    {
        var parameters = method.GetParameters();
        if (parameters.Length == arguments.Length)
        {
            return arguments;
        }

        var expanded = new object?[parameters.Length];
        Array.Copy(arguments, expanded, arguments.Length);
        for (var i = arguments.Length; i < parameters.Length; i++)
        {
            expanded[i] = parameters[i].HasDefaultValue ? parameters[i].DefaultValue : Type.Missing;
        }

        return expanded;
    }
}
