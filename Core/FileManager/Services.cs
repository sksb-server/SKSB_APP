namespace SKSB_App.Core.FileManager;

using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using SKSB_App.Core.FileManager.Constants;
using SKSB_App.Core.ServerManager;

public static class Services
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters =
        {
            new JsonStringEnumConverter()
        }
    };

    public static string Save<T>(T data, bool writeToDatabase = true)
    {
        if (data == null)
            throw new ArgumentNullException(nameof(data), "Data payload cannot be null.");

        Folder.Data.EnsureExists();

        string typeName = ResolveTypeName(typeof(T));
        string targetFile = $"{typeName.ToLowerInvariant()}_data.json";
        string fullPath = Folder.Data.GetFilePath(targetFile);

        string json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(fullPath, json);

        if (writeToDatabase)
        {
            ServerDispatcher.Dispatch(data);
        }

        return fullPath;
    }

    public static async Task<string> SaveAsync<T>(T data, bool writeToDatabase = true)
    {
        return await Task.Run(() => Save(data, writeToDatabase));
    }

    public static T? Load<T>()
    {
        string typeName = ResolveTypeName(typeof(T));
        string targetFile = $"{typeName.ToLowerInvariant()}_data.json";

        if (!Folder.Data.Exists())
            return default;

        string fullPath = Folder.Data.GetFilePath(targetFile);
        if (!File.Exists(fullPath))
            return default;

        string json = File.ReadAllText(fullPath);
        return JsonSerializer.Deserialize<T>(json, JsonOptions);
    }

    private static string ResolveTypeName(Type type)
    {
        if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(System.Collections.Generic.List<>))
        {
            return type.GetGenericArguments()[0].Name;
        }

        if (type.IsArray)
        {
            return type.GetElementType()?.Name ?? "Object";
        }

        return type.Name;
    }
}