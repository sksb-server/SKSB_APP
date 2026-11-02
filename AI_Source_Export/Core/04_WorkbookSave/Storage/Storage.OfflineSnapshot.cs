namespace SKSB_App.Core.WorkbookSave.Storage;

using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using SKSB_App.Core.Bootstrap.Storage;
using SKSB_App.Core.Integration.Database;

public class TolerantEnumConverter<T> : JsonConverter<T> where T : struct, Enum
{
    public override T Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType == JsonTokenType.Number)
        {
            if (reader.TryGetInt32(out int intVal))
            {
                return (T)Enum.ToObject(typeof(T), intVal);
            }
            if (reader.TryGetInt64(out long longVal))
            {
                return (T)Enum.ToObject(typeof(T), (int)longVal);
            }
        }

        if (reader.TokenType == JsonTokenType.String)
        {
            string? str = reader.GetString();
            if (string.IsNullOrWhiteSpace(str))
                return default;

            string normalized = str.Trim().Replace(" ", "_").Replace("-", "_");
            if (Enum.TryParse<T>(normalized, true, out var result))
                return result;

            // Safe fallback to default (None / 0) for unrecognized strings
            return default;
        }

        return default;
    }

    public override void Write(Utf8JsonWriter writer, T value, JsonSerializerOptions options)
    {
        writer.WriteStringValue(value.ToString());
    }
}

public class TolerantEnumConverterFactory : JsonConverterFactory
{
    public override bool CanConvert(Type typeToConvert)
    {
        Type underlying = Nullable.GetUnderlyingType(typeToConvert) ?? typeToConvert;
        return underlying.IsEnum;
    }

    public override JsonConverter? CreateConverter(Type typeToConvert, JsonSerializerOptions options)
    {
        Type underlying = Nullable.GetUnderlyingType(typeToConvert) ?? typeToConvert;
        Type converterType = typeof(TolerantEnumConverter<>).MakeGenericType(underlying);
        return (JsonConverter)Activator.CreateInstance(converterType)!;
    }
}

public static class OfflineSnapshot
{
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        Converters =
        {
            new TolerantEnumConverterFactory()
        }
    };

    public static string Save<T>(T data, bool writeToDatabase = true)
    {
        if (data == null)
            throw new ArgumentNullException(nameof(data), "Data payload cannot be null.");

        FolderRegistry.Data.EnsureExists();
        string typeName = ResolveTypeName(typeof(T));
        string targetFile = $"{typeName.ToLowerInvariant()}_data.json";
        string fullPath = FolderRegistry.Data.GetFilePath(targetFile);

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

        if (!FolderRegistry.Data.Exists())
            return default;

        string fullPath = FolderRegistry.Data.GetFilePath(targetFile);
        if (!File.Exists(fullPath))
            return default;

        try
        {
            string json = File.ReadAllText(fullPath);
            return JsonSerializer.Deserialize<T>(json, JsonOptions);
        }
        catch
        {
            return default;
        }
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

    public static string GetPendingUploadPath<T>() =>
        FolderRegistry.Data.GetFilePath($"{typeof(T).Name}_PendingUpload.json");

    public static string SaveOfflineSnapshot<T>(T data)
    {
        FolderRegistry.Data.EnsureExists();
        string filePath = GetPendingUploadPath<T>();
        string json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(filePath, json);
        return filePath;
    }

    public static bool HasPendingSnapshot<T>() =>
        File.Exists(GetPendingUploadPath<T>());

    public static T? LoadPendingSnapshot<T>()
    {
        string filePath = GetPendingUploadPath<T>();
        if (!File.Exists(filePath)) return default;

        try
        {
            string json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<T>(json, JsonOptions);
        }
        catch
        {
            return default;
        }
    }

    public static void ClearPendingSnapshot<T>()
    {
        string filePath = GetPendingUploadPath<T>();
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }

    public static DateTime? GetServerLatestTimestamp(string tableName)
    {
        try
        {
            var db = new PostgreSqlService();
            string sql = $"SELECT MAX(updated_at) FROM {tableName};";
            return db.ExecuteScalar<DateTime?>(sql);
        }
        catch
        {
            return null;
        }
    }
}