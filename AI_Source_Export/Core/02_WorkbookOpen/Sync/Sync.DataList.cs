namespace SKSB_App.Core.WorkbookOpen.Sync;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Storage;

public static class DataList
{
    public const string TargetSheet = SheetOperations.DataSheetName;

    private class PropertyDescriptor
    {
        public string ListName { get; init; } = string.Empty;
        public Func<object, object?> Getter { get; init; } = null!;
    }

    /// <summary>
    /// Synchronizes specific entity models requested by an operational manager.
    /// Preserves column row-alignment across compound fields (e.g., StaffID and StaffName).
    /// </summary>
    public static int SyncEntities(params Type[] entityTypes)
    {
        if (entityTypes == null || entityTypes.Length == 0) return 0;
        return ExecuteEntityBlockSync(headersOnly: false, entityTypes);
    }

    /// <summary>
    /// Discovers and synchronizes all public entities decorated with [DataList] across all assemblies.
    /// Dedicated strictly for Template Edit Mode.
    /// </summary>
    public static int SyncAllTemplateData(bool headersOnly = false)
    {
        var targetAssemblies = AppDomain.CurrentDomain.GetAssemblies()
            .Where(a => !a.IsDynamic && (a.FullName?.StartsWith("SKSB_App", StringComparison.OrdinalIgnoreCase) ?? false))
            .Distinct()
            .ToList();

        var allClassTypes = new List<Type>();
        foreach (var asm in targetAssemblies)
        {
            try
            {
                var types = asm.GetTypes().Where(t => t.IsClass && t.IsPublic && !t.IsAbstract && !t.IsEnum);
                allClassTypes.AddRange(types);
            }
            catch (ReflectionTypeLoadException ex)
            {
                var loadable = ex.Types.Where(t => t != null && t.IsClass && t.IsPublic && !t.IsAbstract && !t.IsEnum);
                allClassTypes.AddRange(loadable!);
            }
            catch { }
        }

        return ExecuteEntityBlockSync(headersOnly, allClassTypes.ToArray());
    }

    private static int ExecuteEntityBlockSync(bool headersOnly, Type[] targetTypes)
    {
        var childDetailTypes = new HashSet<Type>();
        foreach (var type in targetTypes)
        {
            foreach (var prop in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (prop.GetCustomAttribute<DataDetailAttribute>() != null ||
                    (prop.PropertyType.IsClass && prop.PropertyType != typeof(string) && !prop.PropertyType.IsArray && !typeof(IEnumerable).IsAssignableFrom(prop.PropertyType)))
                {
                    childDetailTypes.Add(prop.PropertyType);
                }
            }
        }

        var rootEntityTypes = targetTypes.Where(t => !childDetailTypes.Contains(t)).Distinct().ToList();

        var loadMethod = typeof(OfflineSnapshot).GetMethod(
            nameof(OfflineSnapshot.Load),
            BindingFlags.Public | BindingFlags.Static,
            Type.EmptyTypes
        );

        var entityBlocks = new List<EntityTableBlock>();
        var registeredListNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var type in rootEntityTypes)
        {
            var descriptors = DiscoverDecoratedProperties(type);
            descriptors = descriptors.Where(d => !registeredListNames.Contains(d.ListName)).ToList();
            if (descriptors.Count == 0) continue;

            IList? cachedRecords = null;
            if (!headersOnly && loadMethod != null)
            {
                try
                {
                    var genericLoad = loadMethod.MakeGenericMethod(typeof(List<>).MakeGenericType(type));
                    cachedRecords = genericLoad.Invoke(null, null) as IList;
                }
                catch { }
            }

            var block = BuildEntityBlock(type.Name, descriptors, cachedRecords, injectPlaceholderIfEmpty: SheetOperations.IsTemplateEditMode);
            if (block != null && block.ColumnNames.Count > 0)
            {
                entityBlocks.Add(block);
                foreach (var col in block.ColumnNames)
                {
                    registeredListNames.Add(col);
                }
            }
        }

        SyncHelper.SyncEntityBlocks(TargetSheet, entityBlocks);
        return entityBlocks.Sum(b => b.ColumnNames.Count);
    }

    public static bool SyncData<TEntity>(IEnumerable<TEntity>? entities = null) where TEntity : class
    {
        var records = entities?.ToList() ?? OfflineSnapshot.Load<List<TEntity>>();
        if (records == null) return false;

        var descriptors = DiscoverDecoratedProperties(typeof(TEntity));
        if (descriptors.Count == 0) return false;

        var block = BuildEntityBlock(typeof(TEntity).Name, descriptors, records, injectPlaceholderIfEmpty: SheetOperations.IsTemplateEditMode);
        if (block == null) return false;

        SyncHelper.SyncEntityBlocks(TargetSheet, new List<EntityTableBlock> { block });
        return true;
    }

    public static bool SyncProperty<TEntity, TProp>(
        Expression<Func<TEntity, TProp>> propertyExpression,
        IEnumerable<TEntity>? entities = null,
        string? customListName = null) where TEntity : class
    {
        var records = entities?.ToList() ?? OfflineSnapshot.Load<List<TEntity>>();
        if (records == null || records.Count == 0) return false;

        string propName = (propertyExpression.Body as MemberExpression)?.Member.Name ?? "Item";
        string listName = customListName ?? $"{propName}List";

        var compiled = propertyExpression.Compile();
        var values = records
            .Select(r => compiled(r)?.ToString()?.Trim())
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToList();

        SyncValues(listName, values!);
        return true;
    }

    public static void SyncValues(string listName, IEnumerable<string> values)
    {
        if (string.IsNullOrWhiteSpace(listName)) return;

        var dict = new Dictionary<string, IEnumerable<string>?>
        {
            { listName, values }
        };

        SyncHelper.SyncDictionary(
            TargetSheet,
            dict,
            WriteMode.SmartOrganize,
            injectPlaceholderIfEmpty: SheetOperations.IsTemplateEditMode
        );
    }

    private static EntityTableBlock? BuildEntityBlock(
        string entityName,
        List<PropertyDescriptor> descriptors,
        IList? cachedRecords,
        bool injectPlaceholderIfEmpty)
    {
        if (descriptors == null || descriptors.Count == 0) return null;

        var colNames = descriptors.Select(d => d.ListName.EndsWith("List", StringComparison.OrdinalIgnoreCase)
            ? d.ListName
            : $"{d.ListName}List").ToList();

        var validRecords = new List<object>();
        if (cachedRecords != null)
        {
            foreach (var rec in cachedRecords)
            {
                if (rec == null) continue;
                object? primaryVal = descriptors[0].Getter(rec);
                string primaryStr = primaryVal?.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(primaryStr) && primaryStr != "0")
                {
                    validRecords.Add(rec);
                }
            }
        }

        // Numeric primary key sort (e.g., StaffID 10, 15, 100, 101)
        if (validRecords.Count > 1)
        {
            validRecords = validRecords.OrderBy(r =>
            {
                object? val = descriptors[0].Getter(r);
                if (val == null) return (object)0L;
                if (val is long l) return l;
                if (val is int i) return (long)i;
                if (long.TryParse(val.ToString(), out long parsedL)) return parsedL;
                return val.ToString()!;
            }).ToList();
        }

        int rowCount = validRecords.Count;
        int colCount = descriptors.Count;

        if (rowCount == 0)
        {
            if (!injectPlaceholderIfEmpty) return null;

            object[,] placeholderGrid = new object[1, colCount];
            for (int c = 0; c < colCount; c++)
            {
                placeholderGrid[0, c] = $"[{colNames[c].Replace("List", "")} Placeholder]";
            }

            return new EntityTableBlock
            {
                EntityName = entityName,
                ColumnNames = colNames,
                Grid = placeholderGrid
            };
        }

        object[,] dataGrid = new object[rowCount, colCount];
        for (int r = 0; r < rowCount; r++)
        {
            var record = validRecords[r];
            for (int c = 0; c < colCount; c++)
            {
                object? rawVal = descriptors[c].Getter(record);
                if (rawVal == null)
                {
                    dataGrid[r, c] = string.Empty;
                }
                else if (rawVal is long l)
                {
                    dataGrid[r, c] = (double)l;
                }
                else if (rawVal is int i)
                {
                    dataGrid[r, c] = (double)i;
                }
                else if (rawVal is double d)
                {
                    dataGrid[r, c] = d;
                }
                else
                {
                    string str = rawVal.ToString()?.Trim() ?? string.Empty;
                    if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedD) && str.Length < 10 && !str.StartsWith("0"))
                    {
                        dataGrid[r, c] = parsedD;
                    }
                    else
                    {
                        dataGrid[r, c] = str;
                    }
                }
            }
        }

        return new EntityTableBlock
        {
            EntityName = entityName,
            ColumnNames = colNames,
            Grid = dataGrid
        };
    }

    private static List<PropertyDescriptor> DiscoverDecoratedProperties(Type rootType)
    {
        var results = new List<PropertyDescriptor>();

        void Scan(Type targetType, Func<object, object?> parentGetter)
        {
            var props = targetType.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var p in props)
            {
                var attr = p.GetCustomAttribute<DataListAttribute>();
                if (attr != null)
                {
                    string listName = !string.IsNullOrWhiteSpace(attr.CustomListName)
                        ? attr.CustomListName!
                        : (p.Name.EndsWith("List", StringComparison.OrdinalIgnoreCase) ? p.Name : $"{p.Name}List");

                    results.Add(new PropertyDescriptor
                    {
                        ListName = listName,
                        Getter = obj =>
                        {
                            var parent = parentGetter(obj);
                            return parent != null ? p.GetValue(parent) : null;
                        }
                    });
                }
                else if (p.GetCustomAttribute<DataDetailAttribute>() != null ||
                         (p.PropertyType.IsClass && p.PropertyType != typeof(string) && !p.PropertyType.IsArray && !typeof(IEnumerable).IsAssignableFrom(p.PropertyType)))
                {
                    Scan(p.PropertyType, obj =>
                    {
                        var parent = parentGetter(obj);
                        return parent != null ? p.GetValue(parent) : null;
                    });
                }
            }
        }

        Scan(rootType, obj => obj);
        return results;
    }
}