namespace SKSB_App.Core.WorkbookOpen.Sync;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using SKSB_App.Core.WorkbookOpen.Sheet;

public static class EnumList
{
    public const string TargetSheet = SheetOperations.LookupsSheetName;

    /// <summary>
    /// Auto-discovers and syncs public Enums across Core and domain assemblies into the 'Lookups' tab.
    /// </summary>
    public static void SyncAllCoreEnums(params Type[] excludeTypes)
    {
        var excludeSet = new HashSet<Type>(excludeTypes ?? Array.Empty<Type>());
        var targetAssemblies = AppDomain.CurrentDomain.GetAssemblies()
            .Where(a => !a.IsDynamic && (a.FullName?.StartsWith("SKSB_App", StringComparison.OrdinalIgnoreCase) ?? false))
            .Distinct()
            .ToList();

        var coreAssembly = typeof(EnumList).Assembly;
        if (!targetAssemblies.Contains(coreAssembly))
        {
            targetAssemblies.Add(coreAssembly);
        }

        var allEnums = new List<Type>();
        foreach (var asm in targetAssemblies)
        {
            try
            {
                var enums = asm.GetTypes().Where(t => t.IsEnum && t.IsPublic && !excludeSet.Contains(t));
                allEnums.AddRange(enums);
            }
            catch (ReflectionTypeLoadException ex)
            {
                var loadable = ex.Types.Where(t => t != null && t.IsEnum && t.IsPublic && !excludeSet.Contains(t!));
                allEnums.AddRange(loadable!);
            }
            catch { }
        }

        var uniqueEnums = allEnums
            .GroupBy(t => t.Name)
            .Select(g => g.First())
            .OrderBy(t => t.Name)
            .ToArray();

        if (uniqueEnums.Length == 0) return;
        Sync(uniqueEnums);
    }

    public static void Sync(params Type[] enumTypes)
    {
        if (enumTypes == null || enumTypes.Length == 0) return;

        var mappings = new Dictionary<string, IEnumerable<string>?>();
        foreach (var type in enumTypes)
        {
            if (!type.IsEnum) continue;
            string[] values = Enum.GetNames(type)
                                  .Where(n => !string.Equals(n, "None", StringComparison.OrdinalIgnoreCase))
                                  .ToArray();
            mappings[type.Name] = values;
        }

        SyncHelper.SyncDictionary(TargetSheet, mappings, WriteMode.SequentialReset, injectPlaceholderIfEmpty: false);
    }

    public static void SyncDictionary(Dictionary<string, string[]> enumMappings)
    {
        var dict = enumMappings.ToDictionary(k => k.Key, v => (IEnumerable<string>?)v.Value);
        SyncHelper.SyncDictionary(TargetSheet, dict, WriteMode.SequentialReset, injectPlaceholderIfEmpty: false);
    }
}