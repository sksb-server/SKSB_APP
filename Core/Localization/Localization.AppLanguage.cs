namespace SKSB_App.Core.Localization;

using System;
using System.Collections.Generic;

/// <summary>UI languages supported by SKSB_App (Portal, Ribbon and Core dialogs).</summary>
public enum AppLanguage
{
    En,
    Ms,
    Zh,
    Id,
    MsEn
}

/// <summary>Codes, display names and the canonical display order of <see cref="AppLanguage"/>.</summary>
public static class AppLanguageInfo
{
    /// <summary>Display order used by the Portal ComboBox and the Ribbon dropdown.</summary>
    public static IReadOnlyList<AppLanguage> All { get; } = new[]
    {
        AppLanguage.En,
        AppLanguage.Ms,
        AppLanguage.Zh,
        AppLanguage.Id,
        AppLanguage.MsEn
    };

    public static string ToCode(AppLanguage language) => language switch
    {
        AppLanguage.Ms => "ms",
        AppLanguage.Zh => "zh",
        AppLanguage.Id => "id",
        AppLanguage.MsEn => "ms-en",
        _ => "en"
    };

    public static AppLanguage FromCode(string? code)
    {
        string c = (code ?? string.Empty).Trim().Replace('_', '-').ToLowerInvariant();
        return c switch
        {
            "ms" or "bm" or "ms-my" => AppLanguage.Ms,
            "zh" or "zh-cn" or "zh-hans" or "zh-my" or "cn" => AppLanguage.Zh,
            "id" or "id-id" => AppLanguage.Id,
            "ms-en" or "msen" or "manglish" => AppLanguage.MsEn,
            _ => AppLanguage.En
        };
    }

    public static string GetDisplayName(AppLanguage language) => language switch
    {
        AppLanguage.Ms => "Bahasa Melayu",
        AppLanguage.Zh => "中文",
        AppLanguage.Id => "Bahasa Indonesia",
        AppLanguage.MsEn => "Manglish (BM-EN)",
        _ => "English"
    };

    /// <summary>Position of <paramref name="language"/> inside <see cref="All"/> (0 when unknown).</summary>
    public static int IndexOf(AppLanguage language)
    {
        for (int i = 0; i < All.Count; i++)
        {
            if (All[i] == language) return i;
        }
        return 0;
    }
}
