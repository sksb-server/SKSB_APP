using SKSB_App.Core.ExcelManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKSB_App.Core.ExcelManager;

public static class Context
{
    private static string _excelPath;
    private static Bitness _officeBitness;

    static Context()
    {
        // 1. Initialise the path first
        _excelPath = ExcelHelper.ResolveExcelExecutable();

        // 2. Resolve bitness immediately from the discovered path
        _officeBitness = BinaryInspector.GetExcelBitness(_excelPath);
    }

    public static string ExcelPath
    {
        get => _excelPath;
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Excel path cannot be null or empty.", nameof(value));

            _excelPath = value;
            _officeBitness = BinaryInspector.GetExcelBitness(value);
        }
    }

    public static Bitness OfficeBitness => _officeBitness;
}