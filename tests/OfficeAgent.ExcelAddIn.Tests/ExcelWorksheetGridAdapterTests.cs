using System;
using System.IO;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ExcelWorksheetGridAdapterTests
    {
        [Fact]
        public void NormalizeToObjectMatrixExpandsScalarToRequestedShape()
        {
            var method = GetNormalizeToObjectMatrixMethod();

            var result = (object[,])method.Invoke(null, new object[] { "scalar-value", 2, 3 });

            Assert.Equal(2, result.GetLength(0));
            Assert.Equal(3, result.GetLength(1));
            for (var row = 0; row < result.GetLength(0); row++)
            {
                for (var column = 0; column < result.GetLength(1); column++)
                {
                    Assert.Equal("scalar-value", Convert.ToString(result[row, column]));
                }
            }
        }

        [Fact]
        public void NormalizeToStringMatrixExpandsScalarToRequestedShape()
        {
            var method = GetNormalizeToStringMatrixMethod();

            var result = (string[,])method.Invoke(null, new object[] { "yyyy-mm-dd", 3, 2 });

            Assert.Equal(3, result.GetLength(0));
            Assert.Equal(2, result.GetLength(1));
            for (var row = 0; row < result.GetLength(0); row++)
            {
                for (var column = 0; column < result.GetLength(1); column++)
                {
                    Assert.Equal("yyyy-mm-dd", result[row, column]);
                }
            }
        }

        [Fact]
        public void NormalizeToObjectMatrixConvertsOneBasedComArrayIntoZeroBasedRequestedShape()
        {
            var method = GetNormalizeToObjectMatrixMethod();
            var source = Array.CreateInstance(typeof(object), new[] { 2, 2 }, new[] { 1, 1 });
            source.SetValue("r1c1", 1, 1);
            source.SetValue("r1c2", 1, 2);
            source.SetValue("r2c1", 2, 1);
            source.SetValue("r2c2", 2, 2);

            var result = (object[,])method.Invoke(null, new object[] { source, 2, 2 });

            Assert.Equal(new object[,] { { "r1c1", "r1c2" }, { "r2c1", "r2c2" } }, result);
        }

        [Fact]
        public void NormalizeToStringMatrixConvertsOneBasedComArrayIntoZeroBasedRequestedShape()
        {
            var method = GetNormalizeToStringMatrixMethod();
            var source = Array.CreateInstance(typeof(object), new[] { 2, 2 }, new[] { 1, 1 });
            source.SetValue("General", 1, 1);
            source.SetValue("yyyy-mm-dd", 1, 2);
            source.SetValue("0.00", 2, 1);
            source.SetValue("@", 2, 2);

            var result = (string[,])method.Invoke(null, new object[] { source, 2, 2 });

            Assert.Equal(new[,] { { "General", "yyyy-mm-dd" }, { "0.00", "@" } }, result);
        }

        private static MethodInfo GetNormalizeToObjectMatrixMethod()
        {
            var method = GetGridAdapterType().GetMethod(
                "NormalizeToObjectMatrix",
                BindingFlags.NonPublic | BindingFlags.Static,
                binder: null,
                types: new[] { typeof(object), typeof(int), typeof(int) },
                modifiers: null);

            return method ?? throw new InvalidOperationException("NormalizeToObjectMatrix(object,int,int) was not found.");
        }

        private static MethodInfo GetNormalizeToStringMatrixMethod()
        {
            var method = GetGridAdapterType().GetMethod(
                "NormalizeToStringMatrix",
                BindingFlags.NonPublic | BindingFlags.Static,
                binder: null,
                types: new[] { typeof(object), typeof(int), typeof(int) },
                modifiers: null);

            return method ?? throw new InvalidOperationException("NormalizeToStringMatrix(object,int,int) was not found.");
        }

        private static Type GetGridAdapterType()
        {
            return Assembly.LoadFrom(ResolveAddInAssemblyPath())
                .GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorksheetGridAdapter", throwOnError: true);
        }

        private static string ResolveAddInAssemblyPath()
        {
            return Path.GetFullPath(
                Path.Combine(
                    AppContext.BaseDirectory,
                    "..",
                    "..",
                    "..",
                    "..",
                    "..",
                    "src",
                    "OfficeAgent.ExcelAddIn",
                    "bin",
                    "Debug",
                    "OfficeAgent.ExcelAddIn.dll"));
        }
    }
}
