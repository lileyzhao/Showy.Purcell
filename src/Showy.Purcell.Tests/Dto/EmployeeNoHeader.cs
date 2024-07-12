namespace Showy.Purcell.Tests.Dto;

internal class EmployeeNoHeader1
{
    [ExcelHeader(IndexName = "A")] public int EmpId { get; set; }

    [ExcelHeader(IndexName = "B")] public string? Xing { get; set; }

    [ExcelHeader(IndexName = "C")] public string? XingMing { get; set; }
}

internal class EmployeeNoHeader2
{
    [ExcelHeader(0)] public int EmpId { get; set; }

    [ExcelHeader(1)] public string? Xing { get; set; }

    [ExcelHeader(2)] public string? XingMing { get; set; }
}