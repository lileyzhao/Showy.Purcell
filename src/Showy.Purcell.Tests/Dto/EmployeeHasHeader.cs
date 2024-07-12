namespace Showy.Purcell.Tests.Dto;

internal class EmployeeHasHeader
{
    [ExcelHeader("员工编号", "员工ID")] public int EmpId { get; set; }

    [ExcelHeader("姓氏")] public string? LastName { get; set; }

    [ExcelHeader("姓名")] public string? RealName { get; set; }
}