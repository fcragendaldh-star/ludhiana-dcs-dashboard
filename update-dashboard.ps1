param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath
)

python .\scripts\update_dashboard_from_excel.py $ExcelPath
