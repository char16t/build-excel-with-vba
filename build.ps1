param(
    [switch]$enableRegistryFix
)

$XLSM_FILE_NAME = "output.xlsm"

$COMPONENT_TYPE_MODULE = 1
$COMPONENT_TYPE_CLASS = 2
$XL_FILE_FORMAT_MACRO_ENABLED = 52

$missing = [Reflection.Missing]::Value

$excel = New-Object -ComObject Excel.Application

if ($enableRegistryFix.IsPresent) {
    $ExcelVersion = $Excel.Version
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings  -Value 1 -Force | Out-Null
}

$book = $excel.Workbooks.Add($missing)

$files = gci src *.* -rec | where { ! $_.PSIsContainer }

Function AddScriptToBook($book, $file)
{
    $extension = $file.Extension.ToLower()
    $lines = New-Object System.Collections.ArrayList

    ForEach ($line in [IO.File]::ReadAllLines($file.FullName))
    {
        $ignore = $lines.Add($line)
    }

    While ($lines[0].StartsWith("VERSION") -or $lines[0].StartsWith("BEGIN") -or $lines[0].StartsWith("  ") -or $lines[0].StartsWith("END") -or $lines[0].StartsWith("Attribute"))
    {
        $ignore = $lines.RemoveAt(0)
    }
    
    $code = [String]::Join("`r`n", $lines.ToArray())
	
	
    $module = $book.VBProject.VBComponents.Add(1)
    $module.CodeModule.AddFromString($code)
    $module.Name = [IO.Path]::GetFileNameWithoutExtension($file.FullName)
}

$files = gci src *.* -rec | where { ! $_.PSIsContainer }

ForEach ($file in $files)
{
    AddScriptToBook $book $file
}

$cwd = (Resolve-Path .\).Path
$build = [IO.Path]::Combine($cwd, "build")
MkDir -Force $build > $null

$filename = [IO.Path]::Combine($build, $XLSM_FILE_NAME)

$excel.DisplayAlerts = $false
$book.SaveAs($filename, $XL_FILE_FORMAT_MACRO_ENABLED)
$excel.DisplayAlerts = $true

if ($enableRegistryFix.IsPresent) {
    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM
    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings
}
$excel.Quit()
