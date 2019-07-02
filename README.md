**Enable PowerShell scripts**

If you can't run the build script at all, use the following command to temporarily enable PowerShell scripts for the duration of your PowerShell session:

```
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process
```

**Trust access to the VBA project object model**

Before you will execute build script you should trust access to the VBA project object model. When you run build script with `-enableRegistryFix` this code will be executed before working with document:

```
$ExcelVersion = $Excel.Version
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings  -Value 1 -Force | Out-Null
```

this code will be executed after working with document:

```
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings
```

If you have error when `$book.VBProject.VBComponents.Add(1)` is `NULL` you should use `-enableRegistryFix` flag:

```
.\build.ps1 -enableRegistryFix
```
