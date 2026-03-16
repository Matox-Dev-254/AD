#Requires -Version 5.1
<#
.SYNOPSIS
    AD Manager - Administrador de Active Directory con interfaz gráfica.
.DESCRIPTION
    Entry point principal. La lógica funcional y de UI se carga desde la carpeta modules.
.NOTES
    Autor: Admin
    Fecha: 2026-03-12
#>

$script:ModuleRoot = Join-Path $PSScriptRoot 'modules'
$moduleOrder = @(
  '00.Bootstrap.ps1',
  '10.Xaml.ps1',
  '20.Window.ps1',
  '30.Domain.ps1',
  '40.Theme.ps1',
  '50.Core.ps1',
  '60.UsersGroups.ps1',
  '70.Operations.ps1',
  '80.AttributesStartup.ps1',
  '90.Launch.ps1'
)

foreach ($moduleName in $moduleOrder) {
  $modulePath = Join-Path $script:ModuleRoot $moduleName
  if (-not (Test-Path -Path $modulePath -PathType Leaf)) {
    throw "No se encontró el módulo requerido: $modulePath"
  }

  . $modulePath
}
