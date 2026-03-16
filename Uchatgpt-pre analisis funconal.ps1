#Requires -Version 5.1
<#
.SYNOPSIS
    AD Manager - Administrador de Active Directory con interfaz gráfica
.DESCRIPTION
    Aplicación WPF para gestionar usuarios, grupos y OUs de Active Directory.
    Requiere módulo ActiveDirectory (RSAT).
.NOTES
    Autor: Admin
    Fecha: 2026-02-19
#>

# Si estamos en PowerShell 7+ (MTA), relanzar con Windows PowerShell (STA)
if ($PSVersionTable.PSVersion.Major -ge 7 -or [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
  $wpExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  if (Test-Path $wpExe) {
    Start-Process $wpExe -ArgumentList "-ExecutionPolicy Bypass -NoProfile -STA -File `"$PSCommandPath`"" -NoNewWindow -Wait
    return
  }
  else {
    Write-Warning "No se encontró Windows PowerShell. Ejecutá con: powershell.exe -STA -File `"$PSCommandPath`""
    return
  }
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Verificar/instalar módulo AD (offline/local)
$script:ADModuleAvailable = $false
$script:ADInstallAttempted = $false
$script:SimulationMode = $false
$script:WhatIfPreference = $false

function Test-IsAdministrator {
  try {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
  }
  catch {
    return $false
  }
}

function Install-RSATFromWindowsUpdate {
  $capabilities = @(
    'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0',
    'Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0'
  )

  foreach ($capName in $capabilities) {
    try {
      if (Get-Command Add-WindowsCapability -ErrorAction SilentlyContinue) {
        Add-WindowsCapability -Online -Name $capName -ErrorAction Stop | Out-Null
      }
      else {
        $dismArgs = "/online /add-capability /capabilityname:$capName /quiet /norestart"
        $proc = Start-Process -FilePath dism.exe -ArgumentList $dismArgs -Wait -PassThru -NoNewWindow
        if (($proc.ExitCode -ne 0) -and ($proc.ExitCode -ne 3010)) {
          throw "DISM devolvió ExitCode $($proc.ExitCode) instalando $capName."
        }
      }
    }
    catch {
      return $false
    }
  }

  return [bool](Get-Module -ListAvailable -Name ActiveDirectory)
}

function Ensure-ADModuleAvailable {
  if (Get-Module -ListAvailable -Name ActiveDirectory) {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    $script:ADModuleAvailable = $true
    return
  }

  $script:ADModuleAvailable = $false
  $script:ADInstallAttempted = $true

  if (-not (Test-IsAdministrator)) {
    Write-Warning "RSAT no está instalado. Para instalar automáticamente desde Windows Update/Internet, ejecutá este script como administrador."
    return
  }

  if (Install-RSATFromWindowsUpdate) {
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
      Import-Module ActiveDirectory -ErrorAction SilentlyContinue
      $script:ADModuleAvailable = $true
      return
    }
  }
}

Ensure-ADModuleAvailable

function Assert-ADModule {
  if (-not $script:ADModuleAvailable) {
    [System.Windows.MessageBox]::Show(
      "El módulo ActiveDirectory (RSAT) no está instalado.`n`nIntento automático:`n- Ejecutar como administrador`n- Instalar desde Windows Update/Internet`n`nSi no se instala, podés hacerlo manualmente desde Características opcionales.`n`nLa GUI se puede ver pero las funciones AD no van a funcionar.",
      "Módulo AD no disponible", "OK", "Warning"
    ) | Out-Null
    return $false
  }
  return $true
}

# ── XAML ──────────────────────────────────────────────────────────────────────
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD Manager - Administrador de Active Directory | Develop by Matias Vaccari"
        Width="1220" Height="820" MinWidth="980" MinHeight="700"
        WindowStartupLocation="CenterScreen"
        Background="#1e1e2e" Foreground="#cdd6f4"
        FontFamily="Segoe UI">
  <Window.Resources>
    <Style x:Key="PanelCard" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource ThemePanelBg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemePanelBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="12"/>
      <Setter Property="Padding" Value="14,12"/>
    </Style>
    <Style x:Key="BadgeText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource ThemeChipFg}"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{DynamicResource ThemeInputBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeInputFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,7"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="MinHeight" Value="34"/>
      <Setter Property="Margin" Value="0,2,0,2"/>
    </Style>
    <Style TargetType="PasswordBox">
      <Setter Property="Background" Value="{DynamicResource ThemeInputBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeInputFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,7"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="MinHeight" Value="34"/>
      <Setter Property="Margin" Value="0,2,0,2"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Background" Value="{DynamicResource ThemeInputBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeInputFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeBorder}"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="MinHeight" Value="34"/>
      <Setter Property="Margin" Value="0,2,0,2"/>
    </Style>
    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource ThemeBtnPrimaryBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeBtnPrimaryFg}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="16,9"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeBtnPrimaryBorder}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="MinHeight" Value="36"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="10" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{DynamicResource ThemeBtnPrimaryHoverBg}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="{DynamicResource ThemeBtnPrimaryPressedBg}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnDanger" TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource ThemeBtnDangerBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeBtnDangerFg}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="16,9"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeBtnDangerBorder}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="MinHeight" Value="36"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="10" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{DynamicResource ThemeBtnDangerHoverBg}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="{DynamicResource ThemeBtnDangerPressedBg}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnSuccess" TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource ThemeBtnSuccessBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeBtnSuccessFg}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="16,9"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeBtnSuccessBorder}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="MinHeight" Value="36"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="10" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{DynamicResource ThemeBtnSuccessHoverBg}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="{DynamicResource ThemeBtnSuccessPressedBg}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnWarn" TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource ThemeBtnWarnBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeBtnWarnFg}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="16,9"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeBtnWarnBorder}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="MinHeight" Value="36"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="10" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{DynamicResource ThemeBtnWarnHoverBg}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="{DynamicResource ThemeBtnWarnPressedBg}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="Label">
      <Setter Property="Foreground" Value="{DynamicResource ThemeLabelFg}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Padding" Value="0,4,0,2"/>
    </Style>
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="{DynamicResource ThemeGridBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeGridFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeGridBorder}"/>
      <Setter Property="RowBackground" Value="{DynamicResource ThemeGridRowBg}"/>
      <Setter Property="AlternatingRowBackground" Value="{DynamicResource ThemeGridRowAltBg}"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource ThemeGridLine}"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="SelectionMode" Value="Single"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowHeaderWidth" Value="0"/>
      <Setter Property="Margin" Value="0,4,0,0"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="{DynamicResource ThemeGridHeaderBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeGridHeaderFg}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeGridBorder}"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Setter Property="MinHeight" Value="32"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{DynamicResource ThemeSelectionBg}"/>
          <Setter Property="Foreground" Value="{DynamicResource ThemeSelectionFg}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{DynamicResource ThemeSelectionBg}"/>
          <Setter Property="Foreground" Value="{DynamicResource ThemeSelectionFg}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="TabControl">
      <Setter Property="Background" Value="{DynamicResource ThemeWindowBg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemePanelBorder}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="{DynamicResource ThemeTabFg}"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Padding" Value="18,10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Grid Margin="3,0,3,0">
              <Border x:Name="Bd" Background="{DynamicResource ThemeTabBg}" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1" CornerRadius="10,10,0,0" Padding="{TemplateBinding Padding}">
                <ContentPresenter ContentSource="Header" HorizontalAlignment="Center"/>
              </Border>
              <Border x:Name="AccentBar" Height="3" Background="{DynamicResource ThemeAccent}" VerticalAlignment="Bottom" Margin="14,0,14,0" Opacity="0"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeTabSelectedBg}"/>
                <Setter Property="Foreground" Value="{DynamicResource ThemeTabSelectedFg}"/>
                <Setter Property="FontWeight" Value="Bold"/>
                <Setter TargetName="AccentBar" Property="Opacity" Value="1"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource ThemeTabHoverBg}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="{DynamicResource ThemeTextPrimary}"/>
      <Setter Property="FontSize" Value="13"/>
    </Style>
    <Style TargetType="TreeView">
      <Setter Property="Background" Value="{DynamicResource ThemeTreeBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ThemeTreeFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemeTreeBorder}"/>
    </Style>
    <Style TargetType="TreeViewItem">
      <Setter Property="Foreground" Value="{DynamicResource ThemeTreeItemFg}"/>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Foreground" Value="{DynamicResource ThemeGroupBoxFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ThemePanelBorder}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="GroupBox">
            <Grid Margin="0,8,0,0" SnapsToDevicePixels="True">
              <Border Margin="0,10,0,0"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="1"
                      CornerRadius="12"
                      Background="{DynamicResource ThemePanelBg}"/>
              <Border Background="{DynamicResource ThemeWindowBg}"
                      HorizontalAlignment="Left"
                      Margin="14,0,0,0"
                      Padding="10,0">
                <ContentPresenter ContentSource="Header"
                                  RecognizesAccessKey="True"
                                  TextElement.Foreground="{TemplateBinding Foreground}"/>
              </Border>
              <ContentPresenter Margin="12,22,12,12"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid Margin="14">
    <DockPanel>
    <Border x:Name="borderHeader" DockPanel.Dock="Top" Background="#181825" Padding="20,18,20,16" CornerRadius="18" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1">
      <DockPanel>
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Top">
          <Border Background="{DynamicResource ThemeChipBg}" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1" CornerRadius="999" Padding="10,4" Margin="0,0,10,0">
            <TextBlock Text="Directorio y auditoria" Style="{StaticResource BadgeText}"/>
          </Border>
          <Button x:Name="btnThemeToggle" Content="☀️ Claro" Style="{StaticResource BtnPrimary}" FontSize="12" Padding="12,6" VerticalAlignment="Center"/>
        </StackPanel>
        <StackPanel VerticalAlignment="Center">
          <TextBlock x:Name="txtTitle" Text="AD Manager" FontSize="28" FontWeight="Bold" Foreground="#89b4fa" FontFamily="Bahnschrift SemiBold"/>
          <TextBlock x:Name="txtSubtitle" Text="Centro de administracion para usuarios, grupos, OUs y procesos masivos de Active Directory" FontSize="13" Foreground="#6c7086" Margin="0,4,0,0"/>
          <WrapPanel Margin="0,14,0,0">
            <Border Background="{DynamicResource ThemeChipBg}" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1" CornerRadius="999" Padding="10,4" Margin="0,0,8,0">
              <TextBlock Text="Usuarios" Style="{StaticResource BadgeText}"/>
            </Border>
            <Border Background="{DynamicResource ThemeChipBg}" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1" CornerRadius="999" Padding="10,4" Margin="0,0,8,0">
              <TextBlock Text="Grupos" Style="{StaticResource BadgeText}"/>
            </Border>
            <Border Background="{DynamicResource ThemeChipBg}" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1" CornerRadius="999" Padding="10,4" Margin="0,0,8,0">
              <TextBlock Text="OUs" Style="{StaticResource BadgeText}"/>
            </Border>
            <Border Background="{DynamicResource ThemeChipBg}" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1" CornerRadius="999" Padding="10,4">
              <TextBlock Text="Lotes y atributos" Style="{StaticResource BadgeText}"/>
            </Border>
          </WrapPanel>
        </StackPanel>
      </DockPanel>
    </Border>

    <Border DockPanel.Dock="Top" Height="4" Margin="10,10,10,0" CornerRadius="999">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
          <GradientStop Color="#89b4fa" Offset="0"/>
          <GradientStop Color="#cba6f7" Offset="0.5"/>
          <GradientStop Color="#f38ba8" Offset="1"/>
        </LinearGradientBrush>
      </Border.Background>
    </Border>

    <Border x:Name="borderDomainBar" DockPanel.Dock="Top" Background="#232336" Padding="16,12" CornerRadius="14" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1" Margin="0,12,0,0">
      <WrapPanel>
        <TextBlock x:Name="txtDomainLabel" Text="Dominio:" Foreground="#bac2de" VerticalAlignment="Center" FontSize="12" Margin="0,0,6,0"/>
        <ComboBox x:Name="cbDomainSelect" Width="260" FontSize="12"/>
        <Button x:Name="btnDomainConnect" Content="➕ Conectar" Style="{StaticResource BtnSuccess}" FontSize="11" Padding="10,4" Margin="8,0,0,0"/>
        <Button x:Name="btnDomainDisconnect" Content="❌ Desconectar" Style="{StaticResource BtnDanger}" FontSize="11" Padding="10,4" Margin="4,0,0,0"/>
        <CheckBox x:Name="chkSimulationMode" Content="🧪 Simulación (WhatIf)" Margin="12,0,0,0" VerticalAlignment="Center" FontSize="12"/>
        <Border Background="{DynamicResource ThemeChipBg}" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1" CornerRadius="999" Padding="10,4" Margin="12,0,0,0">
          <TextBlock x:Name="txtDomainInfo" Text="" Foreground="#6c7086" VerticalAlignment="Center" FontSize="11"/>
        </Border>
      </WrapPanel>
    </Border>

    <Border DockPanel.Dock="Bottom" Background="#181825" Padding="14,10" CornerRadius="14" BorderBrush="{DynamicResource ThemePanelBorder}" BorderThickness="1" Margin="0,12,0,0">
      <TextBlock x:Name="txtStatus" Text="Listo." Foreground="#a6adc8" FontSize="12"/>
    </Border>

    <TabControl x:Name="mainTabControl" Background="#1e1e2e" BorderBrush="#45475a" Margin="0,14,0,0" Padding="0,8">

      <TabItem Header="👤 Usuarios">
        <Grid Margin="10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Border x:Name="borderUserSearchPanel" Grid.Row="0" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="12,8" Margin="0,0,0,8">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="🔍" FontSize="16" VerticalAlignment="Center" Margin="0,0,6,0"/>
              <TextBox x:Name="txtUserSearch" Width="300" ToolTip="Buscar por nombre, SAM o email"/>
              <Button x:Name="btnUserSearch" Content="Buscar" Style="{StaticResource BtnPrimary}" Margin="8,0,0,0"/>
              <Button x:Name="btnUserRefresh" Content="🔄" Style="{StaticResource BtnPrimary}" Margin="4,0,0,0" ToolTip="Actualizar"/>
            </StackPanel>
          </Border>

          <DataGrid x:Name="dgUsers" Grid.Row="1">
            <DataGrid.Columns>
              <DataGridTextColumn Header="SAMAccountName" Binding="{Binding SamAccountName}" Width="130"/>
              <DataGridTextColumn Header="Nombre" Binding="{Binding Name}" Width="180"/>
              <DataGridTextColumn Header="Email" Binding="{Binding EmailAddress}" Width="200"/>
              <DataGridTextColumn Header="Habilitado" Binding="{Binding Enabled}" Width="80"/>
              <DataGridTextColumn Header="Bloqueado" Binding="{Binding LockedOut}" Width="80"/>
              <DataGridTextColumn Header="OU" Binding="{Binding OUPath}" Width="*"/>
            </DataGrid.Columns>
          </DataGrid>

          <Border x:Name="borderUserActionsPanel" Grid.Row="2" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="10,8" Margin="0,8,0,0">
            <WrapPanel>
              <Button x:Name="btnUserNew" Content="➕ Crear Usuario" Style="{StaticResource BtnSuccess}" Margin="0,0,6,4"/>
              <Button x:Name="btnUserCopyProfile" Content="📄 Copiar Perfil" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnUserEdit" Content="✏️ Modificar" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnUserResetPwd" Content="🔑 Resetear Contraseña" Style="{StaticResource BtnWarn}" Margin="0,0,6,4"/>
              <Button x:Name="btnUserDisable" Content="🚫 Deshabilitar" Style="{StaticResource BtnWarn}" Margin="0,0,6,4"/>
              <Button x:Name="btnUserEnable" Content="✅ Habilitar" Style="{StaticResource BtnSuccess}" Margin="0,0,6,4"/>
              <Button x:Name="btnUserUnlock" Content="🔓 Desbloquear" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnUserDelete" Content="🗑️ Eliminar" Style="{StaticResource BtnDanger}" Margin="0,0,6,4"/>
            </WrapPanel>
          </Border>
        </Grid>
      </TabItem>
      <TabItem Header="👥 Grupos">
        <Grid Margin="10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Border x:Name="borderGroupSearchPanel" Grid.Row="0" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="12,8" Margin="0,0,0,8">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="🔍" FontSize="16" VerticalAlignment="Center" Margin="0,0,6,0"/>
              <TextBox x:Name="txtGroupSearch" Width="300" ToolTip="Buscar por nombre de grupo"/>
              <Button x:Name="btnGroupSearch" Content="Buscar" Style="{StaticResource BtnPrimary}" Margin="8,0,0,0"/>
            </StackPanel>
          </Border>

          <DataGrid x:Name="dgGroups" Grid.Row="1">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Nombre" Binding="{Binding Name}" Width="200"/>
              <DataGridTextColumn Header="SAM" Binding="{Binding SamAccountName}" Width="150"/>
              <DataGridTextColumn Header="Categoría" Binding="{Binding GroupCategory}" Width="100"/>
              <DataGridTextColumn Header="Ámbito" Binding="{Binding GroupScope}" Width="100"/>
              <DataGridTextColumn Header="Descripción" Binding="{Binding Description}" Width="*"/>
            </DataGrid.Columns>
          </DataGrid>

          <Border x:Name="borderGroupActionsPanel" Grid.Row="2" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="10,8" Margin="0,8,0,0">
            <WrapPanel>
              <Button x:Name="btnGroupNew" Content="➕ Crear Grupo" Style="{StaticResource BtnSuccess}" Margin="0,0,6,4"/>
              <Button x:Name="btnGroupAddMember" Content="👤+ Agregar Miembro" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnGroupRemoveMember" Content="👤- Quitar Miembro" Style="{StaticResource BtnWarn}" Margin="0,0,6,4"/>
              <Button x:Name="btnGroupMembers" Content="📋 Ver Miembros" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnGroupDelete" Content="🗑️ Eliminar Grupo" Style="{StaticResource BtnDanger}" Margin="0,0,6,4"/>
            </WrapPanel>
          </Border>
        </Grid>
      </TabItem>
<TabItem Header="🔀 Mover">
        <Grid Margin="10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <GroupBox Grid.Column="0" Header="OU Origen y Objetos" Margin="0,0,4,0">
            <StackPanel Margin="8">
              <Label Content="Seleccionar OU origen:"/>
              <Button x:Name="btnMoveRefreshSourceOU" Content="🔄 Recargar OUs origen" Style="{StaticResource BtnPrimary}" Margin="0,0,0,6" HorizontalAlignment="Left"/>
              <TreeView x:Name="tvMoveSourceOUs" Height="170"/>
              <Button x:Name="btnMoveLoadFromSource" Content="📥 Cargar objetos de OU origen" Style="{StaticResource BtnPrimary}" Margin="0,8,0,6" HorizontalAlignment="Left"/>

              <DataGrid x:Name="dgMoveObjects" Height="180" Margin="0,4,0,0" SelectionMode="Extended">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Nombre" Binding="{Binding Name}" Width="160"/>
                  <DataGridTextColumn Header="Tipo" Binding="{Binding ObjectClass}" Width="110"/>
                  <DataGridTextColumn Header="OU Actual" Binding="{Binding OUPath}" Width="*"/>
                </DataGrid.Columns>
              </DataGrid>
            </StackPanel>
          </GroupBox>

          <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="8">
            <Button x:Name="btnMoveExecute" Content="➡️ Mover" Style="{StaticResource BtnWarn}" FontSize="15" Padding="20,12"/>
          </StackPanel>

          <GroupBox Grid.Column="2" Header="OU Destino" Margin="4,0,0,0">
            <StackPanel Margin="8">
              <Label Content="Seleccionar OU destino:"/>
              <Button x:Name="btnMoveRefreshTargetOU" Content="🔄 Recargar OUs destino" Style="{StaticResource BtnPrimary}" Margin="0,0,0,6" HorizontalAlignment="Stretch"/>
              <Grid Margin="0,0,0,6">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button x:Name="btnMoveNewOU" Grid.Column="0" Content="➕ Crear OU" Style="{StaticResource BtnSuccess}" Margin="0,0,4,0"/>
                <Button x:Name="btnMoveDeleteOU" Grid.Column="1" Content="🗑️ Eliminar OU" Style="{StaticResource BtnDanger}" Margin="4,0,0,0"/>
              </Grid>
              <TreeView x:Name="tvMoveTargetOUs" Height="180"/>
              <Button x:Name="btnMoveLoadFromTarget" Content="📥 Cargar objetos de OU destino" Style="{StaticResource BtnPrimary}" Margin="0,8,0,6" HorizontalAlignment="Left"/>
              <DataGrid x:Name="dgMoveTargetObjects" Height="180" Margin="0,4,0,0">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Nombre" Binding="{Binding Name}" Width="160"/>
                  <DataGridTextColumn Header="Tipo" Binding="{Binding ObjectClass}" Width="110"/>
                  <DataGridTextColumn Header="OU Actual" Binding="{Binding OUPath}" Width="*"/>
                </DataGrid.Columns>
              </DataGrid>
            </StackPanel>
          </GroupBox>
        </Grid>
      </TabItem>

      <TabItem Header="📦 Lote CSV">
        <Grid Margin="10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="12,8" Margin="0,0,0,8">
            <StackPanel>
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="📄" FontSize="16" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBox x:Name="txtBatchCsvPath" Width="560" ToolTip="Ruta al archivo CSV"/>
                <Button x:Name="btnBatchBrowseCsv" Content="Examinar" Style="{StaticResource BtnPrimary}" Margin="8,0,0,0"/>
              </StackPanel>
              <TextBlock Text="Acciones soportadas: Alta, Modificar, Baja, Mover (columna requerida: Action)." Foreground="#6c7086" FontSize="11" Margin="0,6,0,0"/>
            </StackPanel>
          </Border>

          <Border Grid.Row="1" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="10,8" Margin="0,0,0,8">
            <WrapPanel>
              <Button x:Name="btnBatchTemplate" Content="🧾 Plantilla CSV" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnBatchTemplateEmpty" Content="📄 Plantilla Vacía" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnBatchPreview" Content="👁️ Preview" Style="{StaticResource BtnWarn}" Margin="0,0,6,4"/>
              <Button x:Name="btnBatchExecute" Content="▶️ Ejecutar Lote" Style="{StaticResource BtnSuccess}" Margin="0,0,6,4"/>
              <Button x:Name="btnBatchExportReport" Content="💾 Exportar Reporte" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
            </WrapPanel>
          </Border>

          <DataGrid x:Name="dgBatchResults" Grid.Row="2">
            <DataGrid.Columns>
              <DataGridTextColumn Header="#" Binding="{Binding Row}" Width="60"/>
              <DataGridTextColumn Header="Acción" Binding="{Binding Action}" Width="90"/>
              <DataGridTextColumn Header="Identidad" Binding="{Binding Identity}" Width="180"/>
              <DataGridTextColumn Header="Estado" Binding="{Binding Status}" Width="90"/>
              <DataGridTextColumn Header="Detalle" Binding="{Binding Message}" Width="*"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </TabItem>

      <TabItem Header="📝 Atributos">
        <Grid Margin="10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Border x:Name="borderAttrSearchPanel" Grid.Row="0" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="12,8" Margin="0,0,0,8">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="🔍" FontSize="16" VerticalAlignment="Center" Margin="0,0,6,0"/>
              <TextBox x:Name="txtAttrSearch" Width="250" ToolTip="Nombre de usuario o grupo"/>
              <Button x:Name="btnAttrSearch" Content="Cargar Atributos" Style="{StaticResource BtnPrimary}" Margin="8,0,0,0"/>
              <Button x:Name="btnAttrClear" Content="Limpiar" Style="{StaticResource BtnWarn}" Margin="6,0,0,0" Padding="10,5" FontSize="11"/>
              <TextBlock x:Name="txtAttrDN" Text="" Foreground="#6c7086" VerticalAlignment="Center" Margin="12,0,0,0" FontSize="11" TextTrimming="CharacterEllipsis" MaxWidth="350"/>
            </StackPanel>
          </Border>

          <DataGrid x:Name="dgAttributes" Grid.Row="1" CanUserSortColumns="True">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Atributo" Binding="{Binding Attribute}" Width="220" IsReadOnly="True"/>
              <DataGridTextColumn Header="Valor" Binding="{Binding Value}" Width="*"/>
            </DataGrid.Columns>
          </DataGrid>

          <Border x:Name="borderAttrActionsPanel" Grid.Row="2" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="10,8" Margin="0,8,0,0">
            <StackPanel Orientation="Horizontal">
              <Button x:Name="btnAttrSave" Content="💾 Guardar Cambios" Style="{StaticResource BtnSuccess}" Margin="0,0,6,0"/>
              <TextBlock Text="Seleccioná un atributo en la grilla, editá el valor y presioná Guardar." Foreground="#6c7086" VerticalAlignment="Center" FontSize="12"/>
            </StackPanel>
          </Border>
        </Grid>
      </TabItem>

      <TabItem Header="⚙️ Configuración">
        <Grid Margin="10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Border x:Name="borderConfigHeaderPanel" Grid.Row="0" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="12,10" Margin="0,0,0,8">
            <StackPanel>
              <TextBlock Text="Configuración de Auditoría Central" FontSize="15" FontWeight="SemiBold" Foreground="#89b4fa"/>
              <TextBlock Text="Definí destino central para logs compartidos entre equipos." Foreground="#6c7086" FontSize="12" Margin="0,4,0,0"/>
            </StackPanel>
          </Border>

          <Border x:Name="borderConfigBodyPanel" Grid.Row="1" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="12,10">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" CanContentScroll="True">
              <StackPanel>
                <CheckBox x:Name="chkAuditEnableCentral" Content="Habilitar auditoría centralizada" Margin="0,0,0,8"/>

                <Label Content="Modo de destino:"/>
                <ComboBox x:Name="cbAuditMode" Width="180" HorizontalAlignment="Left"/>

                <Label Content="Ruta central (FileShare):"/>
                <StackPanel Orientation="Horizontal">
                  <TextBox x:Name="txtAuditCentralFilePath" Width="760" ToolTip="Ejemplo: \\SERVIDOR\AD-Audit\AD_Audit_Central.csv"/>
                  <Button x:Name="btnAuditBrowseCentralFile" Content="Examinar" Style="{StaticResource BtnPrimary}" Margin="8,0,0,0" Padding="10,4"/>
                </StackPanel>

                <Label Content="Connection String (SqlServer):"/>
                <TextBox x:Name="txtAuditSqlConnection" ToolTip="Ejemplo: Server=SQL01;Database=AdminAD;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;"/>

                <Label Content="Tabla SQL (SqlServer):"/>
                <TextBox x:Name="txtAuditSqlTable" Width="300" HorizontalAlignment="Left"/>

                <Label Content="Diagnóstico:" Margin="0,8,0,0"/>
                <TextBox x:Name="txtAuditDiagnostics"
                         IsReadOnly="True"
                         AcceptsReturn="True"
                         TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto"
                         MinHeight="150"
                         FontFamily="Consolas"/>
                <Button x:Name="btnAuditDiagScrollDown"
                        Content="Bajar Diagnóstico"
                        Style="{StaticResource BtnPrimary}"
                        Margin="0,6,0,0"
                        HorizontalAlignment="Right"/>
              </StackPanel>
            </ScrollViewer>
          </Border>

          <Border x:Name="borderConfigActionsPanel" Grid.Row="2" Background="{DynamicResource ThemeSurfaceBg}" CornerRadius="6" Padding="10,8" Margin="0,8,0,0">
            <WrapPanel>
              <Button x:Name="btnAuditApplyConfig" Content="Aplicar Configuración" Style="{StaticResource BtnSuccess}" Margin="0,0,6,4"/>
              <Button x:Name="btnAuditSaveConfig" Content="Guardar Configuración" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnAuditLoadConfig" Content="Cargar Configuración" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnAuditDiagnoseConfig" Content="Diagnosticar" Style="{StaticResource BtnPrimary}" Margin="0,0,6,4"/>
              <Button x:Name="btnAuditTestConfig" Content="Probar Destino Central" Style="{StaticResource BtnWarn}" Margin="0,0,6,4"/>
            </WrapPanel>
          </Border>
        </Grid>
      </TabItem>

    </TabControl>
  </DockPanel>
  </Grid>
</Window>
"@

# ── CARGAR VENTANA ────────────────────────────────────────────────────────────
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

function Get-Control([string]$name) { $window.FindName($name) }

$txtStatus = Get-Control 'txtStatus'
$txtUserSearch = Get-Control 'txtUserSearch'
$btnUserSearch = Get-Control 'btnUserSearch'
$btnUserRefresh = Get-Control 'btnUserRefresh'
$dgUsers = Get-Control 'dgUsers'
$btnUserNew = Get-Control 'btnUserNew'
$btnUserCopyProfile = Get-Control 'btnUserCopyProfile'
$btnUserEdit = Get-Control 'btnUserEdit'
$btnUserResetPwd = Get-Control 'btnUserResetPwd'
$btnUserDisable = Get-Control 'btnUserDisable'
$btnUserEnable = Get-Control 'btnUserEnable'
$btnUserUnlock = Get-Control 'btnUserUnlock'
$btnUserDelete = Get-Control 'btnUserDelete'

$txtGroupSearch = Get-Control 'txtGroupSearch'
$btnGroupSearch = Get-Control 'btnGroupSearch'
$dgGroups = Get-Control 'dgGroups'
$btnGroupNew = Get-Control 'btnGroupNew'
$btnGroupAddMember = Get-Control 'btnGroupAddMember'
$btnGroupRemoveMember = Get-Control 'btnGroupRemoveMember'
$btnGroupMembers = Get-Control 'btnGroupMembers'
$btnGroupDelete = Get-Control 'btnGroupDelete'

$dgMoveObjects = Get-Control 'dgMoveObjects'
$btnMoveExecute = Get-Control 'btnMoveExecute'
$btnMoveRefreshSourceOU = Get-Control 'btnMoveRefreshSourceOU'
$btnMoveRefreshTargetOU = Get-Control 'btnMoveRefreshTargetOU'
$btnMoveLoadFromSource = Get-Control 'btnMoveLoadFromSource'
$btnMoveLoadFromTarget = Get-Control 'btnMoveLoadFromTarget'
$btnMoveNewOU = Get-Control 'btnMoveNewOU'
$btnMoveDeleteOU = Get-Control 'btnMoveDeleteOU'
$dgMoveTargetObjects = Get-Control 'dgMoveTargetObjects'
$tvMoveSourceOUs = Get-Control 'tvMoveSourceOUs'
$tvMoveTargetOUs = Get-Control 'tvMoveTargetOUs'

$txtBatchCsvPath = Get-Control 'txtBatchCsvPath'
$btnBatchBrowseCsv = Get-Control 'btnBatchBrowseCsv'
$btnBatchTemplate = Get-Control 'btnBatchTemplate'
$btnBatchTemplateEmpty = Get-Control 'btnBatchTemplateEmpty'
$btnBatchPreview = Get-Control 'btnBatchPreview'
$btnBatchExecute = Get-Control 'btnBatchExecute'
$btnBatchExportReport = Get-Control 'btnBatchExportReport'
$dgBatchResults = Get-Control 'dgBatchResults'

$txtAttrSearch = Get-Control 'txtAttrSearch'
$btnAttrSearch = Get-Control 'btnAttrSearch'
$btnAttrClear = Get-Control 'btnAttrClear'
$txtAttrDN = Get-Control 'txtAttrDN'
$dgAttributes = Get-Control 'dgAttributes'
$btnAttrSave = Get-Control 'btnAttrSave'
$chkAuditEnableCentral = Get-Control 'chkAuditEnableCentral'
$cbAuditMode = Get-Control 'cbAuditMode'
$txtAuditCentralFilePath = Get-Control 'txtAuditCentralFilePath'
$btnAuditBrowseCentralFile = Get-Control 'btnAuditBrowseCentralFile'
$txtAuditSqlConnection = Get-Control 'txtAuditSqlConnection'
$txtAuditSqlTable = Get-Control 'txtAuditSqlTable'
$txtAuditDiagnostics = Get-Control 'txtAuditDiagnostics'
$btnAuditDiagScrollDown = Get-Control 'btnAuditDiagScrollDown'
$btnAuditApplyConfig = Get-Control 'btnAuditApplyConfig'
$btnAuditSaveConfig = Get-Control 'btnAuditSaveConfig'
$btnAuditLoadConfig = Get-Control 'btnAuditLoadConfig'
$btnAuditDiagnoseConfig = Get-Control 'btnAuditDiagnoseConfig'
$btnAuditTestConfig = Get-Control 'btnAuditTestConfig'
 $borderUserSearchPanel = Get-Control 'borderUserSearchPanel'
 $borderUserActionsPanel = Get-Control 'borderUserActionsPanel'
 $borderGroupSearchPanel = Get-Control 'borderGroupSearchPanel'
 $borderGroupActionsPanel = Get-Control 'borderGroupActionsPanel'
 $borderAttrSearchPanel = Get-Control 'borderAttrSearchPanel'
 $borderAttrActionsPanel = Get-Control 'borderAttrActionsPanel'
 $borderConfigHeaderPanel = Get-Control 'borderConfigHeaderPanel'
 $borderConfigBodyPanel = Get-Control 'borderConfigBodyPanel'
 $borderConfigActionsPanel = Get-Control 'borderConfigActionsPanel'

$borderHeader = Get-Control 'borderHeader'
$txtTitle = Get-Control 'txtTitle'
$txtSubtitle = Get-Control 'txtSubtitle'
$btnThemeToggle = Get-Control 'btnThemeToggle'

$borderDomainBar = Get-Control 'borderDomainBar'
$txtDomainLabel = Get-Control 'txtDomainLabel'
$cbDomainSelect = Get-Control 'cbDomainSelect'
$btnDomainConnect = Get-Control 'btnDomainConnect'
$btnDomainDisconnect = Get-Control 'btnDomainDisconnect'
$chkSimulationMode = Get-Control 'chkSimulationMode'
$txtDomainInfo = Get-Control 'txtDomainInfo'

$borderStatus = $txtStatus.Parent
$tabControl = Get-Control 'mainTabControl'

# ── SISTEMA DE DOMINIOS ──────────────────────────────────────────────────────
$global:domainConnections = @{}
$global:activeDomain = $null
$script:ADDomainCache = @{}
$script:LocalDomainDns = ''
$script:LocalDomainLabel = '(Local)'
$script:DomainConnectionsConfigPath = Join-Path $PSScriptRoot 'domain.connections.json'
$script:StartupDomainChecks = @()
$script:StartupDomainFailures = @()

function Resolve-LocalDomainDns {
  if (-not [string]::IsNullOrWhiteSpace([string]$env:USERDNSDOMAIN)) { return [string]$env:USERDNSDOMAIN }
  if (-not [string]::IsNullOrWhiteSpace([string]$env:USERDOMAIN)) { return [string]$env:USERDOMAIN }
  try {
    $d = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    if ($d -and $d.Name) { return [string]$d.Name }
  }
  catch {}
  return ''
}

function Set-LocalDomainIdentity([string]$dnsRoot) {
  if ([string]::IsNullOrWhiteSpace($dnsRoot)) { $dnsRoot = 'Local sin dominio detectado' }
  $script:LocalDomainDns = $dnsRoot
  $script:LocalDomainLabel = "(Local: $dnsRoot)"
}

Set-LocalDomainIdentity -dnsRoot (Resolve-LocalDomainDns)

$cbDomainSelect.Items.Add($script:LocalDomainLabel) | Out-Null
$cbDomainSelect.SelectedIndex = 0

function Get-ADParams {
  if (-not $global:activeDomain -or $global:activeDomain -eq $script:LocalDomainLabel) { return @{} }
  $conn = $global:domainConnections[$global:activeDomain]
  if (-not $conn) { return @{} }
  $p = @{ Server = $conn.Server }
  if ($conn.Credential) { $p.Credential = $conn.Credential }
  return $p
}

function Clear-ADDomainCache {
  $script:ADDomainCache = @{}
}

function Get-ADDomainCached([hashtable]$adParams = $null, [switch]$ForceRefresh) {
  if ($null -eq $adParams) { $adParams = Get-ADParams }

  $serverKey = '(local)'
  if ($adParams.ContainsKey('Server') -and (-not [string]::IsNullOrWhiteSpace([string]$adParams.Server))) {
    $serverKey = [string]$adParams.Server
  }

  $credKey = ''
  if ($adParams.ContainsKey('Credential') -and $adParams.Credential -and $adParams.Credential.UserName) {
    $credKey = [string]$adParams.Credential.UserName
  }

  $cacheKey = "$serverKey|$credKey"
  if ($ForceRefresh -or (-not $script:ADDomainCache.ContainsKey($cacheKey))) {
    $script:ADDomainCache[$cacheKey] = Get-ADDomain @adParams
  }

  return $script:ADDomainCache[$cacheKey]
}

function Get-DomainConnectionsPersistable {
  $items = @()
  foreach ($label in @($global:domainConnections.Keys | Sort-Object)) {
    if ([string]::IsNullOrWhiteSpace([string]$label)) { continue }
    $conn = $global:domainConnections[$label]
    if (-not $conn) { continue }

    $userName = ''
    if ($conn.Credential -and $conn.Credential.UserName) { $userName = [string]$conn.Credential.UserName }
    elseif ($conn.PersistedUserName) { $userName = [string]$conn.PersistedUserName }

    $items += [PSCustomObject]@{
      Label    = [string]$label
      Server   = [string]$conn.Server
      Domain   = [string]$conn.Domain
      UserName = $userName
    }
  }
  return @($items)
}

function Save-DomainConnectionsToFile {
  try {
    $payload = [PSCustomObject]@{
      LastActive = if (($global:activeDomain) -and ($global:activeDomain -ne $script:LocalDomainLabel)) { [string]$global:activeDomain } else { '' }
      Items      = @(Get-DomainConnectionsPersistable)
    }
    $json = $payload | ConvertTo-Json -Depth 6
    Set-Content -Path $script:DomainConnectionsConfigPath -Value $json -Encoding UTF8
  }
  catch {
    # No interrumpir la app por errores de persistencia.
  }
}

function Load-DomainConnectionsFromFile {
  if (-not (Test-Path -Path $script:DomainConnectionsConfigPath -PathType Leaf)) { return }

  try {
    $raw = Get-Content -Path $script:DomainConnectionsConfigPath -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) { return }

    $cfg = $raw | ConvertFrom-Json -ErrorAction Stop
    $items = @()
    if ($cfg -and $cfg.PSObject.Properties['Items']) {
      $items = @($cfg.Items)
    }

    $global:domainConnections = @{}
    foreach ($item in @($items)) {
      $label = [string]$item.Label
      $server = [string]$item.Server
      if ([string]::IsNullOrWhiteSpace($label) -or [string]::IsNullOrWhiteSpace($server)) { continue }

      $global:domainConnections[$label] = @{
        Server            = $server
        Credential        = $null
        Domain            = [string]$item.Domain
        PersistedUserName = [string]$item.UserName
      }
    }

    Refresh-DomainCombo
    $lastActive = ''
    if ($cfg -and $cfg.PSObject.Properties['LastActive']) { $lastActive = [string]$cfg.LastActive }
    if ((-not [string]::IsNullOrWhiteSpace($lastActive)) -and $cbDomainSelect.Items.Contains($lastActive)) {
      $cbDomainSelect.SelectedItem = $lastActive
    }
  }
  catch {
    # Si el archivo está corrupto, se ignora y se inicia sin conexiones persistidas.
  }
}

function Refresh-DomainCombo {
  $cbDomainSelect.Dispatcher.Invoke([Action] {
      $current = $cbDomainSelect.SelectedItem
      $cbDomainSelect.Items.Clear() | Out-Null
      $cbDomainSelect.Items.Add($script:LocalDomainLabel) | Out-Null

      foreach ($key in @($global:domainConnections.Keys)) {
        if (-not [string]::IsNullOrWhiteSpace($key)) {
          $cbDomainSelect.Items.Add($key) | Out-Null
        }
      }

      if ($null -ne $current) {
        $idx = $cbDomainSelect.Items.IndexOf($current)
        if ($idx -ge 0) { $cbDomainSelect.SelectedIndex = $idx; return }
      }
      $cbDomainSelect.SelectedIndex = 0
    })
}

function Set-SimulationMode([bool]$enabled) {
  $script:SimulationMode = $enabled
  $script:WhatIfPreference = $enabled

  if ($enabled) {
    Set-Status "⚠️ Modo simulación activo (WhatIf): no se aplican cambios reales." '#fab387'
  }
  else {
    Set-Status "✅ Modo simulación desactivado. Cambios reales habilitados." '#a6e3a1'
  }
}

$cbDomainSelect.Add_SelectionChanged({
    $selected = $cbDomainSelect.SelectedItem
    if ($null -eq $selected) { return }

    $global:activeDomain = $selected
    Clear-ADDomainCache
    if ($selected -eq $script:LocalDomainLabel) {
      $txtDomainInfo.Text = "Dominio activo: $($script:LocalDomainDns)"
    }
    else {
      $conn = $global:domainConnections[$selected]
      if ($conn) {
        $domainText = if (-not [string]::IsNullOrWhiteSpace([string]$conn.Domain)) { [string]$conn.Domain } else { [string]$selected }
        $txtDomainInfo.Text = "Dominio activo: $domainText | Servidor: $($conn.Server)"
      }
    }
  })

$txtDomainInfo.Text = "Dominio activo: $($script:LocalDomainDns)"
Load-DomainConnectionsFromFile

if ($chkSimulationMode) {
  $chkSimulationMode.IsChecked = $false
  $chkSimulationMode.ToolTip = "Si está activo, las operaciones de escritura en AD se ejecutan en modo simulación."
  $chkSimulationMode.Add_Checked({ Set-SimulationMode $true })
  $chkSimulationMode.Add_Unchecked({ Set-SimulationMode $false })
}

# ── TEMAS ────────────────────────────────────────────────────────────────────
$script:isDarkMode = $true
$script:ThemeConfigPath = Join-Path $PSScriptRoot 'ui.theme.json'
$script:StatusColor = '#a6adc8'
$script:BrushCache = @{}
$script:themes = @{
  Dark  = @{
    WindowBg = '#111827'; SurfaceBg = '#161f30'; CardBg = '#202b3f'; PanelBg = '#182235'; SurfaceElevated = '#24324b'
    TextPrimary = '#e6edf9'; TextSecondary = '#93a4c3'; TextMuted = '#b5c0d5'
    Accent = '#7cc7ff'; Border = '#31425e'; PanelBorder = '#405473'
    RowBg = '#182235'; RowAltBg = '#1d2940'; GridLine = '#2e415f'
    GridHeaderBg = '#24324b'; GridHeaderFg = '#8fd2ff'
    TreeItemFg = '#8fd2ff'
    BtnPrimaryBg = '#7cc7ff'; BtnPrimaryFg = '#0f1c2e'; BtnPrimaryHoverBg = '#9ad4ff'; BtnPrimaryPressedBg = '#63b9f7'; BtnPrimaryBorder = '#8bd0ff'
    BtnSuccessBg = '#66d3a2'; BtnSuccessFg = '#10261e'; BtnSuccessHoverBg = '#7ee0b3'; BtnSuccessPressedBg = '#51bf90'; BtnSuccessBorder = '#8de5be'
    BtnDangerBg = '#ff8ba7'; BtnDangerFg = '#30101a'; BtnDangerHoverBg = '#ff9fb6'; BtnDangerPressedBg = '#f47795'; BtnDangerBorder = '#ffb4c5'
    BtnWarnBg = '#ffbb70'; BtnWarnFg = '#34220d'; BtnWarnHoverBg = '#ffc88b'; BtnWarnPressedBg = '#f2aa5c'; BtnWarnBorder = '#ffd39d'
    InputBg = '#202b3f'; InputFg = '#e6edf9'; ChipBg = '#21314a'; ChipFg = '#dce9ff'
    HeaderBg = '#0f1726'; StatusBg = '#0f1726'; SelectionBg = '#315783'; SelectionFg = '#f8fbff'
  }
  Light = @{
    WindowBg = '#f4f7fb'; SurfaceBg = '#edf2f8'; CardBg = '#ffffff'; PanelBg = '#ffffff'; SurfaceElevated = '#e7eef8'
    TextPrimary = '#1f2a3d'; TextSecondary = '#5a6f8f'; TextMuted = '#667892'
    Accent = '#1f6fe5'; Border = '#c6d3e5'; PanelBorder = '#d5dfed'
    RowBg = '#ffffff'; RowAltBg = '#f4f8fd'; GridLine = '#d8e1ee'
    GridHeaderBg = '#e7eef8'; GridHeaderFg = '#2654a8'
    TreeItemFg = '#2654a8'
    BtnPrimaryBg = '#2f7cf6'; BtnPrimaryFg = '#f8fbff'; BtnPrimaryHoverBg = '#206fe7'; BtnPrimaryPressedBg = '#165ec8'; BtnPrimaryBorder = '#4c90fb'
    BtnSuccessBg = '#2ea46f'; BtnSuccessFg = '#f8fbff'; BtnSuccessHoverBg = '#259664'; BtnSuccessPressedBg = '#1f8054'; BtnSuccessBorder = '#4cb784'
    BtnDangerBg = '#d95f7c'; BtnDangerFg = '#f8fbff'; BtnDangerHoverBg = '#ca4f6e'; BtnDangerPressedBg = '#b54260'; BtnDangerBorder = '#e18498'
    BtnWarnBg = '#d68d38'; BtnWarnFg = '#fffaf4'; BtnWarnHoverBg = '#c9802d'; BtnWarnPressedBg = '#b56d1f'; BtnWarnBorder = '#e2a45d'
    InputBg = '#fdfefe'; InputFg = '#243248'; ChipBg = '#edf4ff'; ChipFg = '#23478f'
    HeaderBg = '#ffffff'; StatusBg = '#ffffff'; SelectionBg = '#d7e8ff'; SelectionFg = '#17325f'
  }
}

function ConvertTo-Brush([string]$hex) {
  if ([string]::IsNullOrWhiteSpace($hex)) { return $null }
  $key = $hex.Trim().ToLowerInvariant()
  if ($script:BrushCache.ContainsKey($key)) { return $script:BrushCache[$key] }

  $brush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($hex)
  if (($brush -is [System.Windows.Freezable]) -and $brush.CanFreeze) {
    $brush.Freeze()
  }
  $script:BrushCache[$key] = $brush
  return $brush
}
function Get-ActiveTheme { if ($script:isDarkMode) { return $script:themes.Dark } return $script:themes.Light }

function Save-ThemePreference {
  try {
    $payload = [PSCustomObject]@{
      IsDarkMode = [bool]$script:isDarkMode
    }
    $json = $payload | ConvertTo-Json -Depth 4
    Set-Content -Path $script:ThemeConfigPath -Value $json -Encoding UTF8
  }
  catch {
    # No bloquear la UI por fallas al guardar tema.
  }
}

function Load-ThemePreference {
  if (-not (Test-Path -Path $script:ThemeConfigPath -PathType Leaf)) { return }
  try {
    $raw = Get-Content -Path $script:ThemeConfigPath -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) { return }
    $cfg = $raw | ConvertFrom-Json -ErrorAction Stop
    if ($cfg -and $cfg.PSObject.Properties['IsDarkMode']) {
      if ($cfg.IsDarkMode -is [bool]) { $script:isDarkMode = [bool]$cfg.IsDarkMode }
    }
  }
  catch {
    # Si está corrupto, mantener valor por defecto.
  }
}

function Get-VisualDescendants([System.Windows.DependencyObject]$root, [Type]$type) {
  if (-not $root) { return @() }

  $found = New-Object 'System.Collections.Generic.List[object]'
  $stack = New-Object 'System.Collections.Generic.Stack[System.Windows.DependencyObject]'
  $stack.Push($root)

  while ($stack.Count -gt 0) {
    $current = $stack.Pop()

    # VisualTreeHelper solo acepta Visual/Visual3D.
    if ((-not ($current -is [System.Windows.Media.Visual])) -and (-not ($current -is [System.Windows.Media.Media3D.Visual3D]))) { continue }

    $childCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($current)
    for ($i = 0; $i -lt $childCount; $i++) {
      $child = [System.Windows.Media.VisualTreeHelper]::GetChild($current, $i)
      if ($null -eq $child) { continue }

      if ($type.IsAssignableFrom($child.GetType())) {
        [void]$found.Add($child)
      }

      if (($child -is [System.Windows.Media.Visual]) -or ($child -is [System.Windows.Media.Media3D.Visual3D])) {
        $stack.Push($child)
      }
    }
  }

  return @($found.ToArray())
}

function Set-ThemeTabResources($t) {
  $window.Resources['ThemeWindowBg'] = ConvertTo-Brush $t.WindowBg
  $window.Resources['ThemeTabBg'] = ConvertTo-Brush $t.WindowBg
  $window.Resources['ThemeTabHoverBg'] = ConvertTo-Brush $t.SurfaceElevated
  $window.Resources['ThemeTabSelectedBg'] = ConvertTo-Brush $t.PanelBg
  $window.Resources['ThemeTabFg'] = ConvertTo-Brush $t.TextSecondary
  $window.Resources['ThemeTabSelectedFg'] = ConvertTo-Brush $t.Accent
  $window.Resources['ThemeAccent'] = ConvertTo-Brush $t.Accent
  $window.Resources['ThemeBorder'] = ConvertTo-Brush $t.Border
  $window.Resources['ThemePanelBg'] = ConvertTo-Brush $t.PanelBg
  $window.Resources['ThemePanelBorder'] = ConvertTo-Brush $t.PanelBorder
  $window.Resources['ThemeInputBg'] = ConvertTo-Brush $t.InputBg
  $window.Resources['ThemeInputFg'] = ConvertTo-Brush $t.InputFg
  $window.Resources['ThemeSurfaceBg'] = ConvertTo-Brush $t.SurfaceBg
  $window.Resources['ThemeTextPrimary'] = ConvertTo-Brush $t.TextPrimary
  $window.Resources['ThemeLabelFg'] = ConvertTo-Brush $t.TextSecondary
  $window.Resources['ThemeGroupBoxFg'] = ConvertTo-Brush $t.Accent
  $window.Resources['ThemeTreeBg'] = ConvertTo-Brush $t.SurfaceBg
  $window.Resources['ThemeTreeFg'] = ConvertTo-Brush $t.TextPrimary
  $window.Resources['ThemeTreeBorder'] = ConvertTo-Brush $t.Border
  $window.Resources['ThemeTreeItemFg'] = ConvertTo-Brush $t.TreeItemFg
  $window.Resources['ThemeChipBg'] = ConvertTo-Brush $t.ChipBg
  $window.Resources['ThemeChipFg'] = ConvertTo-Brush $t.ChipFg
  $window.Resources['ThemeSelectionBg'] = ConvertTo-Brush $t.SelectionBg
  $window.Resources['ThemeSelectionFg'] = ConvertTo-Brush $t.SelectionFg
}

function Set-ThemeButtonResources($t) {
  $window.Resources['ThemeBtnPrimaryBg'] = ConvertTo-Brush $t.BtnPrimaryBg
  $window.Resources['ThemeBtnPrimaryFg'] = ConvertTo-Brush $t.BtnPrimaryFg
  $window.Resources['ThemeBtnPrimaryHoverBg'] = ConvertTo-Brush $t.BtnPrimaryHoverBg
  $window.Resources['ThemeBtnPrimaryPressedBg'] = ConvertTo-Brush $t.BtnPrimaryPressedBg
  $window.Resources['ThemeBtnPrimaryBorder'] = ConvertTo-Brush $t.BtnPrimaryBorder

  $window.Resources['ThemeBtnSuccessBg'] = ConvertTo-Brush $t.BtnSuccessBg
  $window.Resources['ThemeBtnSuccessFg'] = ConvertTo-Brush $t.BtnSuccessFg
  $window.Resources['ThemeBtnSuccessHoverBg'] = ConvertTo-Brush $t.BtnSuccessHoverBg
  $window.Resources['ThemeBtnSuccessPressedBg'] = ConvertTo-Brush $t.BtnSuccessPressedBg
  $window.Resources['ThemeBtnSuccessBorder'] = ConvertTo-Brush $t.BtnSuccessBorder

  $window.Resources['ThemeBtnDangerBg'] = ConvertTo-Brush $t.BtnDangerBg
  $window.Resources['ThemeBtnDangerFg'] = ConvertTo-Brush $t.BtnDangerFg
  $window.Resources['ThemeBtnDangerHoverBg'] = ConvertTo-Brush $t.BtnDangerHoverBg
  $window.Resources['ThemeBtnDangerPressedBg'] = ConvertTo-Brush $t.BtnDangerPressedBg
  $window.Resources['ThemeBtnDangerBorder'] = ConvertTo-Brush $t.BtnDangerBorder

  $window.Resources['ThemeBtnWarnBg'] = ConvertTo-Brush $t.BtnWarnBg
  $window.Resources['ThemeBtnWarnFg'] = ConvertTo-Brush $t.BtnWarnFg
  $window.Resources['ThemeBtnWarnHoverBg'] = ConvertTo-Brush $t.BtnWarnHoverBg
  $window.Resources['ThemeBtnWarnPressedBg'] = ConvertTo-Brush $t.BtnWarnPressedBg
  $window.Resources['ThemeBtnWarnBorder'] = ConvertTo-Brush $t.BtnWarnBorder
}

function Set-ThemeGridResources($t) {
  $window.Resources['ThemeGridBg'] = ConvertTo-Brush $t.SurfaceBg
  $window.Resources['ThemeGridFg'] = ConvertTo-Brush $t.TextPrimary
  $window.Resources['ThemeGridBorder'] = ConvertTo-Brush $t.Border
  $window.Resources['ThemeGridRowBg'] = ConvertTo-Brush $t.RowBg
  $window.Resources['ThemeGridRowAltBg'] = ConvertTo-Brush $t.RowAltBg
  $window.Resources['ThemeGridLine'] = ConvertTo-Brush $t.GridLine
  $window.Resources['ThemeGridHeaderBg'] = ConvertTo-Brush $t.GridHeaderBg
  $window.Resources['ThemeGridHeaderFg'] = ConvertTo-Brush $t.GridHeaderFg
}

function Apply-ThemeToInputs($t) {
  foreach ($tb in @($txtUserSearch, $txtGroupSearch, $txtAttrSearch, $txtBatchCsvPath)) {
    $tb.Background = ConvertTo-Brush $t.InputBg
    $tb.Foreground = ConvertTo-Brush $t.InputFg
    $tb.BorderBrush = ConvertTo-Brush $t.Border
  }

  if ($cbDomainSelect) {
    $cbDomainSelect.Background = ConvertTo-Brush $t.InputBg
    $cbDomainSelect.Foreground = ConvertTo-Brush $t.InputFg
    $cbDomainSelect.BorderBrush = ConvertTo-Brush $t.Border
  }

  if ($cbAuditMode) {
    $cbAuditMode.Background = ConvertTo-Brush $t.InputBg
    $cbAuditMode.Foreground = ConvertTo-Brush $t.InputFg
    $cbAuditMode.BorderBrush = ConvertTo-Brush $t.Border
  }
}

function Apply-ThemeToGrids($t) {
  foreach ($dg in @($dgUsers, $dgGroups, $dgMoveObjects, $dgMoveTargetObjects, $dgBatchResults, $dgAttributes)) {
    $dg.Background = ConvertTo-Brush $t.SurfaceBg
    $dg.Foreground = ConvertTo-Brush $t.TextPrimary
    $dg.BorderBrush = ConvertTo-Brush $t.Border
    $dg.RowBackground = ConvertTo-Brush $t.RowBg
    $dg.AlternatingRowBackground = ConvertTo-Brush $t.RowAltBg
    $dg.HorizontalGridLinesBrush = ConvertTo-Brush $t.GridLine
  }
}

function Apply-ThemeToTrees($t) {
  foreach ($tv in @($tvMoveSourceOUs, $tvMoveTargetOUs)) {
    $tv.Background = ConvertTo-Brush $t.SurfaceBg
    $tv.Foreground = ConvertTo-Brush $t.TextPrimary
    $tv.BorderBrush = ConvertTo-Brush $t.Border
  }
  foreach ($item in (Get-VisualDescendants -root $window -type ([System.Windows.Controls.TreeViewItem]))) {
    $item.Foreground = ConvertTo-Brush $t.TreeItemFg
  }
}

function Apply-ThemeToGridHeaders($t) {
  foreach ($hdr in (Get-VisualDescendants -root $window -type ([System.Windows.Controls.Primitives.DataGridColumnHeader]))) {
    $hdr.Background = ConvertTo-Brush $t.GridHeaderBg
    $hdr.Foreground = ConvertTo-Brush $t.GridHeaderFg
    $hdr.BorderBrush = ConvertTo-Brush $t.Border
  }
}

function Apply-ThemeToDataGridVisuals($t) {
  $bg = ConvertTo-Brush $t.SurfaceBg
  $rowBg = ConvertTo-Brush $t.RowBg
  $rowAltBg = ConvertTo-Brush $t.RowAltBg
  $fg = ConvertTo-Brush $t.TextPrimary

  foreach ($dg in @($dgUsers, $dgGroups, $dgMoveObjects, $dgMoveTargetObjects, $dgBatchResults, $dgAttributes)) {
    foreach ($sv in (Get-VisualDescendants -root $dg -type ([System.Windows.Controls.ScrollViewer]))) {
      $sv.Background = $bg
    }
    foreach ($row in (Get-VisualDescendants -root $dg -type ([System.Windows.Controls.DataGridRow]))) {
      if (-not $row.IsSelected) {
        if (($row.GetIndex() % 2) -eq 0) { $row.Background = $rowBg } else { $row.Background = $rowAltBg }
      }
      $row.Foreground = $fg
    }
    foreach ($cell in (Get-VisualDescendants -root $dg -type ([System.Windows.Controls.DataGridCell]))) {
      if (-not $cell.IsSelected) {
        $cell.Foreground = $fg
        $cell.BorderBrush = ConvertTo-Brush $t.Border
      }
    }
  }
}

function Apply-ThemeToLabelsAndChecks($t) {
  foreach ($lbl in (Get-VisualDescendants -root $window -type ([System.Windows.Controls.Label]))) {
    $lbl.Foreground = ConvertTo-Brush $t.TextSecondary
  }
  foreach ($cb in (Get-VisualDescendants -root $window -type ([System.Windows.Controls.CheckBox]))) {
    $cb.Foreground = ConvertTo-Brush $t.TextPrimary
  }
}

function Apply-ThemeToLegacyBorders($t) {
  # Mapeo bidireccional: evita estados aleatorios al alternar varias veces claro/oscuro.
  $legacyRoleByHex = @{
    '#ff181825' = 'SurfaceBg'
    '#ffe6e9ef' = 'SurfaceBg'
    '#ff232336' = 'SurfaceBg'
    '#ff1e1e2e' = 'WindowBg'
    '#ffeff1f5' = 'WindowBg'
    '#ff313244' = 'CardBg'
    '#ffdce0e8' = 'CardBg'
    '#ff292940' = 'CardBg'
  }
  foreach ($bd in (Get-VisualDescendants -root $window -type ([System.Windows.Controls.Border]))) {
    if ($bd -eq $borderHeader -or $bd -eq $borderDomainBar -or $bd -eq $borderStatus) { continue }
    if ($bd.Background -is [System.Windows.Media.SolidColorBrush]) {
      $hex = $bd.Background.Color.ToString().ToLowerInvariant()
      if ($legacyRoleByHex.ContainsKey($hex)) {
        $role = $legacyRoleByHex[$hex]
        $bd.Background = ConvertTo-Brush $t[$role]
      }
    }
    if ($bd.BorderBrush -is [System.Windows.Media.SolidColorBrush]) {
      $bd.BorderBrush = ConvertTo-Brush $t.Border
    }
  }
}

function Apply-ThemeToGroupBoxes($t) {
  foreach ($gb in (Get-VisualDescendants -root $window -type ([System.Windows.Controls.GroupBox]))) {
    $gb.BorderBrush = ConvertTo-Brush $t.PanelBorder
    $gb.Foreground = ConvertTo-Brush $t.Accent
  }
}

function Apply-ThemeToPanels($t) {
  foreach ($bd in @(
      $borderUserSearchPanel, $borderUserActionsPanel,
      $borderGroupSearchPanel, $borderGroupActionsPanel,
      $borderAttrSearchPanel, $borderAttrActionsPanel,
      $borderConfigHeaderPanel, $borderConfigBodyPanel, $borderConfigActionsPanel
    )) {
    if ($bd) {
      $bd.Background = ConvertTo-Brush $t.PanelBg
      $bd.BorderBrush = ConvertTo-Brush $t.PanelBorder
      $bd.BorderThickness = '1'
    }
  }
}

function Apply-Theme {
  $t = if ($script:isDarkMode) { $script:themes.Dark } else { $script:themes.Light }
  $window.Background = ConvertTo-Brush $t.WindowBg
  $window.Foreground = ConvertTo-Brush $t.TextPrimary
  $borderHeader.Background = ConvertTo-Brush $t.HeaderBg
  $borderHeader.BorderBrush = ConvertTo-Brush $t.PanelBorder
  $borderDomainBar.Background = ConvertTo-Brush $t.PanelBg
  $borderDomainBar.BorderBrush = ConvertTo-Brush $t.PanelBorder
  $borderStatus.Background = ConvertTo-Brush $t.StatusBg
  $borderStatus.BorderBrush = ConvertTo-Brush $t.PanelBorder
  $tabControl.Background = ConvertTo-Brush $t.WindowBg
  $tabControl.BorderBrush = ConvertTo-Brush $t.PanelBorder
  $txtTitle.Foreground = ConvertTo-Brush $t.Accent
  $txtSubtitle.Foreground = ConvertTo-Brush $t.TextSecondary
  $txtDomainLabel.Foreground = ConvertTo-Brush $t.TextSecondary
  $txtDomainInfo.Foreground = ConvertTo-Brush $t.TextSecondary

  Set-ThemeTabResources $t
  Set-ThemeButtonResources $t
  Set-ThemeGridResources $t
  Apply-ThemeToInputs $t
  Apply-ThemeToGrids $t
  Apply-ThemeToGridHeaders $t
  Apply-ThemeToDataGridVisuals $t
  Apply-ThemeToTrees $t
  Apply-ThemeToLabelsAndChecks $t
  Apply-ThemeToLegacyBorders $t
  Apply-ThemeToGroupBoxes $t
  Apply-ThemeToPanels $t

  if ([string]::IsNullOrWhiteSpace([string]$script:StatusColor)) {
    $script:StatusColor = $t.TextMuted
  }
  $txtStatus.Foreground = ConvertTo-Brush $script:StatusColor
  $txtAttrDN.Foreground = ConvertTo-Brush $t.TextSecondary

  if ($script:isDarkMode) { $btnThemeToggle.Content = "☀️ Claro" } else { $btnThemeToggle.Content = "🌙 Oscuro" }
}

Load-ThemePreference
$btnThemeToggle.Add_Click({
    $script:isDarkMode = -not $script:isDarkMode
    Apply-Theme
    Save-ThemePreference
  })
$window.Add_ContentRendered({ Apply-Theme })
$tabControl.Add_SelectionChanged({
    if ($_.Source -eq $tabControl) { Apply-Theme }
  })

# ── AUX ──────────────────────────────────────────────────────────────────────
function Set-Status([string]$msg, [string]$color = '#a6adc8') {
  $txtStatus.Text = $msg
  $script:StatusColor = $color
  $txtStatus.Foreground = ConvertTo-Brush $color
}

$script:AuditLogDir = Join-Path $PSScriptRoot 'AuditLogs'
$script:AuditLogFilePrefix = 'AD_Audit_'
$script:AuditLocalRetentionMonths = 2
$script:AuditSettings = @{
  EnableCentral      = $false
  CentralMode        = 'FileShare' # FileShare | SqlServer
  CentralFilePath    = '\\SERVIDOR\AD-Audit\AD_Audit_Central.csv'
  SqlConnectionString = '' # Ej: Server=SQL01;Database=AdminAD;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;
  SqlTable           = 'dbo.ADAuditLog'
}
$script:AuditConfigLocalPath = Join-Path $PSScriptRoot 'audit.settings.json'
$script:AuditConfigSharedPath = [string]$env:AD_MANAGER_AUDIT_CONFIG_PATH
$script:AuditConfigPath = $script:AuditConfigLocalPath
$script:AuditLastCentralError = ''
$script:AuditLastCentralOk = ''
$script:AuditUiSyncInProgress = $false
$script:UserPolicyConfigEnvPath = [string]$env:AD_MANAGER_USER_POLICY_PATH
$script:UserPolicyConfigLocalPath = Join-Path $PSScriptRoot 'user.form.policy.json'
$script:UserPolicyConfigProgramDataPath = if (-not [string]::IsNullOrWhiteSpace([string]$env:ProgramData)) { Join-Path $env:ProgramData 'ADManager\user.form.policy.json' } else { '' }
$script:UserPolicyConfigAppDataPath = if (-not [string]::IsNullOrWhiteSpace([string]$env:LOCALAPPDATA)) { Join-Path $env:LOCALAPPDATA 'ADManager\user.form.policy.json' } else { '' }
$script:UserPolicyLegacyPaths = @(
  (Join-Path ([Environment]::GetFolderPath('Desktop')) 'user.form.policy.json'),
  (Join-Path ([Environment]::GetFolderPath('CommonDesktopDirectory')) 'user.form.policy.json')
)
$script:UserPolicyConfigPath = $script:UserPolicyConfigLocalPath
$script:UserPolicy = $null
$script:UserPolicyUnlocked = $false

function ConvertTo-BoolValue([object]$value, [bool]$defaultValue = $false) {
  if ($null -eq $value) { return $defaultValue }
  if ($value -is [bool]) { return [bool]$value }
  $text = ([string]$value).Trim().ToLowerInvariant()
  switch -Regex ($text) {
    '^(1|true|t|si|sí|yes|y|on)$' { return $true }
    '^(0|false|f|no|n|off)$' { return $false }
    default { return $defaultValue }
  }
}

function Set-AuditSettingsFromGui {
  if (-not $chkAuditEnableCentral) { return }

  $mode = [string]$cbAuditMode.SelectedItem
  if ([string]::IsNullOrWhiteSpace($mode)) { $mode = 'FileShare' }
  if (($mode -ne 'FileShare') -and ($mode -ne 'SqlServer')) { $mode = 'FileShare' }

  $script:AuditSettings.EnableCentral = (Get-Checked $chkAuditEnableCentral)
  $script:AuditSettings.CentralMode = $mode
  $rawCentralPath = [string]$txtAuditCentralFilePath.Text.Trim()
  if ($mode -eq 'FileShare') {
    $resolvedCentralPath = Resolve-AuditCentralFilePath -inputPath $rawCentralPath -CreateIfMissing
    $script:AuditSettings.CentralFilePath = $resolvedCentralPath
    if ($txtAuditCentralFilePath) { $txtAuditCentralFilePath.Text = $resolvedCentralPath }
  }
  else {
    $script:AuditSettings.CentralFilePath = $rawCentralPath
  }
  $script:AuditSettings.SqlConnectionString = [string]$txtAuditSqlConnection.Text.Trim()
  $script:AuditSettings.SqlTable = [string]$txtAuditSqlTable.Text.Trim()
}

function Resolve-AuditCentralFilePath([string]$inputPath, [switch]$CreateIfMissing) {
  $path = [string]$inputPath
  if ([string]::IsNullOrWhiteSpace($path)) { return '' }
  $path = $path.Trim()

  # Si no termina en .csv, se interpreta como carpeta.
  if ($path -notmatch '\.csv$') {
    $folder = $path
    $selectedCsv = $null

    if (Test-Path -Path $folder -PathType Container) {
      $preferred = Join-Path $folder 'AD_Audit_Central.csv'
      if (Test-Path -Path $preferred -PathType Leaf) {
        $selectedCsv = $preferred
      }
      else {
        $anyCsv = Get-ChildItem -Path $folder -Filter '*.csv' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
        if ($anyCsv) { $selectedCsv = $anyCsv.FullName }
      }
    }

    if ($null -eq $selectedCsv) {
      $selectedCsv = Join-Path $folder 'AD_Audit_Central.csv'
    }
    $path = $selectedCsv
  }

  if ($CreateIfMissing) {
    Ensure-CentralFileLog -path $path
  }

  return $path
}

function Update-AuditGuiFromSettings {
  if (-not $chkAuditEnableCentral) { return }

  $script:AuditUiSyncInProgress = $true
  try {
    $chkAuditEnableCentral.IsChecked = ConvertTo-BoolValue $script:AuditSettings.EnableCentral $false

    if ($cbAuditMode.Items.Count -eq 0) {
      $cbAuditMode.Items.Add('FileShare') | Out-Null
      $cbAuditMode.Items.Add('SqlServer') | Out-Null
    }

    $mode = [string]$script:AuditSettings.CentralMode
    if (($mode -ne 'FileShare') -and ($mode -ne 'SqlServer')) { $mode = 'FileShare' }
    $cbAuditMode.SelectedItem = $mode

    if (($mode -eq 'FileShare') -and [string]::IsNullOrWhiteSpace([string]$script:AuditSettings.CentralFilePath)) {
      $script:AuditSettings.CentralFilePath = '\\SERVIDOR\AD-Audit\AD_Audit_Central.csv'
    }

    $txtAuditCentralFilePath.Text = [string]$script:AuditSettings.CentralFilePath
    $txtAuditSqlConnection.Text = [string]$script:AuditSettings.SqlConnectionString
    $txtAuditSqlTable.Text = [string]$script:AuditSettings.SqlTable
  }
  finally {
    $script:AuditUiSyncInProgress = $false
  }
}

function Get-AuditConfigReadPaths {
  $paths = @()
  if (-not [string]::IsNullOrWhiteSpace($script:AuditConfigSharedPath)) {
    $paths += $script:AuditConfigSharedPath

    # Compatibilidad: admitir typo histórico "settines" en la ruta compartida.
    if ($script:AuditConfigSharedPath -match 'settines') {
      $paths += ($script:AuditConfigSharedPath -replace 'settines', 'settings')
    }
  }
  if (-not [string]::IsNullOrWhiteSpace($script:AuditConfigLocalPath)) { $paths += $script:AuditConfigLocalPath }
  return @($paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Get-AuditConfigWritePath {
  if (-not [string]::IsNullOrWhiteSpace($script:AuditConfigSharedPath)) { return $script:AuditConfigSharedPath }
  return $script:AuditConfigLocalPath
}

function Get-AuditConfigCandidatesText {
  $paths = Get-AuditConfigReadPaths
  if ($null -eq $paths -or $paths.Count -eq 0) { return '(sin rutas configuradas)' }
  return ($paths -join ' | ')
}

function Write-AuditSettingsFile([string]$path, [string]$json) {
  if ([string]::IsNullOrWhiteSpace($path)) { throw "Ruta de configuración inválida." }
  $parent = Split-Path -Path $path -Parent
  if ((-not [string]::IsNullOrWhiteSpace($parent)) -and (-not (Test-Path -Path $parent -PathType Container))) {
    New-Item -Path $parent -ItemType Directory -Force | Out-Null
  }
  Set-Content -Path $path -Value $json -Encoding UTF8
}

function Save-AuditSettingsToFile {
  $out = [PSCustomObject]@{
    EnableCentral       = [bool]$script:AuditSettings.EnableCentral
    CentralMode         = [string]$script:AuditSettings.CentralMode
    CentralFilePath     = [string]$script:AuditSettings.CentralFilePath
    SqlConnectionString = [string]$script:AuditSettings.SqlConnectionString
    SqlTable            = [string]$script:AuditSettings.SqlTable
  }
  $json = $out | ConvertTo-Json -Depth 4
  $targetPath = Get-AuditConfigWritePath
  try {
    Write-AuditSettingsFile -path $targetPath -json $json
    $script:AuditConfigPath = $targetPath
  }
  catch {
    if ($targetPath -ne $script:AuditConfigLocalPath) {
      Write-AuditSettingsFile -path $script:AuditConfigLocalPath -json $json
      $script:AuditConfigPath = $script:AuditConfigLocalPath
      return
    }
    throw
  }
}

function Load-AuditSettingsFromFile {
  $existingPaths = @()
  foreach ($candidate in (Get-AuditConfigReadPaths)) {
    if (Test-Path -Path $candidate -PathType Leaf) { $existingPaths += $candidate }
  }
  if ($existingPaths.Count -eq 0) { return $false }

  $lastErrorMessage = ''
  foreach ($path in $existingPaths) {
    try {
      $raw = Get-Content -Path $path -Raw -ErrorAction Stop
      if ([string]::IsNullOrWhiteSpace($raw)) { continue }
      $cfg = $raw | ConvertFrom-Json -ErrorAction Stop
      if (-not $cfg) { continue }

      $script:AuditSettings.EnableCentral = ConvertTo-BoolValue $cfg.EnableCentral $false
      $script:AuditSettings.CentralMode = [string]$cfg.CentralMode
      if (($script:AuditSettings.CentralMode -ne 'FileShare') -and ($script:AuditSettings.CentralMode -ne 'SqlServer')) {
        $script:AuditSettings.CentralMode = 'FileShare'
      }

      # Compatibilidad: aceptar nombres de clave legacy y no pisar con vacío.
      $loadedCentralPath = ''
      if ($cfg.PSObject.Properties['CentralFilePath']) { $loadedCentralPath = [string]$cfg.CentralFilePath }
      if ([string]::IsNullOrWhiteSpace($loadedCentralPath) -and $cfg.PSObject.Properties['AuditCentralFilePath']) { $loadedCentralPath = [string]$cfg.AuditCentralFilePath }
      if ([string]::IsNullOrWhiteSpace($loadedCentralPath) -and $cfg.PSObject.Properties['FileSharePath']) { $loadedCentralPath = [string]$cfg.FileSharePath }
      if ([string]::IsNullOrWhiteSpace($loadedCentralPath) -and $cfg.PSObject.Properties['CentralPath']) { $loadedCentralPath = [string]$cfg.CentralPath }
      if (-not [string]::IsNullOrWhiteSpace($loadedCentralPath)) {
        $script:AuditSettings.CentralFilePath = Resolve-AuditCentralFilePath -inputPath $loadedCentralPath
      }

      $script:AuditSettings.SqlConnectionString = [string]$cfg.SqlConnectionString
      $script:AuditSettings.SqlTable = [string]$cfg.SqlTable
      if ([string]::IsNullOrWhiteSpace($script:AuditSettings.SqlTable)) {
        $script:AuditSettings.SqlTable = 'dbo.ADAuditLog'
      }

      $script:AuditConfigPath = $path
      return $true
    }
    catch {
      $lastErrorMessage = $_.Exception.Message
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($lastErrorMessage)) {
    throw "No se pudo cargar la configuración de auditoría desde ninguna ruta disponible. Último error: $lastErrorMessage"
  }
  return $false
}

function Show-AuditDiagnostics([string[]]$lines) {
  if (-not $txtAuditDiagnostics) { return }
  if ($null -eq $lines -or $lines.Count -eq 0) {
    $txtAuditDiagnostics.Text = ''
    return
  }
  $txtAuditDiagnostics.Text = ($lines -join [Environment]::NewLine)
  $txtAuditDiagnostics.ScrollToHome()
}

if ($btnAuditDiagScrollDown) {
  $btnAuditDiagScrollDown.Add_Click({
      if ($txtAuditDiagnostics) { $txtAuditDiagnostics.ScrollToEnd() }
    })
}

if ($btnAuditBrowseCentralFile) {
  $btnAuditBrowseCentralFile.Add_Click({
      try {
        $sfd = New-Object Microsoft.Win32.SaveFileDialog
        $sfd.Filter = "CSV (*.csv)|*.csv|Todos (*.*)|*.*"
        $sfd.AddExtension = $true
        $sfd.DefaultExt = "csv"
        $sfd.FileName = "AD_Audit_Central.csv"

        $current = [string]$txtAuditCentralFilePath.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($current)) {
          try {
            $currentName = Split-Path -Path $current -Leaf
            $currentDir = Split-Path -Path $current -Parent
            if (-not [string]::IsNullOrWhiteSpace($currentName)) { $sfd.FileName = $currentName }
            if ((-not [string]::IsNullOrWhiteSpace($currentDir)) -and (Test-Path -Path $currentDir -PathType Container)) {
              $sfd.InitialDirectory = $currentDir
            }
          }
          catch {}
        }

        if ($sfd.ShowDialog() -eq $true) {
          $txtAuditCentralFilePath.Text = (Resolve-AuditCentralFilePath -inputPath $sfd.FileName)
          Set-Status "✅ Ruta central seleccionada." '#a6e3a1'
        }
      }
      catch { Show-Exception "Error seleccionando ruta central de auditoría:" $_ }
    })
}

function Get-AuditValidationResult([switch]$DeepTest) {
  $lines = @()
  $errors = @()
  $warnings = @()

  $enabled = ConvertTo-BoolValue $script:AuditSettings.EnableCentral $false
  $mode = [string]$script:AuditSettings.CentralMode
  if ([string]::IsNullOrWhiteSpace($mode)) { $mode = 'FileShare' }
  if (($mode -ne 'FileShare') -and ($mode -ne 'SqlServer')) {
    $errors += "Modo inválido: '$mode'. Debe ser FileShare o SqlServer."
  }

  $lines += "Estado general:"
  $lines += " - Auditoria central habilitada: $enabled"
  $lines += " - Modo: $mode"
  $lines += " - Config audit cargada desde: $($script:AuditConfigPath)"
  $lines += " - Rutas de config candidatas: $(Get-AuditConfigCandidatesText)"

  if (-not $enabled) {
    $warnings += "La auditoría central está deshabilitada. Solo se guardará log local."
  }
  elseif ($mode -eq 'FileShare') {
    $path = [string]$script:AuditSettings.CentralFilePath
    $lines += " - Destino FileShare: $path"
    if ([string]::IsNullOrWhiteSpace($path)) {
      $errors += "Ruta central vacía."
    }
    else {
      if ($path -notmatch '^\\\\') {
        $warnings += "La ruta no parece UNC (\\\\servidor\\recurso\\archivo.csv)."
      }
      if ($path -notmatch '\.csv$') {
        $warnings += "La ruta no termina en .csv."
      }

      $parent = Split-Path -Path $path -Parent
      if ([string]::IsNullOrWhiteSpace($parent)) {
        $errors += "No se pudo determinar carpeta padre de la ruta central."
      }
      else {
        try {
          if (Test-Path -Path $parent -PathType Container) {
            $lines += " - Carpeta detectada: OK ($parent)"
          }
          else {
            $warnings += "La carpeta no existe o no es accesible: $parent"
          }
        }
        catch {
          $warnings += "No se pudo validar acceso a carpeta: $($_.Exception.Message)"
        }
      }

      if ($DeepTest -and ($errors.Count -eq 0)) {
        try {
          Ensure-CentralFileLog -path $path
          $probe = [PSCustomObject]@{
            FechaHora = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Operador  = 'Diag'
            Dominio   = 'Diag'
            Accion    = 'DiagWrite'
            Resultado = 'TEST'
            Objeto    = 'FileShare'
            Detalle   = 'Diagnostico'
            Equipo    = $env:COMPUTERNAME
          }
          $line = $probe | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1
          Add-Content -Path $path -Value $line -Encoding UTF8
          $lines += " - Escritura de prueba: OK"
        }
        catch {
          $errors += "Error de escritura en FileShare: $($_.Exception.Message)"
        }
      }
    }
  }
  elseif ($mode -eq 'SqlServer') {
    $connStr = [string]$script:AuditSettings.SqlConnectionString
    $table = [string]$script:AuditSettings.SqlTable
    $lines += " - Tabla SQL: $table"

    if ([string]::IsNullOrWhiteSpace($connStr)) {
      $errors += "Connection string vacío."
    }
    else {
      if ($connStr -notmatch '(?i)(^|;)server=') { $warnings += "Connection string sin 'Server='." }
      if ($connStr -notmatch '(?i)(^|;)database=') { $warnings += "Connection string sin 'Database='." }
    }
    if ([string]::IsNullOrWhiteSpace($table)) {
      $errors += "Nombre de tabla SQL vacío."
    }
    elseif ($table -notmatch '^[A-Za-z0-9_\.\[\]]+$') {
      $errors += "Nombre de tabla SQL con caracteres no permitidos."
    }

    if ($DeepTest -and ($errors.Count -eq 0)) {
      try {
        $cn = New-Object System.Data.SqlClient.SqlConnection($connStr)
        try {
          $cn.Open()
          $cmd = $cn.CreateCommand()
          $cmd.CommandText = "SELECT 1"
          [void]$cmd.ExecuteScalar()

          $cmd2 = $cn.CreateCommand()
          $cmd2.CommandText = "SELECT OBJECT_ID(@t)"
          [void]$cmd2.Parameters.Add('@t', [System.Data.SqlDbType]::NVarChar, 512)
          $cmd2.Parameters['@t'].Value = $table
          $objId = $cmd2.ExecuteScalar()
          if ($null -eq $objId -or $objId -eq [DBNull]::Value) {
            $warnings += "La tabla no existe o no es visible para la conexión: $table"
          }
          else {
            $lines += " - Conexión SQL y tabla: OK"
          }
        }
        finally {
          if ($cn) { $cn.Dispose() }
        }
      }
      catch {
        $errors += "Error de conexión SQL: $($_.Exception.Message)"
      }
    }
  }

  if ($script:AuditLastCentralOk) {
    $lines += " - Ultimo envio central OK: $($script:AuditLastCentralOk)"
  }
  if ($script:AuditLastCentralError) {
    $warnings += "Ultimo error central registrado: $($script:AuditLastCentralError)"
  }

  if ($errors.Count -gt 0) {
    $lines += ""
    $lines += "Errores:"
    foreach ($e in $errors) { $lines += " - $e" }
  }
  if ($warnings.Count -gt 0) {
    $lines += ""
    $lines += "Advertencias:"
    foreach ($w in $warnings) { $lines += " - $w" }
  }
  if (($errors.Count -eq 0) -and ($warnings.Count -eq 0)) {
    $lines += ""
    $lines += "Diagnostico: configuracion valida."
  }

  return [PSCustomObject]@{
    IsValid   = ($errors.Count -eq 0)
    HasWarn   = ($warnings.Count -gt 0)
    Lines     = @($lines)
    ErrorText = ($errors -join ' | ')
  }
}

function Get-AuditOperator {
  if ($global:activeDomain -and $global:activeDomain -ne $script:LocalDomainLabel) {
    $conn = $global:domainConnections[$global:activeDomain]
    if ($conn -and $conn.Credential -and $conn.Credential.UserName) {
      return [string]$conn.Credential.UserName
    }
  }
  try { return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name } catch { return 'Unknown' }
}

function Get-AuditDomain {
  if ($global:activeDomain) { return [string]$global:activeDomain }
  return $script:LocalDomainLabel
}

function Get-AuditLocalLogPath([datetime]$Date = (Get-Date)) {
  $fileName = "{0}{1}.csv" -f $script:AuditLogFilePrefix, $Date.ToString('yyyy-MM-dd')
  return (Join-Path $script:AuditLogDir $fileName)
}

function Remove-AuditLocalExpiredFiles {
  if (-not (Test-Path -Path $script:AuditLogDir -PathType Container)) { return }

  $retentionMonths = [Math]::Max(1, [int]$script:AuditLocalRetentionMonths)
  $oldestToKeep = (Get-Date).Date.AddMonths(-$retentionMonths)
  $files = Get-ChildItem -Path $script:AuditLogDir -Filter ($script:AuditLogFilePrefix + '*.csv') -File -ErrorAction SilentlyContinue

  foreach ($file in @($files)) {
    if ($file.BaseName -match '^AD_Audit_(\d{4}-\d{2}-\d{2})$') {
      try {
        $fileDate = [datetime]::ParseExact($matches[1], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
        if ($fileDate.Date -lt $oldestToKeep) {
          Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
        }
      }
      catch {
        # Ignorar archivos con nombre inesperado.
      }
    }
  }
}

function Ensure-AuditLogFile([string]$path) {
  if ([string]::IsNullOrWhiteSpace($path)) { return }
  if (-not (Test-Path -Path $script:AuditLogDir -PathType Container)) {
    New-Item -Path $script:AuditLogDir -ItemType Directory -Force | Out-Null
  }
  if (-not (Test-Path -Path $path -PathType Leaf)) {
    [PSCustomObject]@{
      FechaHora = ''
      Operador  = ''
      Dominio   = ''
      Accion    = ''
      Resultado = ''
      Objeto    = ''
      Detalle   = ''
    } | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    $existing = Get-Content -Path $path
    if ($existing.Count -gt 1) {
      Set-Content -Path $path -Value $existing[0] -Encoding UTF8
    }
  }
}

function New-AuditEntry(
  [string]$Action,
  [string]$Result,
  [string]$Object,
  [string]$Detail = ''
) {
  return [PSCustomObject]@{
    FechaHora = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Operador  = Get-AuditOperator
    Dominio   = Get-AuditDomain
    Accion    = $Action
    Resultado = $Result
    Objeto    = $Object
    Detalle   = $Detail
    Equipo    = $env:COMPUTERNAME
  }
}

function Ensure-CentralFileLog([string]$path) {
  if ([string]::IsNullOrWhiteSpace($path)) { return }
  $parent = Split-Path -Path $path -Parent
  if (-not [string]::IsNullOrWhiteSpace($parent) -and (-not (Test-Path -Path $parent -PathType Container))) {
    New-Item -Path $parent -ItemType Directory -Force | Out-Null
  }
  if (-not (Test-Path -Path $path -PathType Leaf)) {
    [PSCustomObject]@{
      FechaHora = ''
      Operador  = ''
      Dominio   = ''
      Accion    = ''
      Resultado = ''
      Objeto    = ''
      Detalle   = ''
      Equipo    = ''
    } | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    $existing = Get-Content -Path $path
    if ($existing.Count -gt 1) {
      Set-Content -Path $path -Value $existing[0] -Encoding UTF8
    }
  }
}

function Write-AuditCentral([object]$entry) {
  if (-not $script:AuditSettings.EnableCentral) { return }

  switch ($script:AuditSettings.CentralMode) {
    'FileShare' {
      $centralPath = [string]$script:AuditSettings.CentralFilePath
      if ([string]::IsNullOrWhiteSpace($centralPath)) { return }
      Ensure-CentralFileLog -path $centralPath
      $line = $entry | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1
      Add-Content -Path $centralPath -Value $line -Encoding UTF8
      $script:AuditLastCentralError = ''
      $script:AuditLastCentralOk = ("{0} | FileShare | {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $centralPath)
    }
    'SqlServer' {
      $connStr = [string]$script:AuditSettings.SqlConnectionString
      $table = [string]$script:AuditSettings.SqlTable
      if ([string]::IsNullOrWhiteSpace($connStr) -or [string]::IsNullOrWhiteSpace($table)) { return }
      if ($table -notmatch '^[A-Za-z0-9_\.\[\]]+$') { throw "SqlTable contiene caracteres no permitidos: $table" }

      $cn = New-Object System.Data.SqlClient.SqlConnection($connStr)
      try {
        $cn.Open()
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = "INSERT INTO $table (FechaHora, Operador, Dominio, Accion, Resultado, Objeto, Detalle, Equipo) VALUES (@FechaHora, @Operador, @Dominio, @Accion, @Resultado, @Objeto, @Detalle, @Equipo)"
        [void]$cmd.Parameters.Add('@FechaHora', [System.Data.SqlDbType]::DateTime2)
        [void]$cmd.Parameters.Add('@Operador', [System.Data.SqlDbType]::NVarChar, 255)
        [void]$cmd.Parameters.Add('@Dominio', [System.Data.SqlDbType]::NVarChar, 255)
        [void]$cmd.Parameters.Add('@Accion', [System.Data.SqlDbType]::NVarChar, 255)
        [void]$cmd.Parameters.Add('@Resultado', [System.Data.SqlDbType]::NVarChar, 64)
        [void]$cmd.Parameters.Add('@Objeto', [System.Data.SqlDbType]::NVarChar, 1024)
        [void]$cmd.Parameters.Add('@Detalle', [System.Data.SqlDbType]::NVarChar, -1)
        [void]$cmd.Parameters.Add('@Equipo', [System.Data.SqlDbType]::NVarChar, 255)

        $cmd.Parameters['@FechaHora'].Value = [datetime]::ParseExact($entry.FechaHora, 'yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        $cmd.Parameters['@Operador'].Value = [string]$entry.Operador
        $cmd.Parameters['@Dominio'].Value = [string]$entry.Dominio
        $cmd.Parameters['@Accion'].Value = [string]$entry.Accion
        $cmd.Parameters['@Resultado'].Value = [string]$entry.Resultado
        $cmd.Parameters['@Objeto'].Value = [string]$entry.Objeto
        $cmd.Parameters['@Detalle'].Value = [string]$entry.Detalle
        $cmd.Parameters['@Equipo'].Value = [string]$entry.Equipo
        [void]$cmd.ExecuteNonQuery()
        $script:AuditLastCentralError = ''
        $script:AuditLastCentralOk = ("{0} | SqlServer | {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $table)
      }
      finally {
        if ($cn) { $cn.Dispose() }
      }
    }
    default {
      throw "Modo central no soportado: $($script:AuditSettings.CentralMode)"
    }
  }
}

function Write-AuditLog(
  [string]$Action,
  [string]$Result,
  [string]$Object,
  [string]$Detail = ''
) {
  $entry = New-AuditEntry -Action $Action -Result $Result -Object $Object -Detail $Detail

  try {
    $localLogPath = Get-AuditLocalLogPath
    Ensure-AuditLogFile -path $localLogPath
    $localEntry = $entry | Select-Object FechaHora, Operador, Dominio, Accion, Resultado, Objeto, Detalle
    $line = $localEntry | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1
    Add-Content -Path $localLogPath -Value $line -Encoding UTF8
    Remove-AuditLocalExpiredFiles
  }
  catch {
    # No interrumpir la operación principal por un fallo de auditoría.
  }

  try {
    Write-AuditCentral -entry $entry
  }
  catch {
    $script:AuditLastCentralError = $_.Exception.Message
    # No interrumpir la operación principal por un fallo de auditoría central.
  }
}

function Update-AuditModeUiState {
  if (-not $cbAuditMode) { return }
  $mode = [string]$cbAuditMode.SelectedItem
  if ([string]::IsNullOrWhiteSpace($mode)) { $mode = 'FileShare' }

  $isFileShare = ($mode -eq 'FileShare')
  $txtAuditCentralFilePath.IsEnabled = $isFileShare
  $txtAuditSqlConnection.IsEnabled = (-not $isFileShare)
  $txtAuditSqlTable.IsEnabled = (-not $isFileShare)
}

if ($cbAuditMode) {
  $cbAuditMode.Add_SelectionChanged({
      Update-AuditModeUiState
      $diag = Get-AuditValidationResult
      Show-AuditDiagnostics -lines $diag.Lines
    })
}

if ($chkAuditEnableCentral) {
  $saveAuditToggle = {
    if ($script:AuditUiSyncInProgress) { return }
    try {
      Set-AuditSettingsFromGui
      Save-AuditSettingsToFile
      $diag = Get-AuditValidationResult
      Show-AuditDiagnostics -lines $diag.Lines
    }
    catch {
      # No interrumpir el flujo por fallos de persistencia del tilde.
    }
  }
  $chkAuditEnableCentral.Add_Checked($saveAuditToggle)
  $chkAuditEnableCentral.Add_Unchecked($saveAuditToggle)
}

if ($btnAuditApplyConfig) {
  $btnAuditApplyConfig.Add_Click({
      try {
        Set-AuditSettingsFromGui
        Update-AuditModeUiState
        $diag = Get-AuditValidationResult
        Show-AuditDiagnostics -lines $diag.Lines
        if (-not $diag.IsValid) {
          Set-Status "❌ Configuración inválida. Revisá Diagnóstico." '#f38ba8'
          return
        }
        Set-Status "✅ Configuración de auditoría aplicada." '#a6e3a1'
        Write-AuditLog -Action 'Aplicar Config Auditoria' -Result 'OK' -Object ([string]$script:AuditSettings.CentralMode)
      }
      catch { Show-Exception "Error aplicando configuración de auditoría:" $_ }
    })
}

if ($btnAuditSaveConfig) {
  $btnAuditSaveConfig.Add_Click({
      try {
        Set-AuditSettingsFromGui
        $diag = Get-AuditValidationResult
        Show-AuditDiagnostics -lines $diag.Lines
        if (-not $diag.IsValid) {
          Set-Status "❌ No se guardó: configuración inválida." '#f38ba8'
          return
        }
        Save-AuditSettingsToFile
        Set-Status "✅ Configuración guardada en: $($script:AuditConfigPath)" '#a6e3a1'
        Write-AuditLog -Action 'Guardar Config Auditoria' -Result 'OK' -Object $script:AuditConfigPath
      }
      catch { Show-Exception "Error guardando configuración de auditoría:" $_ }
    })
}

if ($btnAuditLoadConfig) {
  $btnAuditLoadConfig.Add_Click({
      try {
        if (Load-AuditSettingsFromFile) {
          Update-AuditGuiFromSettings
          Update-AuditModeUiState
          $diag = Get-AuditValidationResult
          Show-AuditDiagnostics -lines $diag.Lines
          if ($diag.IsValid) {
            Set-Status "✅ Configuración de auditoría cargada desde: $($script:AuditConfigPath)" '#a6e3a1'
          }
          else {
            Set-Status "⚠️ Configuración cargada con errores." '#fab387'
          }
          Write-AuditLog -Action 'Cargar Config Auditoria' -Result 'OK' -Object $script:AuditConfigPath
        }
        else {
          Show-AuditDiagnostics -lines @(
            "No existe archivo de configuración en ninguna ruta candidata.",
            "Rutas evaluadas: $(Get-AuditConfigCandidatesText)"
          )
          Set-Status "⚠️ No hay archivo de configuración de auditoría." '#fab387'
        }
      }
      catch { Show-Exception "Error cargando configuración de auditoría:" $_ }
    })
}

if ($btnAuditDiagnoseConfig) {
  $btnAuditDiagnoseConfig.Add_Click({
      try {
        Set-AuditSettingsFromGui
        $diag = Get-AuditValidationResult -DeepTest
        Show-AuditDiagnostics -lines $diag.Lines
        if ($diag.IsValid) {
          if ($diag.HasWarn) {
            Set-Status "⚠️ Diagnóstico OK con advertencias." '#fab387'
          }
          else {
            Set-Status "✅ Diagnóstico OK." '#a6e3a1'
          }
        }
        else {
          Set-Status "❌ Diagnóstico con errores." '#f38ba8'
        }
      }
      catch { Show-Exception "Error ejecutando diagnóstico de auditoría:" $_ }
    })
}

if ($btnAuditTestConfig) {
  $btnAuditTestConfig.Add_Click({
      $prevEnabled = [bool]$script:AuditSettings.EnableCentral
      try {
        Set-AuditSettingsFromGui
        $diag = Get-AuditValidationResult -DeepTest
        Show-AuditDiagnostics -lines $diag.Lines
        if (-not $diag.IsValid) {
          Set-Status "❌ No se ejecutó prueba: configuración inválida." '#f38ba8'
          return
        }

        $script:AuditSettings.EnableCentral = $true
        $testEntry = New-AuditEntry -Action 'Test Destino Auditoria' -Result 'TEST' -Object 'GUI Config'
        Write-AuditCentral -entry $testEntry
        $script:AuditSettings.EnableCentral = $prevEnabled
        $diagAfter = Get-AuditValidationResult
        Show-AuditDiagnostics -lines $diagAfter.Lines
        Set-Status "✅ Prueba de auditoría central OK." '#a6e3a1'
      }
      catch {
        $script:AuditSettings.EnableCentral = $prevEnabled
        Show-Exception "Error probando destino central de auditoría:" $_
      }
    })
}

# Persistir configuración de auditoría al cerrar la ventana, incluso si no se presiona "Guardar Configuración".
$window.Add_Closing({
    try {
      Set-AuditSettingsFromGui
      Save-AuditSettingsToFile
      Save-DomainConnectionsToFile
      Save-ThemePreference
    }
    catch {
      # No bloquear el cierre de la app por un error de guardado.
    }
  })

function Show-Error([string]$msg) {
  Set-Status "❌ $msg" '#f38ba8'
  [System.Windows.MessageBox]::Show($msg, "Error", "OK", "Error") | Out-Null
}
function Show-Confirm([string]$msg) {
  return ([System.Windows.MessageBox]::Show($msg, "Confirmar", "YesNo", "Warning") -eq 'Yes')
}
function Get-Checked([System.Windows.Controls.CheckBox]$cb) { return ($cb -and $cb.IsChecked -eq $true) }
function Test-Guard([bool]$condition, [string]$message) {
  if ($condition) { return $true }
  Show-Error $message
  return $false
}
function Get-RequiredTrimmedText([string]$text, [string]$message) {
  $value = $text.Trim()
  if ([string]::IsNullOrWhiteSpace($value)) {
    Show-Error $message
    return $null
  }
  return $value
}
function Merge-ADParams([hashtable]$params, [hashtable]$adParams = $null) {
  if ($null -eq $params) { return @{} }
  $merged = @{}
  foreach ($key in $params.Keys) { $merged[$key] = $params[$key] }
  if ($null -eq $adParams) { $adParams = Get-ADParams }
  foreach ($key in $adParams.Keys) { $merged[$key] = $adParams[$key] }
  return $merged
}
function Get-NormalizedText([object]$value) {
  if ($null -eq $value) { return '' }
  return ([string]$value).Trim()
}
function Escape-ADFilterLiteral([string]$value) {
  if ($null -eq $value) { return '' }
  return $value.Replace("'", "''")
}
function Get-PolicyValue([object]$obj, [string]$name, $defaultValue = $null) {
  if ($null -eq $obj) { return $defaultValue }
  if ($obj -is [System.Collections.IDictionary]) {
    $hasKey = $false
    try {
      if ($obj.PSObject -and $obj.PSObject.Methods['ContainsKey']) {
        $hasKey = [bool]($obj.ContainsKey($name))
      }
      else {
        $hasKey = [bool]($obj.Contains($name))
      }
    }
    catch {
      $hasKey = [bool]($obj.Contains($name))
    }
    if ($hasKey) { return $obj[$name] }
  }
  if ($obj.PSObject -and $obj.PSObject.Properties[$name]) { return $obj.PSObject.Properties[$name].Value }
  return $defaultValue
}
function Test-PolicyPropertyExists([object]$obj, [string]$name) {
  if ($null -eq $obj) { return $false }
  if ($obj -is [System.Collections.IDictionary]) {
    try {
      if ($obj.PSObject -and $obj.PSObject.Methods['ContainsKey']) { return [bool]($obj.ContainsKey($name)) }
      return [bool]($obj.Contains($name))
    }
    catch {
      return [bool]($obj.Contains($name))
    }
  }
  return [bool]($obj.PSObject -and $obj.PSObject.Properties[$name])
}
function Set-PolicyValue([object]$obj, [string]$name, $value) {
  if ($null -eq $obj) { return }
  if ($obj -is [System.Collections.IDictionary]) {
    $obj[$name] = $value
    return
  }
  if ($obj.PSObject -and $obj.PSObject.Properties[$name]) {
    $obj.PSObject.Properties[$name].Value = $value
    return
  }
  if ($obj.PSObject) {
    $obj | Add-Member -MemberType NoteProperty -Name $name -Value $value -Force
  }
}
function Get-UserAdminKeyHash([object]$policy = $null) {
  $p = $policy
  if ($null -eq $p) { $p = $script:UserPolicy }
  if ($null -eq $p) { return '' }

  $hash = [string](Get-PolicyValue $p 'AdminKeyHash' '')
  if (-not [string]::IsNullOrWhiteSpace($hash)) { return $hash.Trim() }

  if ($p -is [System.Collections.IDictionary]) {
    foreach ($k in @($p.Keys)) {
      if ([string]$k -ieq 'AdminKeyHash') {
        return ([string]$p[$k]).Trim()
      }
    }
  }

  return ''
}
function Set-UserAdminKeyHash([string]$hashValue) {
  if ($null -eq $script:UserPolicy) {
    $script:UserPolicy = Normalize-UserPolicy -rawPolicy $null
  }

  Set-PolicyValue -obj $script:UserPolicy -name 'AdminKeyHash' -value ([string]$hashValue)
  if ($script:UserPolicy -is [System.Collections.IDictionary]) {
    $script:UserPolicy['AdminKeyHash'] = [string]$hashValue
  }

  # Canonicalizar para persistencia estable y volver a afirmar el hash.
  $script:UserPolicy = Normalize-UserPolicy -rawPolicy $script:UserPolicy
  Set-PolicyValue -obj $script:UserPolicy -name 'AdminKeyHash' -value ([string]$hashValue)
}
function ConvertTo-Sha256Hex([string]$text) {
  if ([string]::IsNullOrEmpty($text)) { return '' }
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    $hash = $sha.ComputeHash($bytes)
    return ([System.BitConverter]::ToString($hash)).Replace('-', '').ToLowerInvariant()
  }
  finally {
    if ($sha) { $sha.Dispose() }
  }
}
function Get-DefaultUserPolicy {
  $baseFields = @(
    'FirstName', 'Initials', 'LastName', 'FullName', 'DisplayName', 'Description',
    'Office', 'OfficePhone', 'Email', 'WebPage',
    'StreetAddress', 'POBox', 'City', 'State', 'PostalCode', 'Country',
    'UpnUser', 'SamAccountName', 'Password', 'OU',
    'Enabled', 'MustChangePassword', 'CannotChangePassword', 'PasswordNeverExpires'
  )
  return [ordered]@{
    RequireKeyForCreate = $false
    RequireKeyForModify = $false
    RequireKeyForPolicyEdit = $true
    RequireKeyForKeyChange  = $true
    AdminKeyHash        = ''
    Create              = [ordered]@{
      Visible  = @($baseFields)
      Editable = @($baseFields)
      Required = @('FirstName', 'LastName', 'FullName', 'UpnUser', 'SamAccountName', 'Password', 'OU')
    }
    Modify              = [ordered]@{
      Visible  = @($baseFields)
      Editable = @($baseFields)
      Required = @('FirstName', 'LastName', 'FullName', 'UpnUser', 'SamAccountName', 'OU')
    }
  }
}
function Get-UserPolicyFieldCatalog {
  return @(
    [PSCustomObject]@{ Name = 'FirstName'; Label = 'Nombre' },
    [PSCustomObject]@{ Name = 'Initials'; Label = 'Iniciales' },
    [PSCustomObject]@{ Name = 'LastName'; Label = 'Apellido' },
    [PSCustomObject]@{ Name = 'FullName'; Label = 'Nombre completo (CN)' },
    [PSCustomObject]@{ Name = 'DisplayName'; Label = 'Nombre para mostrar' },
    [PSCustomObject]@{ Name = 'Description'; Label = 'Descripción' },
    [PSCustomObject]@{ Name = 'Office'; Label = 'Oficina' },
    [PSCustomObject]@{ Name = 'OfficePhone'; Label = 'Teléfono' },
    [PSCustomObject]@{ Name = 'Email'; Label = 'Email' },
    [PSCustomObject]@{ Name = 'WebPage'; Label = 'Página Web' },
    [PSCustomObject]@{ Name = 'StreetAddress'; Label = 'Calle' },
    [PSCustomObject]@{ Name = 'POBox'; Label = 'Casilla Postal' },
    [PSCustomObject]@{ Name = 'City'; Label = 'Ciudad' },
    [PSCustomObject]@{ Name = 'State'; Label = 'Provincia/Estado' },
    [PSCustomObject]@{ Name = 'PostalCode'; Label = 'Código Postal' },
    [PSCustomObject]@{ Name = 'Country'; Label = 'País/Región' },
    [PSCustomObject]@{ Name = 'UpnUser'; Label = 'UPN (usuario)' },
    [PSCustomObject]@{ Name = 'SamAccountName'; Label = 'SAM' },
    [PSCustomObject]@{ Name = 'Password'; Label = 'Contraseña' },
    [PSCustomObject]@{ Name = 'OU'; Label = 'OU destino' },
    [PSCustomObject]@{ Name = 'Enabled'; Label = 'Cuenta habilitada' },
    [PSCustomObject]@{ Name = 'MustChangePassword'; Label = 'Debe cambiar contraseña' },
    [PSCustomObject]@{ Name = 'CannotChangePassword'; Label = 'No puede cambiar contraseña' },
    [PSCustomObject]@{ Name = 'PasswordNeverExpires'; Label = 'Contraseña no expira' }
  )
}
function Normalize-UserPolicy([object]$rawPolicy) {
  $base = Get-DefaultUserPolicy

  if ($null -eq $rawPolicy) { return $base }

  $base.RequireKeyForCreate = ConvertTo-BoolValue (Get-PolicyValue $rawPolicy 'RequireKeyForCreate' $base.RequireKeyForCreate) $true
  $base.RequireKeyForModify = ConvertTo-BoolValue (Get-PolicyValue $rawPolicy 'RequireKeyForModify' $base.RequireKeyForModify) $true
  $base.RequireKeyForPolicyEdit = ConvertTo-BoolValue (Get-PolicyValue $rawPolicy 'RequireKeyForPolicyEdit' $base.RequireKeyForPolicyEdit) $true
  $base.RequireKeyForKeyChange = ConvertTo-BoolValue (Get-PolicyValue $rawPolicy 'RequireKeyForKeyChange' $base.RequireKeyForKeyChange) $true
  $base.AdminKeyHash = Get-UserAdminKeyHash -policy $rawPolicy

  foreach ($mode in @('Create', 'Modify')) {
    $modeRaw = Get-PolicyValue $rawPolicy $mode $null
    $modeBase = $base[$mode]
    if ($null -eq $modeRaw) { continue }

    $hasVisible = Test-PolicyPropertyExists $modeRaw 'Visible'
    $hasEditable = Test-PolicyPropertyExists $modeRaw 'Editable'
    $hasRequired = Test-PolicyPropertyExists $modeRaw 'Required'
    $visible = @((Get-PolicyValue $modeRaw 'Visible' @()))
    $editable = @((Get-PolicyValue $modeRaw 'Editable' @()))
    $required = @((Get-PolicyValue $modeRaw 'Required' @()))

    if ($hasVisible) { $modeBase.Visible = @($visible | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) }
    if ($hasEditable) { $modeBase.Editable = @($editable | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) }
    if ($hasRequired) { $modeBase.Required = @($required | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) }
  }

  return $base
}
function Get-UserPolicyConfigReadPaths {
  # Ruta canónica: variable de entorno (si existe) o archivo local junto al script.
  # Las rutas secundarias quedan solo como compatibilidad de lectura/migración.
  $paths = @()
  $primary = Get-UserPolicyConfigWritePath
  if (-not [string]::IsNullOrWhiteSpace($primary)) { $paths += $primary }
  if ((-not [string]::IsNullOrWhiteSpace($script:UserPolicyConfigLocalPath)) -and ($script:UserPolicyConfigLocalPath -ne $primary)) { $paths += $script:UserPolicyConfigLocalPath }
  if ((-not [string]::IsNullOrWhiteSpace($script:UserPolicyConfigAppDataPath)) -and ($script:UserPolicyConfigAppDataPath -ne $primary)) { $paths += $script:UserPolicyConfigAppDataPath }
  if ((-not [string]::IsNullOrWhiteSpace($script:UserPolicyConfigProgramDataPath)) -and ($script:UserPolicyConfigProgramDataPath -ne $primary)) { $paths += $script:UserPolicyConfigProgramDataPath }
  foreach ($legacy in @($script:UserPolicyLegacyPaths)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$legacy)) { $paths += [string]$legacy }
  }
  return @($paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}
function Get-UserPolicyConfigWritePath {
  if (-not [string]::IsNullOrWhiteSpace($script:UserPolicyConfigEnvPath)) { return $script:UserPolicyConfigEnvPath }
  if (-not [string]::IsNullOrWhiteSpace($script:UserPolicyConfigLocalPath)) { return $script:UserPolicyConfigLocalPath }
  if (-not [string]::IsNullOrWhiteSpace($script:UserPolicyConfigAppDataPath)) { return $script:UserPolicyConfigAppDataPath }
  if (-not [string]::IsNullOrWhiteSpace($script:UserPolicyConfigPath)) { return $script:UserPolicyConfigPath }
  return (Join-Path (Get-Location).Path 'user.form.policy.json')
}
function Read-UserPolicyFile([string]$path) {
  if ([string]::IsNullOrWhiteSpace($path)) { throw "Ruta de política inválida." }
  if (-not (Test-Path -Path $path -PathType Leaf)) { throw "No existe el archivo de política: $path" }
  $raw = Get-Content -Path $path -Raw -ErrorAction Stop
  if ([string]::IsNullOrWhiteSpace($raw)) { throw "El archivo de política está vacío: $path" }
  return ($raw | ConvertFrom-Json -ErrorAction Stop)
}
function Get-UserPolicyConfigCandidatesText {
  $paths = Get-UserPolicyConfigReadPaths
  if ($null -eq $paths -or $paths.Count -eq 0) { return '(sin rutas configuradas)' }
  return ($paths -join ' | ')
}
function Write-UserPolicyFile([string]$path, [string]$json) {
  if ([string]::IsNullOrWhiteSpace($path)) { throw "Ruta de política inválida." }
  $parent = Split-Path -Path $path -Parent
  if ((-not [string]::IsNullOrWhiteSpace($parent)) -and (-not (Test-Path -Path $parent -PathType Container))) {
    New-Item -Path $parent -ItemType Directory -Force | Out-Null
  }
  Set-Content -Path $path -Value $json -Encoding UTF8
}
function Ensure-UserPolicyLoaded([switch]$ForceReload) {
  if ((-not $ForceReload) -and ($null -ne $script:UserPolicy)) { return }

  $currentPolicy = $script:UserPolicy
  $existingPaths = @()
  foreach ($candidate in (Get-UserPolicyConfigReadPaths)) {
    if (Test-Path -Path $candidate -PathType Leaf) { $existingPaths += $candidate }
  }

  # Respetar orden de prioridad definido en Get-UserPolicyConfigReadPaths.

  $loaded = $null
  $loadedPath = ''
  $lastErrorMessage = ''
  foreach ($path in @($existingPaths)) {
    try {
      $loaded = Read-UserPolicyFile -path $path
      if ($loaded) {
        $loadedPath = $path
        break
      }
    }
    catch {
      $lastErrorMessage = $_.Exception.Message
    }
  }

  if ($loaded) {
    $script:UserPolicy = Normalize-UserPolicy -rawPolicy $loaded
    $script:UserPolicyConfigPath = $loadedPath

    # Migración automática: si vino de ruta legacy, persistir en ruta principal.
    $writePath = Get-UserPolicyConfigWritePath
    if ((-not [string]::IsNullOrWhiteSpace($writePath)) -and ($writePath -ne $loadedPath)) {
      try { [void](Save-UserPolicyToFile) } catch {}
    }
    return
  }

  if ((-not [string]::IsNullOrWhiteSpace($lastErrorMessage)) -and ($existingPaths.Count -gt 0)) {
    # Fail-safe: conservar política en memoria si falla la lectura/parseo.
    if ($null -ne $currentPolicy) {
      $script:UserPolicy = Normalize-UserPolicy -rawPolicy $currentPolicy
      return
    }
  }

  if ($null -ne $currentPolicy) {
    $script:UserPolicy = Normalize-UserPolicy -rawPolicy $currentPolicy
    return
  }

  $script:UserPolicy = Normalize-UserPolicy -rawPolicy $null
  $script:UserPolicyConfigPath = Get-UserPolicyConfigWritePath
}
function Save-UserPolicyToFile {
  if ($null -eq $script:UserPolicy) { Ensure-UserPolicyLoaded }
  if ($null -eq $script:UserPolicy) { return $false }
  $script:UserPolicy = Normalize-UserPolicy -rawPolicy $script:UserPolicy
  Set-PolicyValue -obj $script:UserPolicy -name 'AdminKeyHash' -value (Get-UserAdminKeyHash -policy $script:UserPolicy)
  $json = $script:UserPolicy | ConvertTo-Json -Depth 8

  $targetPath = Get-UserPolicyConfigWritePath
  if ([string]::IsNullOrWhiteSpace($targetPath)) {
    Set-Status "⚠️ No se pudo guardar la política de usuarios: ruta destino vacía." '#fab387'
    return $false
  }

  try {
    Write-UserPolicyFile -path $targetPath -json $json
    if (-not (Test-Path -Path $targetPath -PathType Leaf)) {
      throw "El archivo no quedó creado en disco."
    }
    $reloaded = Normalize-UserPolicy -rawPolicy (Read-UserPolicyFile -path $targetPath)
    try {
      $resolved = (Resolve-Path -Path $targetPath -ErrorAction Stop).Path
      $script:UserPolicyConfigPath = [string]$resolved
    }
    catch {
      $script:UserPolicyConfigPath = [string]$targetPath
    }
    $script:UserPolicy = $reloaded
    return $true
  }
  catch {
    $detail = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { 'Error desconocido.' }
    Set-Status "⚠️ No se pudo guardar la política de usuarios en: $targetPath" '#fab387'
    Write-AuditLog -Action 'Guardar Politica Usuarios' -Result 'ERROR' -Object $targetPath -Detail $detail
    return $false
  }
}
function Get-UserPolicyList([ValidateSet('Create', 'Modify')] [string]$Mode, [ValidateSet('Visible', 'Editable', 'Required')] [string]$ListName) {
  Ensure-UserPolicyLoaded
  $modeObj = Get-PolicyValue $script:UserPolicy $Mode $null
  $list = @((Get-PolicyValue $modeObj $ListName @()))
  if ($list.Count -eq 0) { return @() }
  return @($list | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}
function Test-UserFieldVisible([ValidateSet('Create', 'Modify')] [string]$Mode, [string]$Field) {
  return $true
}
function Test-UserFieldEditable([ValidateSet('Create', 'Modify')] [string]$Mode, [string]$Field) {
  return $true
}
function Test-UserFieldRequired([ValidateSet('Create', 'Modify')] [string]$Mode, [string]$Field) {
  return $false
}
function Test-UserFieldValue([ValidateSet('Create', 'Modify')] [string]$Mode, [string]$Field, [string]$Value, [string]$ErrorMessage) {
  if ((Test-UserFieldRequired -Mode $Mode -Field $Field) -and [string]::IsNullOrWhiteSpace($Value)) {
    Show-Error $ErrorMessage
    return $false
  }
  return $true
}
function Apply-UserFieldUiPolicy([ValidateSet('Create', 'Modify')] [string]$Mode, [string]$Field, $LabelControl, $InputControl) {
  $isVisible = Test-UserFieldVisible -Mode $Mode -Field $Field
  $isEditable = Test-UserFieldEditable -Mode $Mode -Field $Field
  $visibility = if ($isVisible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
  if ($LabelControl) { $LabelControl.Visibility = $visibility }
  if ($InputControl) {
    $InputControl.Visibility = $visibility
    if ($InputControl.PSObject -and $InputControl.PSObject.Properties['IsEnabled']) {
      $InputControl.IsEnabled = $isEditable
    }
  }
}
function New-DefaultSam([string]$firstName, [string]$lastName, [string]$fallback = 'usuario') {
  $fn = Get-NormalizedText $firstName
  $ln = Get-NormalizedText $lastName
  $base = ''
  if ((-not [string]::IsNullOrWhiteSpace($fn)) -and (-not [string]::IsNullOrWhiteSpace($ln))) {
    $base = "$($fn.Substring(0, 1))$ln"
  }
  elseif (-not [string]::IsNullOrWhiteSpace($ln)) {
    $base = $ln
  }
  elseif (-not [string]::IsNullOrWhiteSpace($fn)) {
    $base = $fn
  }
  else {
    $base = $fallback
  }
  $norm = (($base.ToLowerInvariant()) -replace '[^a-z0-9._-]', '')
  if ([string]::IsNullOrWhiteSpace($norm)) { return 'usuario' }
  return $norm
}
function Request-UserAdminUnlock([switch]$ForcePrompt) {
  Ensure-UserPolicyLoaded
  $hash = Get-UserAdminKeyHash
  if ([string]::IsNullOrWhiteSpace($hash)) {
    return $false
  }
  if ((-not $ForcePrompt) -and $script:UserPolicyUnlocked) { return $true }

  $dlg = New-DialogWindow "Desbloquear administración de usuarios" 430 230
  $sp = New-Object System.Windows.Controls.StackPanel
  $sp.Margin = '16'
  $sp.Children.Add((New-DlgLabel "Ingresá la clave de administración:"))
  $tbKey = New-DlgPasswordBox
  $sp.Children.Add($tbKey)
  $btnOk = New-DlgButton "Desbloquear" 'Warn'
  $btnCancel = New-DlgButton "Cancelar" 'Primary'
  $btnCancel.Margin = '0,6,0,0'
  $sp.Children.Add($btnOk)
  $sp.Children.Add($btnCancel)
  $dlg.Content = $sp

  $approved = $false
  $tbKey.Add_KeyDown({
      if ($_.Key -eq 'Return') {
        $btnOk.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
      }
    }.GetNewClosure())
  $btnOk.Add_Click({
      $input = [string]$tbKey.Password
      $inputHash = ConvertTo-Sha256Hex $input
      if ($inputHash -eq $hash) {
        $approved = $true
        $dlg.Close()
      }
      else {
        Show-Error "Clave incorrecta."
      }
    }.GetNewClosure())
  $btnCancel.Add_Click({ $dlg.Close() })
  $dlg.ShowDialog() | Out-Null

  if ($approved) {
    $script:UserPolicyUnlocked = $true
    Set-Status "✅ Administración de usuarios desbloqueada para esta sesión." '#a6e3a1'
    return $true
  }
  return $false
}
function Assert-UserAdminPermission([ValidateSet('Create', 'Modify')] [string]$Operation) {
  # Uso diario sin clave: crear/modificar usuarios no requiere desbloqueo.
  # La clave admin se reserva para editar política de campos y cambiar la clave.
  return $true
}
function Assert-UserPolicyAdminPermission([ValidateSet('PolicyEdit', 'KeyChange')] [string]$Operation) {
  Ensure-UserPolicyLoaded
  $hash = Get-UserAdminKeyHash
  if ([string]::IsNullOrWhiteSpace($hash)) {
    Ensure-UserPolicyLoaded -ForceReload
    $hash = Get-UserAdminKeyHash
  }

  if ([string]::IsNullOrWhiteSpace($hash)) {
    if ($Operation -eq 'PolicyEdit') {
      Show-Error "Primero configurá una clave de administración en 'Configurar Clave Admin'.`n`nArchivo de política: $script:UserPolicyConfigPath"
      return $false
    }
    return $true
  }

  $requireKey = $true
  if ($Operation -eq 'PolicyEdit') {
    $requireKey = ConvertTo-BoolValue (Get-PolicyValue $script:UserPolicy 'RequireKeyForPolicyEdit' $true) $true
  }
  else {
    $requireKey = ConvertTo-BoolValue (Get-PolicyValue $script:UserPolicy 'RequireKeyForKeyChange' $true) $true
  }

  if (-not $requireKey) { return $true }
  if (Request-UserAdminUnlock -ForcePrompt) { return $true }
  Show-Error "No autorizado. Se requiere clave de administración."
  return $false
}
function Show-UserPolicyEditor {
  Ensure-UserPolicyLoaded
  if (-not (Assert-UserPolicyAdminPermission -Operation 'PolicyEdit')) { return }

  $dlg = New-DialogWindow "Política de campos de usuarios" 1120 680
  $root = New-Object System.Windows.Controls.Grid
  $root.Margin = '12'
  $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = 'Auto' }))
  $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = '*' }))
  $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = 'Auto' }))

  $hint = New-Object System.Windows.Controls.TextBlock
  $hint.Text = "Marcá qué campos se ven/editan/requieren en Crear y Modificar."
  $hint.Margin = '0,0,0,8'
  $hint.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextSecondary
  [System.Windows.Controls.Grid]::SetRow($hint, 0)
  $root.Children.Add($hint) | Out-Null

  $createVisible = Get-UserPolicyList -Mode 'Create' -ListName 'Visible'
  $createEditable = Get-UserPolicyList -Mode 'Create' -ListName 'Editable'
  $createRequired = Get-UserPolicyList -Mode 'Create' -ListName 'Required'
  $modifyVisible = Get-UserPolicyList -Mode 'Modify' -ListName 'Visible'
  $modifyEditable = Get-UserPolicyList -Mode 'Modify' -ListName 'Editable'
  $modifyRequired = Get-UserPolicyList -Mode 'Modify' -ListName 'Required'

  $rows = @()
  foreach ($f in (Get-UserPolicyFieldCatalog)) {
    $rows += [PSCustomObject]@{
      FieldName       = [string]$f.Name
      FieldLabel      = [string]$f.Label
      C_Visible       = ($createVisible -contains [string]$f.Name)
      C_Editable      = ($createEditable -contains [string]$f.Name)
      C_Required      = ($createRequired -contains [string]$f.Name)
      M_Visible       = ($modifyVisible -contains [string]$f.Name)
      M_Editable      = ($modifyEditable -contains [string]$f.Name)
      M_Required      = ($modifyRequired -contains [string]$f.Name)
    }
  }

  $dg = New-Object System.Windows.Controls.DataGrid
  $dg.AutoGenerateColumns = $false
  $dg.CanUserAddRows = $false
  $dg.IsReadOnly = $false
  $dg.EnableRowVirtualization = $false
  $dg.EnableColumnVirtualization = $false
  $dg.ItemsSource = @($rows)
  $dg.SelectionMode = 'Single'
  $dg.Margin = '0,0,0,4'
  [void]$dg.Columns.Add((New-Object System.Windows.Controls.DataGridTextColumn -Property @{ Header = 'Campo'; Binding = (New-Object System.Windows.Data.Binding('FieldLabel')); IsReadOnly = $true; Width = 220 }))
  function New-PolicyCheckColumn([string]$header, [string]$path, [double]$width = 95) {
    $template = New-Object System.Windows.DataTemplate
    $factory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.CheckBox])
    $factory.SetValue([System.Windows.Controls.CheckBox]::HorizontalAlignmentProperty, [System.Windows.HorizontalAlignment]::Center)
    $factory.SetValue([System.Windows.Controls.CheckBox]::VerticalAlignmentProperty, [System.Windows.VerticalAlignment]::Center)
    $factory.SetValue([System.Windows.Controls.CheckBox]::IsThreeStateProperty, $false)
    # Single-click real en DataGrid: evita el primer click solo para enfocar celda.
    $factory.SetValue([System.Windows.Controls.CheckBox]::FocusableProperty, $false)
    $factory.SetValue([System.Windows.Controls.CheckBox]::IsTabStopProperty, $false)
    $binding = New-Object System.Windows.Data.Binding($path)
    $binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
    $binding.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
    $factory.SetBinding([System.Windows.Controls.CheckBox]::IsCheckedProperty, $binding)
    $template.VisualTree = $factory

    $col = New-Object System.Windows.Controls.DataGridTemplateColumn
    $col.Header = $header
    $col.Width = $width
    $col.CellTemplate = $template
    $col.CellEditingTemplate = $template
    return $col
  }

  [void]$dg.Columns.Add((New-PolicyCheckColumn -header 'Crear Ver' -path 'C_Visible'))
  [void]$dg.Columns.Add((New-PolicyCheckColumn -header 'Crear Edit' -path 'C_Editable'))
  [void]$dg.Columns.Add((New-PolicyCheckColumn -header 'Crear Req' -path 'C_Required'))
  [void]$dg.Columns.Add((New-PolicyCheckColumn -header 'Mod Ver' -path 'M_Visible'))
  [void]$dg.Columns.Add((New-PolicyCheckColumn -header 'Mod Edit' -path 'M_Editable'))
  [void]$dg.Columns.Add((New-PolicyCheckColumn -header 'Mod Req' -path 'M_Required'))
  # UX: un solo clic en la celda de permiso alterna el valor (sin doble clic).
  $dg.Add_PreviewMouseLeftButtonDown({
      param($sender, $e)
      try {
        if ($null -eq $e) { return }
        $dep = $e.OriginalSource -as [System.Windows.DependencyObject]
        while ($dep -and (-not ($dep -is [System.Windows.Controls.DataGridCell]))) {
          $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep)
        }
        if (-not $dep) { return }

        $cell = [System.Windows.Controls.DataGridCell]$dep
        $colHeader = [string]$cell.Column.Header
        $prop = switch ($colHeader) {
          'Crear Ver' { 'C_Visible' }
          'Crear Edit' { 'C_Editable' }
          'Crear Req' { 'C_Required' }
          'Mod Ver' { 'M_Visible' }
          'Mod Edit' { 'M_Editable' }
          'Mod Req' { 'M_Required' }
          default { '' }
        }
        if ([string]::IsNullOrWhiteSpace($prop)) { return }

        $row = $cell.DataContext
        if ($null -eq $row) { return }

        $current = ConvertTo-BoolValue (Get-PolicyValue $row $prop $false) $false
        Set-PolicyValue -obj $row -name $prop -value (-not $current)
        # El origen es PSCustomObject (sin INotifyPropertyChanged), refrescar para reflejar el tilde.
        $dg.Items.Refresh()
        $e.Handled = $true
      }
      catch {
        # Mantener comportamiento estándar si falla el toggle asistido.
      }
    }.GetNewClosure())
  [System.Windows.Controls.Grid]::SetRow($dg, 1)
  $root.Children.Add($dg) | Out-Null

  $actions = New-Object System.Windows.Controls.WrapPanel
  $actions.HorizontalAlignment = 'Right'
  [System.Windows.Controls.Grid]::SetRow($actions, 2)
  $root.Children.Add($actions) | Out-Null

  $btnSave = New-DlgButton "Guardar y aplicar" 'Success'
  $btnSave.Margin = '0,8,8,0'
  $btnJson = New-DlgButton "Editar JSON" 'Warn'
  $btnJson.Margin = '0,8,8,0'
  $btnCancel = New-DlgButton "Cancelar" 'Primary'
  $btnCancel.Margin = '0,8,0,0'
  $actions.Children.Add($btnSave) | Out-Null
  $actions.Children.Add($btnJson) | Out-Null
  $actions.Children.Add($btnCancel) | Out-Null

  $btnSave.Add_Click({
      try {
        # Mover foco para cerrar edición activa en checkbox/texto.
        $btnSave.Focus()
        # Si hay una celda en edición, forzar commit antes de leer ItemsSource.
        [void]$dg.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        [void]$dg.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
        $dg.UpdateLayout()

        $gridRows = @()
        foreach ($item in @($dg.Items)) {
          if ($null -eq $item) { continue }
          if ($item -is [System.Windows.Data.CollectionViewGroup]) { continue }
          if ($item -is [System.Windows.Controls.ItemCollection]) { continue }
          if ([string]$item -eq '{NewItemPlaceholder}') { continue }

          $fieldName = [string](Get-PolicyValue $item 'FieldName' '')
          if ([string]::IsNullOrWhiteSpace($fieldName)) { continue }

          $gridRows += $item
        }
        if ($gridRows.Count -eq 0) { $gridRows = @($rows) }

        $createVisibleOut = @()
        $createEditableOut = @()
        $createRequiredOut = @()
        $modifyVisibleOut = @()
        $modifyEditableOut = @()
        $modifyRequiredOut = @()
        foreach ($r in @($gridRows)) {
          if ($null -eq $r) { continue }
          $name = [string]$r.FieldName
          if ([string]::IsNullOrWhiteSpace($name)) { continue }
          $cv = ConvertTo-BoolValue $r.C_Visible $false
          $ce = ConvertTo-BoolValue $r.C_Editable $false
          $cr = ConvertTo-BoolValue $r.C_Required $false
          $mv = ConvertTo-BoolValue $r.M_Visible $false
          $me = ConvertTo-BoolValue $r.M_Editable $false
          $mr = ConvertTo-BoolValue $r.M_Required $false

          if ($cv) { $createVisibleOut += $name }
          if ($ce) { $createEditableOut += $name }
          if ($cr) { $createRequiredOut += $name }
          if ($mv) { $modifyVisibleOut += $name }
          if ($me) { $modifyEditableOut += $name }
          if ($mr) { $modifyRequiredOut += $name }
        }

        $script:UserPolicy = Normalize-UserPolicy -rawPolicy $script:UserPolicy
        Set-PolicyValue -obj (Get-PolicyValue $script:UserPolicy 'Create' $null) -name 'Visible' -value @($createVisibleOut | Select-Object -Unique)
        Set-PolicyValue -obj (Get-PolicyValue $script:UserPolicy 'Create' $null) -name 'Editable' -value @($createEditableOut | Select-Object -Unique)
        Set-PolicyValue -obj (Get-PolicyValue $script:UserPolicy 'Create' $null) -name 'Required' -value @($createRequiredOut | Select-Object -Unique)
        Set-PolicyValue -obj (Get-PolicyValue $script:UserPolicy 'Modify' $null) -name 'Visible' -value @($modifyVisibleOut | Select-Object -Unique)
        Set-PolicyValue -obj (Get-PolicyValue $script:UserPolicy 'Modify' $null) -name 'Editable' -value @($modifyEditableOut | Select-Object -Unique)
        Set-PolicyValue -obj (Get-PolicyValue $script:UserPolicy 'Modify' $null) -name 'Required' -value @($modifyRequiredOut | Select-Object -Unique)

        $script:UserPolicy = Normalize-UserPolicy -rawPolicy $script:UserPolicy
        if (-not (Save-UserPolicyToFile)) {
          throw "No se pudo persistir la política en archivo."
        }
        # Verificar persistencia real: recargar desde disco y aplicar.
        Ensure-UserPolicyLoaded -ForceReload
        $script:UserPolicy = Normalize-UserPolicy -rawPolicy $script:UserPolicy
        $savedPath = [string]$script:UserPolicyConfigPath
        if ([string]::IsNullOrWhiteSpace($savedPath)) { $savedPath = [string](Get-UserPolicyConfigWritePath) }
        if ([string]::IsNullOrWhiteSpace($savedPath)) { $savedPath = '(sin ruta resuelta)' }
        Set-Status "✅ Política guardada y aplicada. Ruta: $savedPath" '#a6e3a1'
        Write-AuditLog -Action 'Guardar Politica Usuarios' -Result 'OK' -Object $script:UserPolicyConfigPath
        $dlg.Close()
      }
      catch {
        $detail = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { 'Error desconocido.' }
        Show-Error "No se pudo guardar la política de campos. Verificá que la clave de administración esté autenticada.`n`nDetalle: $detail"
      }
    }.GetNewClosure())
  $btnJson.Add_Click({
      $jsonDlg = New-DialogWindow "Política JSON (avanzado)" 780 620
      $jsonRoot = New-Object System.Windows.Controls.Grid
      $jsonRoot.Margin = '12'
      $jsonRoot.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = '*' }))
      $jsonRoot.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = 'Auto' }))
      $tbJson = New-Object System.Windows.Controls.TextBox
      $tbJson.AcceptsReturn = $true
      $tbJson.AcceptsTab = $true
      $tbJson.TextWrapping = 'NoWrap'
      $tbJson.VerticalScrollBarVisibility = 'Auto'
      $tbJson.HorizontalScrollBarVisibility = 'Auto'
      $tbJson.FontFamily = 'Consolas'
      $tbJson.FontSize = 12
      $tbJson.Text = ($script:UserPolicy | ConvertTo-Json -Depth 8)
      [System.Windows.Controls.Grid]::SetRow($tbJson, 0)
      $jsonRoot.Children.Add($tbJson) | Out-Null
      $jsonActions = New-Object System.Windows.Controls.WrapPanel
      $jsonActions.HorizontalAlignment = 'Right'
      [System.Windows.Controls.Grid]::SetRow($jsonActions, 1)
      $jsonRoot.Children.Add($jsonActions) | Out-Null
      $btnJsonSave = New-DlgButton "Guardar JSON" 'Success'
      $btnJsonCancel = New-DlgButton "Cerrar" 'Primary'
      $btnJsonSave.Margin = '0,8,8,0'
      $btnJsonCancel.Margin = '0,8,0,0'
      $jsonActions.Children.Add($btnJsonSave) | Out-Null
      $jsonActions.Children.Add($btnJsonCancel) | Out-Null
      $btnJsonSave.Add_Click({
          try {
            $parsed = $tbJson.Text | ConvertFrom-Json -ErrorAction Stop
            $script:UserPolicy = Normalize-UserPolicy -rawPolicy $parsed
            Save-UserPolicyToFile
            Set-Status "✅ Política guardada desde JSON." '#a6e3a1'
            $jsonDlg.Close()
            $dlg.Close()
          }
          catch { Show-Error "JSON inválido." }
        }.GetNewClosure())
      $btnJsonCancel.Add_Click({ $jsonDlg.Close() })
      $jsonDlg.Content = $jsonRoot
      $jsonDlg.ShowDialog() | Out-Null
    }.GetNewClosure())
  $btnCancel.Add_Click({ $dlg.Close() })

  $dlg.Content = $root
  $dlg.ShowDialog() | Out-Null
}
function Set-UserAdminKey {
  Ensure-UserPolicyLoaded
  $script:UserPolicy = Normalize-UserPolicy -rawPolicy $script:UserPolicy
  $currentHash = Get-UserAdminKeyHash
  if (-not [string]::IsNullOrWhiteSpace($currentHash)) {
    if (-not (Assert-UserPolicyAdminPermission -Operation 'KeyChange')) { return }
  }

  $dlg = New-DialogWindow "Configurar clave de administración" 460 300
  $sp = New-Object System.Windows.Controls.StackPanel
  $sp.Margin = '16'
  $sp.Children.Add((New-DlgLabel "Nueva clave:"))
  $tbNew = New-DlgPasswordBox
  $sp.Children.Add($tbNew)
  $sp.Children.Add((New-DlgLabel "Repetir clave:"))
  $tbConfirm = New-DlgPasswordBox
  $sp.Children.Add($tbConfirm)
  $btnSave = New-DlgButton "Guardar clave" 'Success'
  $btnClear = New-DlgButton "Quitar clave" 'Danger'
  $btnClear.Margin = '0,6,0,0'
  $sp.Children.Add($btnSave)
  $sp.Children.Add($btnClear)
  $dlg.Content = $sp

  $btnSave.Add_Click({
      if (-not [string]::IsNullOrWhiteSpace($currentHash)) {
        if (-not (Request-UserAdminUnlock -ForcePrompt)) { return }
      }
      $newKey = [string]$tbNew.Password
      $confirmKey = [string]$tbConfirm.Password
      if ([string]::IsNullOrWhiteSpace($newKey)) {
        Show-Error "Ingresá una clave."
        return
      }
      if ($newKey -ne $confirmKey) {
        Show-Error "La confirmación no coincide."
        return
      }
      $newHash = ConvertTo-Sha256Hex $newKey
      if ([string]::IsNullOrWhiteSpace($newHash)) {
        Show-Error "No se pudo generar el hash de la clave."
        return
      }
      Set-UserAdminKeyHash -hashValue $newHash
      $hashInMemory = Get-UserAdminKeyHash
      if ([string]::IsNullOrWhiteSpace($hashInMemory)) {
        Show-Error "No se pudo aplicar la clave en memoria. Revisá permisos/estado de la política."
        return
      }
      $script:UserPolicyUnlocked = $true
      if (-not (Save-UserPolicyToFile)) {
        Show-Error "No se pudo guardar la clave en archivo.`nRuta: $script:UserPolicyConfigPath"
        return
      }
      Ensure-UserPolicyLoaded -ForceReload
      $hashReloaded = Get-UserAdminKeyHash
      if ([string]::IsNullOrWhiteSpace($hashReloaded)) {
        Show-Error "La clave no quedó persistida luego de recargar la política.`nRuta: $script:UserPolicyConfigPath"
        return
      }
      Set-Status "✅ Clave de administración actualizada." '#a6e3a1'
      Write-AuditLog -Action 'Configurar Clave Admin Usuarios' -Result 'OK' -Object 'UserPolicy'
      $dlg.Close()
    }.GetNewClosure())

  $btnClear.Add_Click({
      if ([string]::IsNullOrWhiteSpace($currentHash)) {
        Show-Error "No hay clave configurada para quitar."
        return
      }
      if (-not (Request-UserAdminUnlock -ForcePrompt)) { return }
      if (-not (Show-Confirm "¿Quitar la clave de administración de usuarios?")) { return }
      Set-UserAdminKeyHash -hashValue ''
      $script:UserPolicyUnlocked = $false
      if (-not (Save-UserPolicyToFile)) {
        Show-Error "No se pudo guardar la eliminación de clave.`nRuta: $script:UserPolicyConfigPath"
        return
      }
      Set-Status "✅ Clave de administración eliminada." '#a6e3a1'
      Write-AuditLog -Action 'Eliminar Clave Admin Usuarios' -Result 'OK' -Object 'UserPolicy'
      $dlg.Close()
    }.GetNewClosure())

  $dlg.ShowDialog() | Out-Null
}
function Get-TreeSelectionDN(
  [System.Windows.Controls.TreeView]$treeView,
  [string]$selectionMessage,
  [string]$invalidDnMessage,
  [switch]$DenyProtected,
  [string]$protectedMessage = "Por seguridad no se permite operar sobre 'Domain Controllers'."
) {
  $item = $treeView.SelectedItem
  if (-not (Test-Guard ($null -ne $item) $selectionMessage)) { return $null }

  $dn = [string]$item.Tag
  if (-not (Test-Guard (-not [string]::IsNullOrWhiteSpace($dn)) $invalidDnMessage)) { return $null }

  if ($DenyProtected -and (Is-ProtectedContainerDN $dn)) {
    Show-Error $protectedMessage
    return $null
  }
  return $dn
}

function Show-Exception([string]$context, $err) {
  $msgText = ''
  if ($err -and $err.Exception -and $err.Exception.Message) { $msgText = [string]$err.Exception.Message }
  Write-AuditLog -Action $context -Result 'ERROR' -Object '' -Detail $msgText

  $msg = @(
    "$context",
    "",
    "Mensaje: $($err.Exception.Message)",
    "",
    "Tipo: $($err.Exception.GetType().FullName)",
    "",
    "Stack:",
    ($err.ScriptStackTrace | Out-String)
  ) -join "`n"
  Show-Error $msg
}

function Test-DCConnectivityAtStartup([int]$TimeoutSeconds = 5, [object]$SplashUi = $null) {
  $script:StartupDcError = ''
  $script:StartupDomainChecks = @()
  $script:StartupDomainFailures = @()
  $timeout = [Math]::Max(2, $TimeoutSeconds)

  $targets = @()
  $targets += [PSCustomObject]@{
    Label      = $script:LocalDomainLabel
    Server     = ''
    Credential = $null
    IsLocal    = $true
  }

  foreach ($label in @($global:domainConnections.Keys | Sort-Object)) {
    $conn = $global:domainConnections[$label]
    if (-not $conn) { continue }
    $targets += [PSCustomObject]@{
      Label      = [string]$label
      Server     = [string]$conn.Server
      Credential = $conn.Credential
      IsLocal    = $false
    }
  }

  if ($targets.Count -eq 0) { return $false }

  $checkSingleTarget = {
    param(
      [string]$Label,
      [string]$Server,
      $Credential,
      [bool]$IsLocalTarget,
      [int]$ProgressStart,
      [int]$ProgressEnd
    )

    $job = $null
    try {
      $job = Start-Job -ScriptBlock {
        param($serverArg, $credArg)

        Import-Module ActiveDirectory -ErrorAction Stop
        $params = @{}
        if (-not [string]::IsNullOrWhiteSpace([string]$serverArg)) { $params.Server = [string]$serverArg }
        if ($null -ne $credArg) { $params.Credential = $credArg }

        $domain = Get-ADDomain @params -ErrorAction Stop
        Get-ADRootDSE @params -ErrorAction Stop | Out-Null

        [PSCustomObject]@{
          Ok                 = $true
          DistinguishedName  = [string]$domain.DistinguishedName
          DNSRoot            = [string]$domain.DNSRoot
          UsersContainer     = [string]$domain.UsersContainer
          Server             = [string]$serverArg
          CredUser           = if ($credArg) { [string]$credArg.UserName } else { '' }
        }
      } -ArgumentList $Server, $Credential

      $waitTickSeconds = 1
      $elapsedSeconds = 0.0
      while ($true) {
        if (Wait-Job -Job $job -Timeout $waitTickSeconds) { break }
        $elapsedSeconds += $waitTickSeconds

        if ($SplashUi) {
          $ratio = [Math]::Min(1.0, ($elapsedSeconds / [double]$timeout))
          $currentProgress = $ProgressStart + [int]([Math]::Round(($ProgressEnd - $ProgressStart) * $ratio))
          Set-StartupSplashProgress -splashUi $SplashUi -percent $currentProgress -message "Validando dominio: $Label"
        }

        if ($elapsedSeconds -ge $timeout) {
          Stop-Job -Job $job -ErrorAction SilentlyContinue
          return [PSCustomObject]@{
            Ok      = $false
            Label   = $Label
            IsLocal = $IsLocalTarget
            Error   = "Tiempo de espera agotado (${timeout}s)."
          }
        }
      }

      $result = Receive-Job -Job $job -ErrorAction Stop
      if (-not $result -or (-not $result.Ok)) {
        return [PSCustomObject]@{
          Ok      = $false
          Label   = $Label
          IsLocal = $IsLocalTarget
          Error   = "No se obtuvo respuesta válida del controlador de dominio."
        }
      }

      $cacheServer = '(local)'
      if (-not [string]::IsNullOrWhiteSpace([string]$result.Server)) { $cacheServer = [string]$result.Server }
      $cacheCred = [string]$result.CredUser
      $cacheKey = "$cacheServer|$cacheCred"
      $script:ADDomainCache[$cacheKey] = [PSCustomObject]@{
        DistinguishedName = [string]$result.DistinguishedName
        DNSRoot           = [string]$result.DNSRoot
        UsersContainer    = [string]$result.UsersContainer
      }

      if ($IsLocalTarget -and (-not [string]::IsNullOrWhiteSpace([string]$result.DNSRoot))) {
        Set-LocalDomainIdentity -dnsRoot ([string]$result.DNSRoot)
      }

      return [PSCustomObject]@{
        Ok      = $true
        Label   = $Label
        IsLocal = $IsLocalTarget
        Error   = ''
      }
    }
    catch {
      $errMsg = if ($_.Exception -and $_.Exception.Message) { [string]$_.Exception.Message } else { "No se pudo contactar un controlador de dominio." }
      return [PSCustomObject]@{
        Ok      = $false
        Label   = $Label
        IsLocal = $IsLocalTarget
        Error   = $errMsg
      }
    }
    finally {
      if ($job) {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
      }
    }
  }

  $count = [Math]::Max(1, $targets.Count)
  $rangeStart = 50
  $rangeEnd = 88
  $span = [Math]::Max(1, ($rangeEnd - $rangeStart))
  $localOk = $false

  for ($idx = 0; $idx -lt $targets.Count; $idx++) {
    $target = $targets[$idx]
    $currentIndex = $idx + 1
    $segmentStart = $rangeStart + [int]([Math]::Floor(($idx * $span) / $count))
    $segmentEnd = $rangeStart + [int]([Math]::Floor((($currentIndex) * $span) / $count))

    if ($SplashUi) {
      Set-StartupSplashProgress -splashUi $SplashUi -percent $segmentStart -message "Chequeando dominio [$currentIndex/$count]: $($target.Label)"
    }

    $check = & $checkSingleTarget `
      -Label ([string]$target.Label) `
      -Server ([string]$target.Server) `
      -Credential $target.Credential `
      -IsLocalTarget ([bool]$target.IsLocal) `
      -ProgressStart $segmentStart `
      -ProgressEnd $segmentEnd

    $script:StartupDomainChecks += $check

    if ($check.Ok) {
      if ($target.IsLocal) { $localOk = $true }
    }
    else {
      $script:StartupDomainFailures += $check
      if ($target.IsLocal) {
        $script:StartupDcError = [string]$check.Error
      }
    }
  }

  if (-not $localOk -and [string]::IsNullOrWhiteSpace($script:StartupDcError)) {
    $script:StartupDcError = "No se pudo validar el dominio local."
  }

  return $localOk
}

function Show-StartupDomainFailuresSummary {
  if ($null -eq $script:StartupDomainFailures -or $script:StartupDomainFailures.Count -eq 0) { return }

  $lines = @()
  foreach ($f in @($script:StartupDomainFailures)) {
    $lines += " - $($f.Label): $($f.Error)"
  }

  $msg = @(
    "Se detectaron dominios con falla de conectividad al iniciar:",
    "",
    ($lines -join "`n"),
    "",
    "Podés seguir usando la app con los dominios que sí respondieron."
  ) -join "`n"

  $failedLabels = @($script:StartupDomainFailures | Select-Object -ExpandProperty Label)
  Set-Status ("⚠️ Falló conectividad en: " + ($failedLabels -join ', ')) '#fab387'
  [System.Windows.MessageBox]::Show($msg, "Conectividad de dominios", "OK", "Warning") | Out-Null
}

function Show-FriendlyDcUnavailableMessage {
  $detail = if ([string]::IsNullOrWhiteSpace($script:StartupDcError)) {
    "No se pudo contactar un controlador de dominio."
  }
  else {
    $script:StartupDcError
  }

  $msg = @(
    "No se pudo conectar al controlador de dominio al iniciar.",
    "",
    "Qué podés revisar:",
    " - Conectividad de red/VPN",
    " - DNS del equipo",
    " - Nombre del dominio o servidor configurado",
    "",
    "Detalle técnico:",
    $detail
  ) -join "`n"

  Set-Status "⚠️ Sin conexión al DC. La app abrió en modo limitado." '#fab387'
  [System.Windows.MessageBox]::Show($msg, "Conexión a DC no disponible", "OK", "Warning") | Out-Null
}

function Get-OUFromDN([string]$dn) {
  $parts = $dn -split ',', 2
  if ($parts.Count -gt 1) { return $parts[1] } else { return $dn }
}

function Get-ContainerLabelFromDN([string]$dn) {
  if ([string]::IsNullOrWhiteSpace($dn)) { return '' }
  $segments = @()
  foreach ($part in ($dn -split ',')) {
    if ($part -match '^(OU|CN)=(.+)$') {
      $segments += $matches[2]
    }
  }
  if (-not $segments -or $segments.Count -eq 0) { return $dn }
  [array]::Reverse($segments)
  return ($segments -join ' / ')
}

function Is-ProtectedContainerDN([string]$dn) {
  if ([string]::IsNullOrWhiteSpace($dn)) { return $false }
  return ($dn -match '^(?i)(OU|CN)=Domain Controllers,')
}

function Get-SelectedUser {
  $sel = $dgUsers.SelectedItem
  if (-not $sel) { Show-Error "Seleccioná un usuario de la lista."; return $null }
  return $sel
}
function Get-SelectedGroup {
  $sel = $dgGroups.SelectedItem
  if (-not $sel) { Show-Error "Seleccioná un grupo de la lista."; return $null }
  return $sel
}

function Raise-ButtonClick([System.Windows.Controls.Button]$button) {
  if ($button) {
    $button.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
  }
}

function Bind-EnterKeyToButton([System.Windows.Controls.TextBox]$textBox, [System.Windows.Controls.Button]$button) {
  $textBox.Add_KeyDown({
      if ($_.Key -eq 'Return') {
        Raise-ButtonClick $button
      }
    }.GetNewClosure())
}

function Invoke-SelectedUserAction(
  [scriptblock]$Action,
  [string]$SuccessMessage,
  [string]$ErrorContext,
  [string]$ConfirmMessage = '',
  [string]$AuditAction = 'Acción Usuario'
) {
  if (-not (Assert-ADModule)) { return }
  $sel = Get-SelectedUser
  if (-not $sel) { return }

  if (-not [string]::IsNullOrWhiteSpace($ConfirmMessage)) {
    if (-not (Show-Confirm ($ConfirmMessage -f $sel.SamAccountName))) { return }
  }

  try {
    & $Action $sel
    Set-Status ($SuccessMessage -f $sel.SamAccountName) '#a6e3a1'
    Write-AuditLog -Action $AuditAction -Result 'OK' -Object ([string]$sel.SamAccountName)
    Raise-ButtonClick $btnUserSearch
  }
  catch {
    Write-AuditLog -Action $AuditAction -Result 'ERROR' -Object ([string]$sel.SamAccountName) -Detail $_.Exception.Message
    Show-Exception $ErrorContext $_
  }
}

function Invoke-SelectedGroupAction(
  [scriptblock]$Action,
  [string]$SuccessMessage,
  [string]$ErrorContext,
  [string]$ConfirmMessage = '',
  [string]$AuditAction = 'Acción Grupo'
) {
  if (-not (Assert-ADModule)) { return }
  $sel = Get-SelectedGroup
  if (-not $sel) { return }

  if (-not [string]::IsNullOrWhiteSpace($ConfirmMessage)) {
    if (-not (Show-Confirm ($ConfirmMessage -f $sel.Name))) { return }
  }

  try {
    & $Action $sel
    Set-Status ($SuccessMessage -f $sel.Name) '#a6e3a1'
    Write-AuditLog -Action $AuditAction -Result 'OK' -Object ([string]$sel.Name)
    Raise-ButtonClick $btnGroupSearch
  }
  catch {
    Write-AuditLog -Action $AuditAction -Result 'ERROR' -Object ([string]$sel.Name) -Detail $_.Exception.Message
    Show-Exception $ErrorContext $_
  }
}

function New-DialogWindow([string]$title, [int]$w = 450, [int]$h = 500) {
  $t = Get-ActiveTheme
  $dlg = New-Object System.Windows.Window
  $dlg.Title = "$title | Develop by Matias Vaccari"
  $dlg.Width = $w; $dlg.Height = $h
  $dlg.MinWidth = [Math]::Min($w, 420)
  $dlg.MinHeight = [Math]::Min($h, 260)
  $dlg.WindowStartupLocation = 'CenterOwner'
  $dlg.Owner = $window
  $dlg.Background = ConvertTo-Brush $t.WindowBg
  $dlg.Foreground = ConvertTo-Brush $t.TextPrimary
  $dlg.FontFamily = 'Segoe UI'
  return $dlg
}
function New-DlgLabel([string]$text) {
  $t = Get-ActiveTheme
  $l = New-Object System.Windows.Controls.Label
  $l.Content = $text
  $l.Foreground = ConvertTo-Brush $t.TextSecondary
  $l.FontSize = 13; $l.Padding = '0,6,0,2'
  return $l
}
function New-DlgTextBox([string]$text = '') {
  $theme = Get-ActiveTheme
  $tb = New-Object System.Windows.Controls.TextBox
  $tb.Text = $text; $tb.FontSize = 13; $tb.Padding = '6,4'
  $tb.Background = ConvertTo-Brush $theme.InputBg
  $tb.Foreground = ConvertTo-Brush $theme.InputFg
  $tb.BorderBrush = ConvertTo-Brush $theme.Border
  return $tb
}
function New-DlgPasswordBox {
  $t = Get-ActiveTheme
  $p = New-Object System.Windows.Controls.PasswordBox
  $p.FontSize = 13; $p.Padding = '6,4'
  $p.Background = ConvertTo-Brush $t.InputBg
  $p.Foreground = ConvertTo-Brush $t.InputFg
  $p.BorderBrush = ConvertTo-Brush $t.Border
  return $p
}
function New-DlgButton([string]$text, [string]$style = 'Primary') {
  $b = New-Object System.Windows.Controls.Button
  $b.Content = $text; $b.FontSize = 13; $b.FontWeight = 'SemiBold'; $b.Padding = '16,8'
  $b.BorderThickness = '0'; $b.Cursor = 'Hand'; $b.Margin = '0,10,0,0'
  $styleKey = "Btn$style"
  if ($window -and $window.Resources.Contains($styleKey)) {
    $b.Style = $window.Resources[$styleKey]
  }
  else {
    $colors = @{ Primary = '#89b4fa'; Success = '#a6e3a1'; Danger = '#f38ba8'; Warn = '#fab387' }
    $b.Background = ConvertTo-Brush $colors[$style]
    $b.Foreground = ConvertTo-Brush '#1e1e2e'
  }
  return $b
}
function New-DlgComboBox([string[]]$items, [int]$selected = 0) {
  $t = Get-ActiveTheme
  $c = New-Object System.Windows.Controls.ComboBox
  $items | ForEach-Object { $c.Items.Add($_) | Out-Null }
  $c.SelectedIndex = $selected; $c.FontSize = 13; $c.Padding = '6,4'
  $c.Background = ConvertTo-Brush $t.InputBg
  $c.Foreground = ConvertTo-Brush $t.InputFg
  $c.BorderBrush = ConvertTo-Brush $t.Border
  return $c
}
function New-ConfigActionButton([string]$content, [string]$styleKey) {
  $b = New-Object System.Windows.Controls.Button
  $b.Content = $content
  $b.Margin = '0,0,6,4'
  $b.Padding = '12,6'
  if ($window.Resources.Contains($styleKey)) {
    $b.Style = $window.Resources[$styleKey]
  }
  return $b
}
function Add-UserPolicyConfigButtons {
  if (-not $borderConfigActionsPanel) { return }
  $actionsHost = $borderConfigActionsPanel.Child
  if (-not ($actionsHost -is [System.Windows.Controls.WrapPanel])) { return }

  if (($actionsHost.Children | Where-Object { $_ -is [System.Windows.Controls.Button] -and $_.Name -eq 'btnUserAdminUnlock' }).Count -gt 0) {
    return
  }

  $btnUnlock = New-ConfigActionButton -content "🔐 Desbloquear Admin Usuarios" -styleKey "BtnWarn"
  $btnUnlock.Name = 'btnUserAdminUnlock'
  $btnPolicy = New-ConfigActionButton -content "🛠️ Política Campos Usuarios" -styleKey "BtnPrimary"
  $btnPolicy.Name = 'btnUserPolicyEditor'
  $btnKey = New-ConfigActionButton -content "🔑 Configurar Clave Admin" -styleKey "BtnDanger"
  $btnKey.Name = 'btnUserAdminKey'

  $btnUnlock.Add_Click({ [void](Request-UserAdminUnlock) })
  $btnPolicy.Add_Click({ Show-UserPolicyEditor })
  $btnKey.Add_Click({ Set-UserAdminKey })

  $actionsHost.Children.Add($btnUnlock) | Out-Null
  $actionsHost.Children.Add($btnPolicy) | Out-Null
  $actionsHost.Children.Add($btnKey) | Out-Null
}
# Administracion de campos deshabilitada. Se mantienen todos visibles/editables.

# ── FIX CORE: resolver host y testear RootDSE ────────────────────────────────
function Resolve-ServerForAD([string]$server) {
  $s = $server.Trim()
  if ($s -match '^\d{1,3}(\.\d{1,3}){3}$') {
    try {
      $entry = [System.Net.Dns]::GetHostEntry($s)
      if ($entry -and $entry.HostName) { return $entry.HostName }
    }
    catch { }
  }
  return $s
}

function Test-ADConnection([string]$server, [System.Management.Automation.PSCredential]$cred) {
  $srv = Resolve-ServerForAD $server
  $p = @{ Server = $srv }
  if ($cred) { $p.Credential = $cred }

  # RootDSE es el test más neutro (no depende de "Administrator" ni de OU)
  Get-ADRootDSE @p -ErrorAction Stop | Out-Null

  return $srv
}

# ── DOMAIN CONNECT / DISCONNECT ───────────────────────────────────────────────
$btnDomainConnect.Add_Click({
    if (-not (Assert-ADModule)) { return }

    $dlg = New-DialogWindow "Conectar a Dominio" 440 350
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'

    $sp.Children.Add((New-DlgLabel "Servidor / Controlador de Dominio:"))
    $sp.Children.Add((New-DlgLabel "(ej: dc01.empresa.com o empresa.com)"))
    $tbServer = New-DlgTextBox; $sp.Children.Add($tbServer)

    $sp.Children.Add((New-DlgLabel "Usuario (DOMINIO\usuario o user@dominio):"))
    $tbUser = New-DlgTextBox; $sp.Children.Add($tbUser)

    $sp.Children.Add((New-DlgLabel "Contraseña:"))
    $tbPwd = New-DlgPasswordBox; $sp.Children.Add($tbPwd)

    $cbNoCred = New-Object System.Windows.Controls.CheckBox
    $cbNoCred.Content = "Usar credenciales actuales (sin usuario/contraseña)"
    $cbNoCred.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
    $cbNoCred.Margin = '0,8,0,0'
    $cbNoCred.IsThreeState = $false
    $cbNoCred.IsChecked = $true
    $sp.Children.Add($cbNoCred)

    $btnConn = New-DlgButton "Conectar" 'Success'; $sp.Children.Add($btnConn)
    $scroll = New-Object System.Windows.Controls.ScrollViewer; $scroll.Content = $sp
    $dlg.Content = $scroll

    $btnConn.Add_Click({
        try {
          $serverInput = Get-RequiredTrimmedText $tbServer.Text "Ingresá un servidor."
          if ($null -eq $serverInput) { return }

          $useCurrentCred = Get-Checked $cbNoCred
          $cred = $null

          if (-not $useCurrentCred) {
            $u = Get-RequiredTrimmedText $tbUser.Text "Ingresá usuario y contraseña, o marcá 'Usar credenciales actuales'."
            if ($null -eq $u -or [string]::IsNullOrWhiteSpace($tbPwd.Password)) {
              Show-Error "Ingresá usuario y contraseña, o marcá 'Usar credenciales actuales'."
              return
            }
            $secPwd = ConvertTo-SecureString $tbPwd.Password -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($u, $secPwd)
          }

          Set-Status "Probando conexión..." '#89b4fa'

          $resolvedServer = Test-ADConnection -server $serverInput -cred $cred
          $domainParams = @{ Server = $resolvedServer }
          if ($cred) { $domainParams.Credential = $cred }
          $connectedDomain = ''
          try { $connectedDomain = [string](Get-ADDomain @domainParams -ErrorAction Stop).DNSRoot } catch {}
          if ([string]::IsNullOrWhiteSpace($connectedDomain)) { $connectedDomain = $resolvedServer }

          $label = if ($cred) { "$connectedDomain [$resolvedServer] ($($cred.UserName))" } else { "$connectedDomain [$resolvedServer]" }

          $global:domainConnections[$label] = @{
            Server     = $resolvedServer
            Credential = $cred
            Domain     = $connectedDomain
          }

          Clear-ADDomainCache
          Refresh-DomainCombo
          $cbDomainSelect.SelectedItem = $label
          Save-DomainConnectionsToFile

          Set-Status "✅ Conectado a '$connectedDomain' ($resolvedServer)." '#a6e3a1'
          Write-AuditLog -Action 'Conectar Dominio' -Result 'OK' -Object "$connectedDomain ($resolvedServer)"
          [System.Windows.MessageBox]::Show("Conexión exitosa a '$connectedDomain' usando '$resolvedServer'.", "Conectado", "OK", "Information") | Out-Null
          $dlg.Close()
        }
        catch {
          Show-Exception "Error al conectar a '$($tbServer.Text.Trim())':" $_
        }
      }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
  })

$btnDomainDisconnect.Add_Click({
    $selected = $cbDomainSelect.SelectedItem
    if ($selected -eq $script:LocalDomainLabel) {
      [System.Windows.MessageBox]::Show("No se puede desconectar el dominio local.", "Info", "OK", "Information") | Out-Null
      return
    }
    if ($selected -and $global:domainConnections.ContainsKey($selected)) {
      $global:domainConnections.Remove($selected)
      Clear-ADDomainCache
      Refresh-DomainCombo
      Save-DomainConnectionsToFile
      Set-Status "✅ Desconectado de '$selected'." '#a6e3a1'
      Write-AuditLog -Action 'Desconectar Dominio' -Result 'OK' -Object ([string]$selected)
    }
  })

# ── USUARIOS ─────────────────────────────────────────────────────────────────
$btnUserSearch.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $filter = Get-NormalizedText $txtUserSearch.Text
    $maxResults = 500

    try {
      [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
      Set-Status "Buscando usuarios..." '#89b4fa'
      $adp = Get-ADParams

      if ([string]::IsNullOrWhiteSpace($filter)) {
        $users = Get-ADUser @adp `
          -Filter * `
          -ResultSetSize ($maxResults + 1) `
          -Properties SamAccountName, Name, EmailAddress, Enabled, LockedOut, DistinguishedName |
        Select-Object SamAccountName, Name, EmailAddress, Enabled, LockedOut,
        @{N = 'OUPath'; E = { Get-OUFromDN $_.DistinguishedName } },
        DistinguishedName
      }
      else {
        # Escapar comillas simples para que el filtro AD no falle (ej: O'Connor)
        $safeFilter = Escape-ADFilterLiteral $filter
        $users = Get-ADUser @adp `
          -Filter "Name -like '*$safeFilter*' -or SamAccountName -like '*$safeFilter*' -or EmailAddress -like '*$safeFilter*'" `
          -ResultSetSize ($maxResults + 1) `
          -Properties SamAccountName, Name, EmailAddress, Enabled, LockedOut, DistinguishedName |
        Select-Object SamAccountName, Name, EmailAddress, Enabled, LockedOut,
        @{N = 'OUPath'; E = { Get-OUFromDN $_.DistinguishedName } },
        DistinguishedName
      }

      $totalFound = @($users).Count
      if ($totalFound -gt $maxResults) {
        $users = @($users | Select-Object -First $maxResults)
      }

      $dgUsers.ItemsSource = @($users)
      if ($totalFound -gt $maxResults) {
        Set-Status "⚠️ Se muestran los primeros $maxResults usuarios. Refiná la búsqueda para acotar resultados." '#fab387'
      }
      else {
        Set-Status "✅ Se encontraron $($users.Count) usuario(s)." '#a6e3a1'
      }
    }
    catch [Microsoft.ActiveDirectory.Management.ADException] {
      Show-Error "No se pudo interpretar el filtro de búsqueda. Probá con texto simple (nombre, usuario o email)."
    }
    catch { Show-Exception "Error buscando usuarios:" $_ }
    finally { [System.Windows.Input.Mouse]::OverrideCursor = $null }
  })
Bind-EnterKeyToButton -textBox $txtUserSearch -button $btnUserSearch
$btnUserRefresh.Add_Click({ Raise-ButtonClick $btnUserSearch })

$dgUsers.Add_MouseDoubleClick({
    if ($dgUsers.SelectedItem) {
      $btnUserEdit.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
  })

# Crear usuario
$btnUserNew.Add_Click({
    if (-not (Assert-ADModule)) { return }
    if (-not (Assert-UserAdminPermission -Operation 'Create')) { return }

    $dlg = New-DialogWindow "Crear Nuevo Usuario" 620 780
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'

    $lblFirstName = New-DlgLabel "Nombre:"; $sp.Children.Add($lblFirstName)
    $tbFirstName = New-DlgTextBox; $sp.Children.Add($tbFirstName)
    $lblInitials = New-DlgLabel "Iniciales:"; $sp.Children.Add($lblInitials)
    $tbInitials = New-DlgTextBox; $sp.Children.Add($tbInitials)
    $lblLastName = New-DlgLabel "Apellido:"; $sp.Children.Add($lblLastName)
    $tbLastName = New-DlgTextBox; $sp.Children.Add($tbLastName)
    $lblFullName = New-DlgLabel "Nombre completo (CN):"; $sp.Children.Add($lblFullName)
    $tbFullName = New-DlgTextBox; $sp.Children.Add($tbFullName)
    $lblDisplayName = New-DlgLabel "Nombre para mostrar:"; $sp.Children.Add($lblDisplayName)
    $tbDisplayName = New-DlgTextBox; $sp.Children.Add($tbDisplayName)
    $lblDescription = New-DlgLabel "Descripción:"; $sp.Children.Add($lblDescription)
    $tbDescription = New-DlgTextBox; $sp.Children.Add($tbDescription)
    $lblOffice = New-DlgLabel "Oficina:"; $sp.Children.Add($lblOffice)
    $tbOffice = New-DlgTextBox; $sp.Children.Add($tbOffice)
    $lblOfficePhone = New-DlgLabel "Teléfono:"; $sp.Children.Add($lblOfficePhone)
    $tbOfficePhone = New-DlgTextBox; $sp.Children.Add($tbOfficePhone)
    $lblEmail = New-DlgLabel "Email:"; $sp.Children.Add($lblEmail)
    $tbEmail = New-DlgTextBox; $sp.Children.Add($tbEmail)
    $lblWebPage = New-DlgLabel "Página Web:"; $sp.Children.Add($lblWebPage)
    $tbWebPage = New-DlgTextBox; $sp.Children.Add($tbWebPage)

    $lblStreetAddress = New-DlgLabel "Calle:"; $sp.Children.Add($lblStreetAddress)
    $tbStreetAddress = New-DlgTextBox
    $tbStreetAddress.AcceptsReturn = $true
    $tbStreetAddress.Height = 60
    $tbStreetAddress.TextWrapping = 'Wrap'
    $tbStreetAddress.VerticalScrollBarVisibility = 'Auto'
    $sp.Children.Add($tbStreetAddress)
    $lblPOBox = New-DlgLabel "Casilla Postal:"; $sp.Children.Add($lblPOBox)
    $tbPOBox = New-DlgTextBox; $sp.Children.Add($tbPOBox)
    $lblCity = New-DlgLabel "Ciudad:"; $sp.Children.Add($lblCity)
    $tbCity = New-DlgTextBox; $sp.Children.Add($tbCity)
    $lblState = New-DlgLabel "Provincia/Estado:"; $sp.Children.Add($lblState)
    $tbState = New-DlgTextBox; $sp.Children.Add($tbState)
    $lblPostalCode = New-DlgLabel "Código Postal:"; $sp.Children.Add($lblPostalCode)
    $tbPostalCode = New-DlgTextBox; $sp.Children.Add($tbPostalCode)
    $lblCountry = New-DlgLabel "País/Región:"; $sp.Children.Add($lblCountry)
    $tbCountry = New-DlgTextBox; $sp.Children.Add($tbCountry)

    $lblUpn = New-DlgLabel "Nombre de inicio de sesión (UPN):"; $sp.Children.Add($lblUpn)
    $upnGrid = New-Object System.Windows.Controls.Grid
    $upnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
    $upnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = 'Auto' }))
    $tbUpnUser = New-DlgTextBox
    $tbUpnUser.Margin = '0,0,6,0'
    [System.Windows.Controls.Grid]::SetColumn($tbUpnUser, 0)
    $lblUpnSuffix = New-Object System.Windows.Controls.TextBlock
    $lblUpnSuffix.VerticalAlignment = 'Center'
    $lblUpnSuffix.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextSecondary
    [System.Windows.Controls.Grid]::SetColumn($lblUpnSuffix, 1)
    $upnGrid.Children.Add($tbUpnUser) | Out-Null
    $upnGrid.Children.Add($lblUpnSuffix) | Out-Null
    $sp.Children.Add($upnGrid)

    $lblSam = New-DlgLabel "Nombre de inicio de sesión (anterior a Windows 2000):"; $sp.Children.Add($lblSam)
    $tbSam = New-DlgTextBox
    $sp.Children.Add($tbSam)
    $lblPwd = New-DlgLabel "Contraseña:"; $sp.Children.Add($lblPwd)
    $tbPwd = New-DlgPasswordBox; $sp.Children.Add($tbPwd)
    $lblPwd2 = New-DlgLabel "Repetir contraseña:"; $sp.Children.Add($lblPwd2)
    $tbPwdConfirm = New-DlgPasswordBox; $sp.Children.Add($tbPwdConfirm)

    $lblOU = New-DlgLabel "Unidad Organizativa (OU):"; $sp.Children.Add($lblOU)
    $cbOU = New-DlgComboBox @() 0
    $cbOU.MinWidth = 480
    $cbOU.MaxDropDownHeight = 320
    $cbOU.DisplayMemberPath = 'Label'
    $cbOU.SelectedValuePath = 'DistinguishedName'
    $cbOU.ToolTip = "Seleccioná la OU o contenedor destino."
    $sp.Children.Add($cbOU)

    try {
      $adpForOu = Get-ADParams
      $domainInfo = Get-ADDomainCached -adParams $adpForOu
      $usersContainerDn = $domainInfo.UsersContainer

      $ouDns = Get-ADOrganizationalUnit @adpForOu -Filter * -Properties DistinguishedName |
      Sort-Object DistinguishedName |
      Select-Object -ExpandProperty DistinguishedName

      if ((-not [string]::IsNullOrWhiteSpace($usersContainerDn)) -and (-not ($ouDns -contains $usersContainerDn))) {
        $ouDns += $usersContainerDn
      }

      # Seguridad: excluir Domain Controllers del desplegable para evitar errores humanos.
      $ouDns = @(
        $ouDns |
        Where-Object { $_ -and ($_ -notmatch '^(?i)(OU|CN)=Domain Controllers,') } |
        Sort-Object -Unique
      )

      $cbOU.Items.Clear() | Out-Null

      foreach ($ouDn in @($ouDns)) {
        if (-not [string]::IsNullOrWhiteSpace($ouDn)) {
          $label = Get-ContainerLabelFromDN $ouDn
          $cbOU.Items.Add([PSCustomObject]@{
              Label             = $label
              DistinguishedName = $ouDn
            }) | Out-Null
        }
      }
      if ($cbOU.Items.Count -gt 0) { $cbOU.SelectedIndex = 0 }
      else { Show-Error "No hay OUs disponibles para crear usuarios (se excluye 'Domain Controllers')." }
    }
    catch {
      Show-Exception "Error cargando OUs para crear usuario:" $_
    }

    # Similar a ADUC: completar campos de identidad automáticamente, pero editables.
    $updateIdentity = {
      $fn = Get-NormalizedText $tbFirstName.Text
      $ln = Get-NormalizedText $tbLastName.Text

      $nameParts = @()
      if (-not [string]::IsNullOrWhiteSpace($fn)) { $nameParts += $fn }
      if (-not [string]::IsNullOrWhiteSpace($ln)) { $nameParts += $ln }
      $full = ($nameParts -join ' ').Trim()
      if (-not [string]::IsNullOrWhiteSpace($full)) {
        $tbFullName.Text = $full
        $tbDisplayName.Text = $full
      }

      if ((-not [string]::IsNullOrWhiteSpace($fn)) -and (-not [string]::IsNullOrWhiteSpace($ln))) {
        $samAuto = (($fn.Substring(0, 1) + $ln).ToLower() -replace '[^a-z0-9._-]', '')
        $tbSam.Text = $samAuto
        $tbUpnUser.Text = $samAuto
      }
      elseif (-not [string]::IsNullOrWhiteSpace($ln)) {
        $samAuto = ($ln.ToLower() -replace '[^a-z0-9._-]', '')
        $tbSam.Text = $samAuto
        $tbUpnUser.Text = $samAuto
      }
    }
    $tbFirstName.Add_TextChanged($updateIdentity)
    $tbLastName.Add_TextChanged($updateIdentity)

    $adpForDomain = Get-ADParams
    $domainRootForUpn = (Get-ADDomainCached -adParams $adpForDomain).DNSRoot
    $lblUpnSuffix.Text = "@$domainRootForUpn"

    $cbMustChange = New-Object System.Windows.Controls.CheckBox
    $cbMustChange.Content = "El usuario debe cambiar la contraseña en el próximo inicio"
    $cbMustChange.IsChecked = $true
    $cbMustChange.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
    $cbMustChange.Margin = '0,10,0,0'
    $cbMustChange.IsThreeState = $false
    $sp.Children.Add($cbMustChange)

    $cbCannotChange = New-Object System.Windows.Controls.CheckBox
    $cbCannotChange.Content = "El usuario no puede cambiar la contraseña"
    $cbCannotChange.IsChecked = $false
    $cbCannotChange.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
    $cbCannotChange.Margin = '0,2,0,0'
    $cbCannotChange.IsThreeState = $false
    $sp.Children.Add($cbCannotChange)

    $cbPwdNeverExpires = New-Object System.Windows.Controls.CheckBox
    $cbPwdNeverExpires.Content = "La contraseña nunca expira"
    $cbPwdNeverExpires.IsChecked = $false
    $cbPwdNeverExpires.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
    $cbPwdNeverExpires.Margin = '0,2,0,0'
    $cbPwdNeverExpires.IsThreeState = $false
    $sp.Children.Add($cbPwdNeverExpires)

    $cbEnabled = New-Object System.Windows.Controls.CheckBox
    $cbEnabled.Content = "Cuenta habilitada"
    $cbEnabled.IsChecked = $true
    $cbEnabled.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
    $cbEnabled.Margin = '0,2,0,0'
    $cbEnabled.IsThreeState = $false
    $sp.Children.Add($cbEnabled)

    $modeUi = 'Create'
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'FirstName' -LabelControl $lblFirstName -InputControl $tbFirstName
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Initials' -LabelControl $lblInitials -InputControl $tbInitials
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'LastName' -LabelControl $lblLastName -InputControl $tbLastName
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'FullName' -LabelControl $lblFullName -InputControl $tbFullName
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'DisplayName' -LabelControl $lblDisplayName -InputControl $tbDisplayName
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Description' -LabelControl $lblDescription -InputControl $tbDescription
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Office' -LabelControl $lblOffice -InputControl $tbOffice
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'OfficePhone' -LabelControl $lblOfficePhone -InputControl $tbOfficePhone
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Email' -LabelControl $lblEmail -InputControl $tbEmail
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'WebPage' -LabelControl $lblWebPage -InputControl $tbWebPage
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'StreetAddress' -LabelControl $lblStreetAddress -InputControl $tbStreetAddress
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'POBox' -LabelControl $lblPOBox -InputControl $tbPOBox
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'City' -LabelControl $lblCity -InputControl $tbCity
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'State' -LabelControl $lblState -InputControl $tbState
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'PostalCode' -LabelControl $lblPostalCode -InputControl $tbPostalCode
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Country' -LabelControl $lblCountry -InputControl $tbCountry
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'UpnUser' -LabelControl $lblUpn -InputControl $upnGrid
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'SamAccountName' -LabelControl $lblSam -InputControl $tbSam
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Password' -LabelControl $lblPwd -InputControl $tbPwd
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Password' -LabelControl $lblPwd2 -InputControl $tbPwdConfirm
    Apply-UserFieldUiPolicy -Mode $modeUi -Field 'OU' -LabelControl $lblOU -InputControl $cbOU

    $cbMustChange.Visibility = if (Test-UserFieldVisible -Mode $modeUi -Field 'MustChangePassword') { 'Visible' } else { 'Collapsed' }
    $cbMustChange.IsEnabled = Test-UserFieldEditable -Mode $modeUi -Field 'MustChangePassword'
    $cbCannotChange.Visibility = if (Test-UserFieldVisible -Mode $modeUi -Field 'CannotChangePassword') { 'Visible' } else { 'Collapsed' }
    $cbCannotChange.IsEnabled = Test-UserFieldEditable -Mode $modeUi -Field 'CannotChangePassword'
    $cbPwdNeverExpires.Visibility = if (Test-UserFieldVisible -Mode $modeUi -Field 'PasswordNeverExpires') { 'Visible' } else { 'Collapsed' }
    $cbPwdNeverExpires.IsEnabled = Test-UserFieldEditable -Mode $modeUi -Field 'PasswordNeverExpires'
    $cbEnabled.Visibility = if (Test-UserFieldVisible -Mode $modeUi -Field 'Enabled') { 'Visible' } else { 'Collapsed' }
    $cbEnabled.IsEnabled = Test-UserFieldEditable -Mode $modeUi -Field 'Enabled'

    $btnCreate = New-DlgButton "Crear Usuario" 'Success'
    $sp.Children.Add($btnCreate)

    $scroll = New-Object System.Windows.Controls.ScrollViewer
    $scroll.Content = $sp
    $dlg.Content = $scroll

    $btnCreate.Add_Click({
        try {
          $mode = 'Create'
          $fn = Get-NormalizedText $tbFirstName.Text
          $ln = Get-NormalizedText $tbLastName.Text
          $cnName = Get-NormalizedText $tbFullName.Text
          $displayName = Get-NormalizedText $tbDisplayName.Text
          $upnUser = Get-NormalizedText $tbUpnUser.Text
          $sam = Get-NormalizedText $tbSam.Text
          $selectedOU = [string]$cbOU.SelectedValue
          $pwdPlain = $tbPwd.Password
          $pwdConfirm = $tbPwdConfirm.Password

          $effectiveSam = if (Test-UserFieldEditable -Mode $mode -Field 'SamAccountName') { $sam } else { New-DefaultSam -firstName $fn -lastName $ln }
          if ([string]::IsNullOrWhiteSpace($effectiveSam)) { $effectiveSam = New-DefaultSam -firstName $fn -lastName $ln }

          $effectiveCn = if (Test-UserFieldEditable -Mode $mode -Field 'FullName') { $cnName } else { '' }
          if ([string]::IsNullOrWhiteSpace($effectiveCn)) {
            $effectiveCn = ((@($fn, $ln) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' ').Trim()
          }
          if ([string]::IsNullOrWhiteSpace($effectiveCn)) { $effectiveCn = $effectiveSam }

          $effectiveUpnUser = if (Test-UserFieldEditable -Mode $mode -Field 'UpnUser') { $upnUser } else { '' }
          if ([string]::IsNullOrWhiteSpace($effectiveUpnUser)) { $effectiveUpnUser = $effectiveSam }

          if (-not (Test-UserFieldValue -Mode $mode -Field 'FirstName' -Value $fn -ErrorMessage "El campo 'Nombre' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'LastName' -Value $ln -ErrorMessage "El campo 'Apellido' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'FullName' -Value $effectiveCn -ErrorMessage "El campo 'Nombre completo (CN)' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'UpnUser' -Value $effectiveUpnUser -ErrorMessage "El campo 'UPN' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'SamAccountName' -Value $effectiveSam -ErrorMessage "El campo 'SamAccountName' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'OU' -Value $selectedOU -ErrorMessage "Seleccioná una OU o contenedor destino (requerido por política).")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Initials' -Value (Get-NormalizedText $tbInitials.Text) -ErrorMessage "El campo 'Iniciales' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Description' -Value (Get-NormalizedText $tbDescription.Text) -ErrorMessage "El campo 'Descripción' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Office' -Value (Get-NormalizedText $tbOffice.Text) -ErrorMessage "El campo 'Oficina' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'OfficePhone' -Value (Get-NormalizedText $tbOfficePhone.Text) -ErrorMessage "El campo 'Teléfono' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Email' -Value (Get-NormalizedText $tbEmail.Text) -ErrorMessage "El campo 'Email' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'WebPage' -Value (Get-NormalizedText $tbWebPage.Text) -ErrorMessage "El campo 'Página Web' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'StreetAddress' -Value (Get-NormalizedText $tbStreetAddress.Text) -ErrorMessage "El campo 'Calle' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'POBox' -Value (Get-NormalizedText $tbPOBox.Text) -ErrorMessage "El campo 'Casilla Postal' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'City' -Value (Get-NormalizedText $tbCity.Text) -ErrorMessage "El campo 'Ciudad' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'State' -Value (Get-NormalizedText $tbState.Text) -ErrorMessage "El campo 'Provincia/Estado' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'PostalCode' -Value (Get-NormalizedText $tbPostalCode.Text) -ErrorMessage "El campo 'Código Postal' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Country' -Value (Get-NormalizedText $tbCountry.Text) -ErrorMessage "El campo 'País/Región' es requerido por política.")) { return }
          $canSetPassword = Test-UserFieldEditable -Mode $mode -Field 'Password'
          if (($canSetPassword) -and ((Test-UserFieldRequired -Mode $mode -Field 'Password') -or (-not [string]::IsNullOrWhiteSpace($pwdPlain)) -or (-not [string]::IsNullOrWhiteSpace($pwdConfirm)))) {
            if ([string]::IsNullOrWhiteSpace($pwdPlain) -or [string]::IsNullOrWhiteSpace($pwdConfirm)) {
              Show-Error "El campo contraseña es requerido por política."
              return
            }
          }
          if ($canSetPassword -and ((-not [string]::IsNullOrWhiteSpace($pwdPlain)) -or (-not [string]::IsNullOrWhiteSpace($pwdConfirm))) -and ($pwdPlain -ne $pwdConfirm)) {
            Show-Error "Las contraseñas no coinciden."
            return
          }
          if ((Test-UserFieldEditable -Mode $mode -Field 'MustChangePassword') -and (Test-UserFieldEditable -Mode $mode -Field 'CannotChangePassword') -and (Get-Checked $cbMustChange) -and (Get-Checked $cbCannotChange)) {
            Show-Error "No podés marcar a la vez 'debe cambiar contraseña' y 'no puede cambiar contraseña'."
            return
          }

          $adp = Get-ADParams
          $domainRoot = (Get-ADDomainCached -adParams $adp).DNSRoot

          $upn = "$effectiveUpnUser@$domainRoot"
          $safeSam = Escape-ADFilterLiteral $effectiveSam
          $safeUpn = Escape-ADFilterLiteral $upn
          $safeCn = Escape-ADFilterLiteral $effectiveCn

          $existingSam = Get-ADUser @adp -Filter "SamAccountName -eq '$safeSam'" -ResultSetSize 1 -ErrorAction SilentlyContinue
          if ($existingSam) {
            Show-Error "El SAM '$effectiveSam' ya existe. Corregí el dato y volvé a intentar."
            return
          }

          $existingUpn = Get-ADUser @adp -Filter "UserPrincipalName -eq '$safeUpn'" -ResultSetSize 1 -ErrorAction SilentlyContinue
          if ($existingUpn) {
            Show-Error "El UPN '$upn' ya existe. Corregí el dato y volvé a intentar."
            return
          }

          $params = @{
            Name              = $effectiveCn
            SamAccountName    = $effectiveSam
            UserPrincipalName = $upn
            ErrorAction       = 'Stop'
          }

          if ((Test-UserFieldEditable -Mode $mode -Field 'FirstName') -and (-not [string]::IsNullOrWhiteSpace($fn))) { $params.GivenName = $fn }
          if ((Test-UserFieldEditable -Mode $mode -Field 'LastName') -and (-not [string]::IsNullOrWhiteSpace($ln))) { $params.Surname = $ln }
          if ((Test-UserFieldEditable -Mode $mode -Field 'DisplayName') -and (-not [string]::IsNullOrWhiteSpace($displayName))) { $params.DisplayName = $displayName }
          if ((Test-UserFieldEditable -Mode $mode -Field 'Initials') -and (-not [string]::IsNullOrWhiteSpace($tbInitials.Text))) { $params.Initials = $tbInitials.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'Description') -and (-not [string]::IsNullOrWhiteSpace($tbDescription.Text))) { $params.Description = $tbDescription.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'Email') -and (-not [string]::IsNullOrWhiteSpace($tbEmail.Text))) { $params.EmailAddress = $tbEmail.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'Office') -and (-not [string]::IsNullOrWhiteSpace($tbOffice.Text))) { $params.Office = $tbOffice.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'OfficePhone') -and (-not [string]::IsNullOrWhiteSpace($tbOfficePhone.Text))) { $params.OfficePhone = $tbOfficePhone.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'WebPage') -and (-not [string]::IsNullOrWhiteSpace($tbWebPage.Text))) { $params.HomePage = $tbWebPage.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'StreetAddress') -and (-not [string]::IsNullOrWhiteSpace($tbStreetAddress.Text))) { $params.StreetAddress = $tbStreetAddress.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'POBox') -and (-not [string]::IsNullOrWhiteSpace($tbPOBox.Text))) { $params.POBox = $tbPOBox.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'City') -and (-not [string]::IsNullOrWhiteSpace($tbCity.Text))) { $params.City = $tbCity.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'State') -and (-not [string]::IsNullOrWhiteSpace($tbState.Text))) { $params.State = $tbState.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'PostalCode') -and (-not [string]::IsNullOrWhiteSpace($tbPostalCode.Text))) { $params.PostalCode = $tbPostalCode.Text.Trim() }
          if ((Test-UserFieldEditable -Mode $mode -Field 'Country') -and (-not [string]::IsNullOrWhiteSpace($tbCountry.Text))) { $params.Country = $tbCountry.Text.Trim() }

          $hasPassword = $canSetPassword -and (-not [string]::IsNullOrWhiteSpace($pwdPlain))
          if ($hasPassword) {
            $params.AccountPassword = (ConvertTo-SecureString $pwdPlain -AsPlainText -Force)
          }

          if (Test-UserFieldEditable -Mode $mode -Field 'Enabled') {
            $params.Enabled = if ($hasPassword) { (Get-Checked $cbEnabled) } else { $false }
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'MustChangePassword') { $params.ChangePasswordAtLogon = (Get-Checked $cbMustChange) }
          if (Test-UserFieldEditable -Mode $mode -Field 'CannotChangePassword') { $params.CannotChangePassword = (Get-Checked $cbCannotChange) }
          if (Test-UserFieldEditable -Mode $mode -Field 'PasswordNeverExpires') { $params.PasswordNeverExpires = (Get-Checked $cbPwdNeverExpires) }

          if ($selectedOU -and ($selectedOU -match '^(?i)(OU|CN)=Domain Controllers,')) {
            Show-Error "Por seguridad no se permite crear usuarios en 'Domain Controllers'. Seleccioná otra OU."
            return
          }
          if (-not [string]::IsNullOrWhiteSpace($selectedOU)) {
            $existingName = Get-ADObject @adp -SearchBase $selectedOU -SearchScope OneLevel -Filter "Name -eq '$safeCn'" -ResultSetSize 1 -ErrorAction SilentlyContinue
            if ($existingName) {
              Show-Error "Ya existe un objeto con nombre '$effectiveCn' en el contenedor seleccionado. Corregí el nombre y volvé a intentar."
              return
            }
            if (Test-UserFieldEditable -Mode $mode -Field 'OU') { $params.Path = $selectedOU }
          }

          $newUserParams = Merge-ADParams -params $params -adParams $adp
          New-ADUser @newUserParams

          [System.Windows.MessageBox]::Show("Usuario '$effectiveCn' creado correctamente.", "Éxito", "OK", "Information") | Out-Null
          Set-Status "✅ Usuario '$effectiveCn' creado." '#a6e3a1'
          Write-AuditLog -Action 'Crear Usuario' -Result 'OK' -Object $effectiveSam
          $dlg.Close()
        }
        catch {
          $msg = $_.Exception.Message
          Write-AuditLog -Action 'Crear Usuario' -Result 'ERROR' -Object $effectiveSam -Detail $msg
          Show-Error "No se pudo crear el usuario. Corregí los datos y volvé a intentar.`n`nDetalle: $msg"
        }
      }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
  })

# Copiar perfil de usuario (similar a ADUC)
$btnUserCopyProfile.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $sel = Get-SelectedUser; if (-not $sel) { return }

    try {
      $adp = Get-ADParams
      $src = Get-ADUser @adp -Identity $sel.SamAccountName -Properties GivenName, Surname, EmailAddress, Company, Office, DistinguishedName, Name, SamAccountName

      $dlg = New-DialogWindow "Copiar Perfil: $($src.SamAccountName)" 480 620
      $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'

      $sp.Children.Add((New-DlgLabel "Usuario origen: $($src.SamAccountName)"))

      $sp.Children.Add((New-DlgLabel "Nombre:")); $tbFirstName = New-DlgTextBox $src.GivenName; $sp.Children.Add($tbFirstName)
      $sp.Children.Add((New-DlgLabel "Apellido:")); $tbLastName = New-DlgTextBox $src.Surname; $sp.Children.Add($tbLastName)

      $sp.Children.Add((New-DlgLabel "SamAccountName (nuevo):"))
      $tbSam = New-DlgTextBox
      $sp.Children.Add($tbSam)

      $updateSam = {
        $f = Get-NormalizedText $tbFirstName.Text
        $l = Get-NormalizedText $tbLastName.Text
        if ((-not [string]::IsNullOrWhiteSpace($f)) -and (-not [string]::IsNullOrWhiteSpace($l))) {
          $tbSam.Text = ($f.Substring(0, 1) + $l).ToLower() -replace '\s', ''
        }
        else {
          $tbSam.Text = ''
        }
      }
      $tbFirstName.Add_TextChanged($updateSam)
      $tbLastName.Add_TextChanged($updateSam)
      & $updateSam

      $sp.Children.Add((New-DlgLabel "Email:")); $tbEmail = New-DlgTextBox $src.EmailAddress; $sp.Children.Add($tbEmail)
      $sp.Children.Add((New-DlgLabel "Empresa:")); $tbCompany = New-DlgTextBox $src.Company; $sp.Children.Add($tbCompany)
      $sp.Children.Add((New-DlgLabel "Oficina:")); $tbOffice = New-DlgTextBox $src.Office; $sp.Children.Add($tbOffice)
      $sp.Children.Add((New-DlgLabel "Contraseña inicial:")); $tbPwd = New-DlgPasswordBox; $sp.Children.Add($tbPwd)
      $sp.Children.Add((New-DlgLabel "Repetir contraseña inicial:")); $tbPwdConfirm = New-DlgPasswordBox; $sp.Children.Add($tbPwdConfirm)

      $sp.Children.Add((New-DlgLabel "Unidad Organizativa (OU):"))
      $cbOU = New-DlgComboBox @() 0
      $cbOU.MinWidth = 380
      $cbOU.MaxDropDownHeight = 320
      $cbOU.DisplayMemberPath = 'Label'
      $cbOU.SelectedValuePath = 'DistinguishedName'
      $sp.Children.Add($cbOU)

      $srcOuDn = Get-OUFromDN $src.DistinguishedName
      try {
        $domainInfo = Get-ADDomainCached -adParams $adp
        $usersContainerDn = $domainInfo.UsersContainer
        $ouDns = Get-ADOrganizationalUnit @adp -Filter * -Properties DistinguishedName |
        Sort-Object DistinguishedName |
        Select-Object -ExpandProperty DistinguishedName

        if ((-not [string]::IsNullOrWhiteSpace($usersContainerDn)) -and (-not ($ouDns -contains $usersContainerDn))) {
          $ouDns += $usersContainerDn
        }

        $ouDns = @(
          $ouDns |
          Where-Object { $_ -and ($_ -notmatch '^(?i)(OU|CN)=Domain Controllers,') } |
          Sort-Object -Unique
        )

        $cbOU.Items.Clear() | Out-Null
        foreach ($ouDn in @($ouDns)) {
          if (-not [string]::IsNullOrWhiteSpace($ouDn)) {
            $cbOU.Items.Add([PSCustomObject]@{
                Label             = (Get-ContainerLabelFromDN $ouDn)
                DistinguishedName = $ouDn
              }) | Out-Null
          }
        }
        if ($cbOU.Items.Count -gt 0) {
          $prefIdx = -1
          for ($i = 0; $i -lt $cbOU.Items.Count; $i++) {
            if ([string]$cbOU.Items[$i].DistinguishedName -eq [string]$srcOuDn) { $prefIdx = $i; break }
          }
          if ($prefIdx -ge 0) { $cbOU.SelectedIndex = $prefIdx } else { $cbOU.SelectedIndex = 0 }
        }
      }
      catch {
        Show-Exception "Error cargando OUs para copiar perfil:" $_
      }

      $cbEnabled = New-Object System.Windows.Controls.CheckBox
      $cbEnabled.Content = "Cuenta habilitada"
      $cbEnabled.IsChecked = $true
      $cbEnabled.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
      $cbEnabled.Margin = '0,10,0,0'
      $sp.Children.Add($cbEnabled)

      $btnCreate = New-DlgButton "Crear Copia de Perfil" 'Success'
      $sp.Children.Add($btnCreate)

      $scroll = New-Object System.Windows.Controls.ScrollViewer
      $scroll.Content = $sp
      $dlg.Content = $scroll

      $btnCreate.Add_Click({
          try {
            $fn = Get-NormalizedText $tbFirstName.Text
            $ln = Get-NormalizedText $tbLastName.Text
            $sam = Get-NormalizedText $tbSam.Text
            $pwdPlain = $tbPwd.Password
            $pwdConfirm = $tbPwdConfirm.Password

            if ([string]::IsNullOrWhiteSpace($fn) -or [string]::IsNullOrWhiteSpace($ln) -or [string]::IsNullOrWhiteSpace($sam) -or [string]::IsNullOrWhiteSpace($pwdPlain) -or [string]::IsNullOrWhiteSpace($pwdConfirm)) {
              Show-Error "Completá Nombre, Apellido, SamAccountName y ambas contraseñas."
              return
            }
            if ($pwdPlain -ne $pwdConfirm) {
              Show-Error "Las contraseñas no coinciden."
              return
            }
            if ($sam -eq $src.SamAccountName) {
              Show-Error "El SamAccountName del nuevo usuario debe ser distinto al usuario origen."
              return
            }
            if (($fn -eq (Get-NormalizedText $src.GivenName)) -and ($ln -eq (Get-NormalizedText $src.Surname))) {
              Show-Error "Debés cambiar el nombre/apellido respecto del usuario origen."
              return
            }

            $attrChanged = $false
            foreach ($pair in @(
                @{ New = (Get-NormalizedText $tbEmail.Text); Source = (Get-NormalizedText $src.EmailAddress) },
                @{ New = (Get-NormalizedText $tbCompany.Text); Source = (Get-NormalizedText $src.Company) },
                @{ New = (Get-NormalizedText $tbOffice.Text); Source = (Get-NormalizedText $src.Office) }
              )) {
              if ($pair.New -ne $pair.Source) { $attrChanged = $true; break }
            }
            if (-not $attrChanged) {
              Show-Error "Debés cambiar al menos un atributo (email, empresa u oficina)."
              return
            }

            $selectedOU = [string]$cbOU.SelectedValue
            if ([string]::IsNullOrWhiteSpace($selectedOU)) { Show-Error "Seleccioná una OU o contenedor destino."; return }
            if ($selectedOU -match '^(?i)(OU|CN)=Domain Controllers,') { Show-Error "Por seguridad no se permite crear usuarios en 'Domain Controllers'."; return }

            $adp2 = Get-ADParams
            $domainRoot = (Get-ADDomainCached -adParams $adp2).DNSRoot
            $name = "$fn $ln"

            $newParams = @{
              Name              = $name
              GivenName         = $fn
              Surname           = $ln
              SamAccountName    = $sam
              UserPrincipalName = "$sam@$domainRoot"
              AccountPassword   = (ConvertTo-SecureString $pwdPlain -AsPlainText -Force)
              Enabled           = (Get-Checked $cbEnabled)
              Path              = $selectedOU
              EmailAddress      = (Get-NormalizedText $tbEmail.Text)
              Company           = (Get-NormalizedText $tbCompany.Text)
              Office            = (Get-NormalizedText $tbOffice.Text)
            }

            $createParams = @{}
            foreach ($k in $newParams.Keys) {
              if (-not [string]::IsNullOrWhiteSpace([string]$newParams[$k])) { $createParams[$k] = $newParams[$k] }
            }

            $createParams = Merge-ADParams -params $createParams -adParams $adp2
            New-ADUser @createParams

            $newUser = Get-ADUser @adp2 -Identity $sam -Properties DistinguishedName, SamAccountName

            $groupErrors = @()
            $sourceGroups = Get-ADPrincipalGroupMembership @adp2 -Identity $src.DistinguishedName
            foreach ($g in @($sourceGroups)) {
              if ($null -eq $g) { continue }
              $sidValue = ''
              if ($g.SID) { $sidValue = $g.SID.Value }
              if ($sidValue -match '-513$') { continue } # Domain Users (grupo primario habitual)
              try {
                Add-ADGroupMember @adp2 -Identity $g.DistinguishedName -Members $newUser.DistinguishedName -ErrorAction Stop
              }
              catch {
                $groupErrors += "$($g.Name): $($_.Exception.Message)"
              }
            }

            if ($groupErrors.Count -gt 0) {
              $preview = ($groupErrors | Select-Object -First 10) -join "`n - "
              [System.Windows.MessageBox]::Show(
                "Usuario creado, pero hubo errores copiando algunos grupos:`n - $preview",
                "Copia de perfil con advertencias",
                "OK",
                "Warning"
              ) | Out-Null
              Set-Status "⚠️ Usuario '$sam' creado. Revisá grupos con errores." '#fab387'
              Write-AuditLog -Action 'Copiar Perfil Usuario' -Result 'WARN' -Object $sam -Detail ("Errores de grupos: " + ($groupErrors.Count))
            }
            else {
              [System.Windows.MessageBox]::Show("Perfil copiado correctamente para '$sam'.", "Éxito", "OK", "Information") | Out-Null
              Set-Status "✅ Perfil copiado a '$sam' respetando grupos/permisos." '#a6e3a1'
              Write-AuditLog -Action 'Copiar Perfil Usuario' -Result 'OK' -Object $sam
            }

            Raise-ButtonClick $btnUserSearch
            $dlg.Close()
          }
          catch { Show-Exception "Error copiando perfil de usuario:" $_ }
        }.GetNewClosure())

      $dlg.ShowDialog() | Out-Null
    }
    catch { Show-Exception "Error preparando copia de perfil:" $_ }
  })

# Modificar usuario
$btnUserEdit.Add_Click({
    if (-not (Assert-ADModule)) { return }
    if (-not (Assert-UserAdminPermission -Operation 'Modify')) { return }
    $sel = Get-SelectedUser; if (-not $sel) { return }

    try {
      $adp = Get-ADParams
      $user = Get-ADUser @adp -Identity $sel.SamAccountName -Properties GivenName, Initials, Surname, DisplayName, Description, UserPrincipalName, SamAccountName, EmailAddress, Office, OfficePhone, HomePage, StreetAddress, POBox, City, State, PostalCode, Country, DistinguishedName, Name, Enabled, CannotChangePassword, PasswordNeverExpires, pwdLastSet

      $dlg = New-DialogWindow "Modificar Usuario: $($user.SamAccountName)" 560 860
      $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'

      $lblFirstName = New-DlgLabel "Nombre:"; $sp.Children.Add($lblFirstName)
      $tbFirstName = New-DlgTextBox $user.GivenName; $sp.Children.Add($tbFirstName)
      $lblInitials = New-DlgLabel "Iniciales:"; $sp.Children.Add($lblInitials)
      $tbInitials = New-DlgTextBox $user.Initials; $sp.Children.Add($tbInitials)
      $lblLastName = New-DlgLabel "Apellido:"; $sp.Children.Add($lblLastName)
      $tbLastName = New-DlgTextBox $user.Surname; $sp.Children.Add($tbLastName)
      $lblFullName = New-DlgLabel "Nombre completo (CN):"; $sp.Children.Add($lblFullName)
      $tbFullName = New-DlgTextBox $user.Name; $sp.Children.Add($tbFullName)
      $lblDisplayName = New-DlgLabel "Nombre para mostrar:"; $sp.Children.Add($lblDisplayName)
      $tbDisplayName = New-DlgTextBox $user.DisplayName; $sp.Children.Add($tbDisplayName)
      $lblDescription = New-DlgLabel "Descripción:"; $sp.Children.Add($lblDescription)
      $tbDescription = New-DlgTextBox $user.Description; $sp.Children.Add($tbDescription)
      $lblOffice = New-DlgLabel "Oficina:"; $sp.Children.Add($lblOffice)
      $tbOffice = New-DlgTextBox $user.Office; $sp.Children.Add($tbOffice)
      $lblOfficePhone = New-DlgLabel "Teléfono:"; $sp.Children.Add($lblOfficePhone)
      $tbOfficePhone = New-DlgTextBox $user.OfficePhone; $sp.Children.Add($tbOfficePhone)
      $lblEmail = New-DlgLabel "Email:"; $sp.Children.Add($lblEmail)
      $tbEmail = New-DlgTextBox $user.EmailAddress; $sp.Children.Add($tbEmail)
      $lblWebPage = New-DlgLabel "Página Web:"; $sp.Children.Add($lblWebPage)
      $tbWebPage = New-DlgTextBox $user.HomePage; $sp.Children.Add($tbWebPage)

      $lblStreetAddress = New-DlgLabel "Calle:"; $sp.Children.Add($lblStreetAddress)
      $tbStreetAddress = New-DlgTextBox $user.StreetAddress
      $tbStreetAddress.AcceptsReturn = $true
      $tbStreetAddress.Height = 60
      $tbStreetAddress.TextWrapping = 'Wrap'
      $tbStreetAddress.VerticalScrollBarVisibility = 'Auto'
      $sp.Children.Add($tbStreetAddress)
      $lblPOBox = New-DlgLabel "Casilla Postal:"; $sp.Children.Add($lblPOBox)
      $tbPOBox = New-DlgTextBox $user.POBox; $sp.Children.Add($tbPOBox)
      $lblCity = New-DlgLabel "Ciudad:"; $sp.Children.Add($lblCity)
      $tbCity = New-DlgTextBox $user.City; $sp.Children.Add($tbCity)
      $lblState = New-DlgLabel "Provincia/Estado:"; $sp.Children.Add($lblState)
      $tbState = New-DlgTextBox $user.State; $sp.Children.Add($tbState)
      $lblPostalCode = New-DlgLabel "Código Postal:"; $sp.Children.Add($lblPostalCode)
      $tbPostalCode = New-DlgTextBox $user.PostalCode; $sp.Children.Add($tbPostalCode)
      $lblCountry = New-DlgLabel "País/Región:"; $sp.Children.Add($lblCountry)
      $tbCountry = New-DlgTextBox $user.Country; $sp.Children.Add($tbCountry)

      $lblUpn = New-DlgLabel "Nombre de inicio de sesión (UPN):"; $sp.Children.Add($lblUpn)
      $upnGrid = New-Object System.Windows.Controls.Grid
      $upnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
      $upnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = 'Auto' }))
      $upnUserPart = ''
      if ((-not [string]::IsNullOrWhiteSpace($user.UserPrincipalName)) -and ($user.UserPrincipalName -match '^([^@]+)@')) { $upnUserPart = $matches[1] }
      if ([string]::IsNullOrWhiteSpace($upnUserPart)) { $upnUserPart = $user.SamAccountName }
      $tbUpnUser = New-DlgTextBox $upnUserPart
      $tbUpnUser.Margin = '0,0,6,0'
      [System.Windows.Controls.Grid]::SetColumn($tbUpnUser, 0)
      $lblUpnSuffix = New-Object System.Windows.Controls.TextBlock
      $lblUpnSuffix.VerticalAlignment = 'Center'
      $lblUpnSuffix.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextSecondary
      [System.Windows.Controls.Grid]::SetColumn($lblUpnSuffix, 1)
      $upnGrid.Children.Add($tbUpnUser) | Out-Null
      $upnGrid.Children.Add($lblUpnSuffix) | Out-Null
      $sp.Children.Add($upnGrid)

      $lblSam = New-DlgLabel "Nombre de inicio de sesión (anterior a Windows 2000):"; $sp.Children.Add($lblSam)
      $tbSam = New-DlgTextBox $user.SamAccountName
      $sp.Children.Add($tbSam)

      $lblPwd = New-DlgLabel "Nueva contraseña (opcional):"; $sp.Children.Add($lblPwd)
      $tbPwd = New-DlgPasswordBox; $sp.Children.Add($tbPwd)
      $lblPwd2 = New-DlgLabel "Repetir nueva contraseña:"; $sp.Children.Add($lblPwd2)
      $tbPwdConfirm = New-DlgPasswordBox; $sp.Children.Add($tbPwdConfirm)

      $lblOU = New-DlgLabel "Unidad Organizativa (OU):"; $sp.Children.Add($lblOU)
      $cbOU = New-DlgComboBox @() 0
      $cbOU.MinWidth = 420
      $cbOU.MaxDropDownHeight = 320
      $cbOU.DisplayMemberPath = 'Label'
      $cbOU.SelectedValuePath = 'DistinguishedName'
      $cbOU.ToolTip = "Seleccioná la OU o contenedor destino."
      $sp.Children.Add($cbOU)

      try {
        $domainInfo = Get-ADDomainCached -adParams $adp
        $usersContainerDn = $domainInfo.UsersContainer
        $ouDns = Get-ADOrganizationalUnit @adp -Filter * -Properties DistinguishedName |
        Sort-Object DistinguishedName |
        Select-Object -ExpandProperty DistinguishedName

        if ((-not [string]::IsNullOrWhiteSpace($usersContainerDn)) -and (-not ($ouDns -contains $usersContainerDn))) {
          $ouDns += $usersContainerDn
        }

        $ouDns = @(
          $ouDns |
          Where-Object { $_ -and ($_ -notmatch '^(?i)(OU|CN)=Domain Controllers,') } |
          Sort-Object -Unique
        )

        $cbOU.Items.Clear() | Out-Null
        foreach ($ouDn in @($ouDns)) {
          if (-not [string]::IsNullOrWhiteSpace($ouDn)) {
            $cbOU.Items.Add([PSCustomObject]@{
                Label             = (Get-ContainerLabelFromDN $ouDn)
                DistinguishedName = $ouDn
              }) | Out-Null
          }
        }

        $currentOuDn = Get-OUFromDN $user.DistinguishedName
        if ($cbOU.Items.Count -gt 0) {
          $prefIdx = -1
          for ($i = 0; $i -lt $cbOU.Items.Count; $i++) {
            if ([string]$cbOU.Items[$i].DistinguishedName -eq [string]$currentOuDn) { $prefIdx = $i; break }
          }
          if ($prefIdx -ge 0) { $cbOU.SelectedIndex = $prefIdx } else { $cbOU.SelectedIndex = 0 }
        }
      }
      catch {
        Show-Exception "Error cargando OUs para modificar usuario:" $_
      }

      $adpForDomain = Get-ADParams
      $domainRootForUpn = (Get-ADDomainCached -adParams $adpForDomain).DNSRoot
      $lblUpnSuffix.Text = "@$domainRootForUpn"

      $cbMustChange = New-Object System.Windows.Controls.CheckBox
      $cbMustChange.Content = "El usuario debe cambiar la contraseña en el próximo inicio"
      $mustChangeAtLogon = $false
      if ($null -ne $user.pwdLastSet) {
        try { $mustChangeAtLogon = ([Int64]$user.pwdLastSet -eq 0) } catch { $mustChangeAtLogon = $false }
      }
      $cbMustChange.IsChecked = $mustChangeAtLogon
      $cbMustChange.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
      $cbMustChange.Margin = '0,10,0,0'
      $cbMustChange.IsThreeState = $false
      $sp.Children.Add($cbMustChange)

      $cbCannotChange = New-Object System.Windows.Controls.CheckBox
      $cbCannotChange.Content = "El usuario no puede cambiar la contraseña"
      $cbCannotChange.IsChecked = [bool]$user.CannotChangePassword
      $cbCannotChange.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
      $cbCannotChange.Margin = '0,2,0,0'
      $cbCannotChange.IsThreeState = $false
      $sp.Children.Add($cbCannotChange)

      $cbPwdNeverExpires = New-Object System.Windows.Controls.CheckBox
      $cbPwdNeverExpires.Content = "La contraseña nunca expira"
      $cbPwdNeverExpires.IsChecked = [bool]$user.PasswordNeverExpires
      $cbPwdNeverExpires.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
      $cbPwdNeverExpires.Margin = '0,2,0,0'
      $cbPwdNeverExpires.IsThreeState = $false
      $sp.Children.Add($cbPwdNeverExpires)

      $cbEnabled = New-Object System.Windows.Controls.CheckBox
      $cbEnabled.Content = "Cuenta habilitada"
      $cbEnabled.IsChecked = [bool]$user.Enabled
      $cbEnabled.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
      $cbEnabled.Margin = '0,2,0,0'
      $cbEnabled.IsThreeState = $false
      $sp.Children.Add($cbEnabled)

      $modeUi = 'Modify'
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'FirstName' -LabelControl $lblFirstName -InputControl $tbFirstName
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Initials' -LabelControl $lblInitials -InputControl $tbInitials
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'LastName' -LabelControl $lblLastName -InputControl $tbLastName
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'FullName' -LabelControl $lblFullName -InputControl $tbFullName
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'DisplayName' -LabelControl $lblDisplayName -InputControl $tbDisplayName
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Description' -LabelControl $lblDescription -InputControl $tbDescription
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Office' -LabelControl $lblOffice -InputControl $tbOffice
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'OfficePhone' -LabelControl $lblOfficePhone -InputControl $tbOfficePhone
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Email' -LabelControl $lblEmail -InputControl $tbEmail
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'WebPage' -LabelControl $lblWebPage -InputControl $tbWebPage
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'StreetAddress' -LabelControl $lblStreetAddress -InputControl $tbStreetAddress
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'POBox' -LabelControl $lblPOBox -InputControl $tbPOBox
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'City' -LabelControl $lblCity -InputControl $tbCity
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'State' -LabelControl $lblState -InputControl $tbState
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'PostalCode' -LabelControl $lblPostalCode -InputControl $tbPostalCode
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Country' -LabelControl $lblCountry -InputControl $tbCountry
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'UpnUser' -LabelControl $lblUpn -InputControl $upnGrid
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'SamAccountName' -LabelControl $lblSam -InputControl $tbSam
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Password' -LabelControl $lblPwd -InputControl $tbPwd
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'Password' -LabelControl $lblPwd2 -InputControl $tbPwdConfirm
      Apply-UserFieldUiPolicy -Mode $modeUi -Field 'OU' -LabelControl $lblOU -InputControl $cbOU

      $cbMustChange.Visibility = if (Test-UserFieldVisible -Mode $modeUi -Field 'MustChangePassword') { 'Visible' } else { 'Collapsed' }
      $cbMustChange.IsEnabled = Test-UserFieldEditable -Mode $modeUi -Field 'MustChangePassword'
      $cbCannotChange.Visibility = if (Test-UserFieldVisible -Mode $modeUi -Field 'CannotChangePassword') { 'Visible' } else { 'Collapsed' }
      $cbCannotChange.IsEnabled = Test-UserFieldEditable -Mode $modeUi -Field 'CannotChangePassword'
      $cbPwdNeverExpires.Visibility = if (Test-UserFieldVisible -Mode $modeUi -Field 'PasswordNeverExpires') { 'Visible' } else { 'Collapsed' }
      $cbPwdNeverExpires.IsEnabled = Test-UserFieldEditable -Mode $modeUi -Field 'PasswordNeverExpires'
      $cbEnabled.Visibility = if (Test-UserFieldVisible -Mode $modeUi -Field 'Enabled') { 'Visible' } else { 'Collapsed' }
      $cbEnabled.IsEnabled = Test-UserFieldEditable -Mode $modeUi -Field 'Enabled'

      $btnSave = New-DlgButton "Guardar Cambios" 'Primary'
      $sp.Children.Add($btnSave)

      $scroll = New-Object System.Windows.Controls.ScrollViewer
      $scroll.Content = $sp
      $dlg.Content = $scroll

      $btnSave.Add_Click({
        try {
          $mode = 'Modify'
          $fn = Get-NormalizedText $tbFirstName.Text
          $ln = Get-NormalizedText $tbLastName.Text
          $cnName = Get-NormalizedText $tbFullName.Text
            $displayName = Get-NormalizedText $tbDisplayName.Text
            $upnUser = Get-NormalizedText $tbUpnUser.Text
            $sam = Get-NormalizedText $tbSam.Text
          $selectedOU = [string]$cbOU.SelectedValue
          $pwdPlain = $tbPwd.Password
          $pwdConfirm = $tbPwdConfirm.Password

          $effectiveSam = if (Test-UserFieldEditable -Mode $mode -Field 'SamAccountName') { $sam } else { [string]$user.SamAccountName }
          if ([string]::IsNullOrWhiteSpace($effectiveSam)) { $effectiveSam = [string]$user.SamAccountName }
          $effectiveCn = if (Test-UserFieldEditable -Mode $mode -Field 'FullName') { $cnName } else { [string]$user.Name }
          if ([string]::IsNullOrWhiteSpace($effectiveCn)) { $effectiveCn = [string]$user.Name }
          $effectiveUpnUser = if (Test-UserFieldEditable -Mode $mode -Field 'UpnUser') { $upnUser } else { '' }
          if ([string]::IsNullOrWhiteSpace($effectiveUpnUser)) {
            if ((-not [string]::IsNullOrWhiteSpace([string]$user.UserPrincipalName)) -and ([string]$user.UserPrincipalName -match '^([^@]+)@')) {
              $effectiveUpnUser = [string]$matches[1]
            }
            else {
              $effectiveUpnUser = $effectiveSam
            }
          }

          if (-not (Test-UserFieldValue -Mode $mode -Field 'FirstName' -Value $fn -ErrorMessage "El campo 'Nombre' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'LastName' -Value $ln -ErrorMessage "El campo 'Apellido' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'FullName' -Value $effectiveCn -ErrorMessage "El campo 'Nombre completo (CN)' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'UpnUser' -Value $effectiveUpnUser -ErrorMessage "El campo 'UPN' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'SamAccountName' -Value $effectiveSam -ErrorMessage "El campo 'SamAccountName' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'OU' -Value $selectedOU -ErrorMessage "Seleccioná una OU o contenedor destino (requerido por política).")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Initials' -Value (Get-NormalizedText $tbInitials.Text) -ErrorMessage "El campo 'Iniciales' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Description' -Value (Get-NormalizedText $tbDescription.Text) -ErrorMessage "El campo 'Descripción' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Office' -Value (Get-NormalizedText $tbOffice.Text) -ErrorMessage "El campo 'Oficina' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'OfficePhone' -Value (Get-NormalizedText $tbOfficePhone.Text) -ErrorMessage "El campo 'Teléfono' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Email' -Value (Get-NormalizedText $tbEmail.Text) -ErrorMessage "El campo 'Email' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'WebPage' -Value (Get-NormalizedText $tbWebPage.Text) -ErrorMessage "El campo 'Página Web' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'StreetAddress' -Value (Get-NormalizedText $tbStreetAddress.Text) -ErrorMessage "El campo 'Calle' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'POBox' -Value (Get-NormalizedText $tbPOBox.Text) -ErrorMessage "El campo 'Casilla Postal' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'City' -Value (Get-NormalizedText $tbCity.Text) -ErrorMessage "El campo 'Ciudad' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'State' -Value (Get-NormalizedText $tbState.Text) -ErrorMessage "El campo 'Provincia/Estado' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'PostalCode' -Value (Get-NormalizedText $tbPostalCode.Text) -ErrorMessage "El campo 'Código Postal' es requerido por política.")) { return }
          if (-not (Test-UserFieldValue -Mode $mode -Field 'Country' -Value (Get-NormalizedText $tbCountry.Text) -ErrorMessage "El campo 'País/Región' es requerido por política.")) { return }

          if ((Test-UserFieldEditable -Mode $mode -Field 'MustChangePassword') -and (Test-UserFieldEditable -Mode $mode -Field 'CannotChangePassword') -and (Get-Checked $cbMustChange) -and (Get-Checked $cbCannotChange)) {
            Show-Error "No podés marcar a la vez 'debe cambiar contraseña' y 'no puede cambiar contraseña'."
            return
          }
          $canSetPassword = Test-UserFieldEditable -Mode $mode -Field 'Password'
          if ($canSetPassword -and ((-not [string]::IsNullOrWhiteSpace($pwdPlain)) -or (-not [string]::IsNullOrWhiteSpace($pwdConfirm)))) {
            if ([string]::IsNullOrWhiteSpace($pwdPlain) -or [string]::IsNullOrWhiteSpace($pwdConfirm)) {
              Show-Error "Si querés cambiar contraseña, completá ambos campos."
              return
            }
            if ($pwdPlain -ne $pwdConfirm) {
                Show-Error "Las contraseñas no coinciden."
              return
            }
          }
          if ($selectedOU -and ($selectedOU -match '^(?i)(OU|CN)=Domain Controllers,')) { Show-Error "Por seguridad no se permite mover usuarios a 'Domain Controllers'."; return }

          $adp2 = Get-ADParams
          $domainRoot = (Get-ADDomainCached -adParams $adp2).DNSRoot
          $params = @{ Identity = [string]$user.SamAccountName }
          $hasUserChanges = $false
          $clearAttrs = @()
          if ((Test-UserFieldEditable -Mode $mode -Field 'FirstName') -and (-not [string]::IsNullOrWhiteSpace($fn))) { $params.GivenName = $fn; $hasUserChanges = $true }
          if (Test-UserFieldEditable -Mode $mode -Field 'Initials') {
            $initialsValue = Get-NormalizedText $tbInitials.Text
            if (-not [string]::IsNullOrWhiteSpace($initialsValue)) { $params.Initials = $initialsValue } else { $clearAttrs += 'initials' }
            $hasUserChanges = $true
          }
          if ((Test-UserFieldEditable -Mode $mode -Field 'LastName') -and (-not [string]::IsNullOrWhiteSpace($ln))) { $params.Surname = $ln; $hasUserChanges = $true }
          if (Test-UserFieldEditable -Mode $mode -Field 'SamAccountName') { $params.SamAccountName = $effectiveSam; $hasUserChanges = $true }
          if (Test-UserFieldEditable -Mode $mode -Field 'UpnUser') { $params.UserPrincipalName = "$effectiveUpnUser@$domainRoot"; $hasUserChanges = $true }
          if (Test-UserFieldEditable -Mode $mode -Field 'Enabled') { $params.Enabled = (Get-Checked $cbEnabled); $hasUserChanges = $true }
          if (Test-UserFieldEditable -Mode $mode -Field 'MustChangePassword') { $params.ChangePasswordAtLogon = (Get-Checked $cbMustChange); $hasUserChanges = $true }
          if (Test-UserFieldEditable -Mode $mode -Field 'CannotChangePassword') { $params.CannotChangePassword = (Get-Checked $cbCannotChange); $hasUserChanges = $true }
          if (Test-UserFieldEditable -Mode $mode -Field 'PasswordNeverExpires') { $params.PasswordNeverExpires = (Get-Checked $cbPwdNeverExpires); $hasUserChanges = $true }

          if (Test-UserFieldEditable -Mode $mode -Field 'DisplayName') {
            if (-not [string]::IsNullOrWhiteSpace($displayName)) { $params.DisplayName = $displayName } else { $clearAttrs += 'DisplayName' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'Description') {
            if (-not [string]::IsNullOrWhiteSpace($tbDescription.Text)) { $params.Description = $tbDescription.Text.Trim() } else { $clearAttrs += 'description' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'Email') {
            if (-not [string]::IsNullOrWhiteSpace($tbEmail.Text)) { $params.EmailAddress = $tbEmail.Text.Trim() } else { $clearAttrs += 'mail' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'Office') {
            if (-not [string]::IsNullOrWhiteSpace($tbOffice.Text)) { $params.Office = $tbOffice.Text.Trim() } else { $clearAttrs += 'physicalDeliveryOfficeName' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'OfficePhone') {
            if (-not [string]::IsNullOrWhiteSpace($tbOfficePhone.Text)) { $params.OfficePhone = $tbOfficePhone.Text.Trim() } else { $clearAttrs += 'telephoneNumber' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'WebPage') {
            if (-not [string]::IsNullOrWhiteSpace($tbWebPage.Text)) { $params.HomePage = $tbWebPage.Text.Trim() } else { $clearAttrs += 'wWWHomePage' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'StreetAddress') {
            if (-not [string]::IsNullOrWhiteSpace($tbStreetAddress.Text)) { $params.StreetAddress = $tbStreetAddress.Text.Trim() } else { $clearAttrs += 'streetAddress' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'POBox') {
            if (-not [string]::IsNullOrWhiteSpace($tbPOBox.Text)) { $params.POBox = $tbPOBox.Text.Trim() } else { $clearAttrs += 'postOfficeBox' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'City') {
            if (-not [string]::IsNullOrWhiteSpace($tbCity.Text)) { $params.City = $tbCity.Text.Trim() } else { $clearAttrs += 'l' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'State') {
            if (-not [string]::IsNullOrWhiteSpace($tbState.Text)) { $params.State = $tbState.Text.Trim() } else { $clearAttrs += 'st' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'PostalCode') {
            if (-not [string]::IsNullOrWhiteSpace($tbPostalCode.Text)) { $params.PostalCode = $tbPostalCode.Text.Trim() } else { $clearAttrs += 'postalCode' }
            $hasUserChanges = $true
          }
          if (Test-UserFieldEditable -Mode $mode -Field 'Country') {
            if (-not [string]::IsNullOrWhiteSpace($tbCountry.Text)) { $params.Country = $tbCountry.Text.Trim() } else { $clearAttrs += 'co' }
            $hasUserChanges = $true
          }
          if ($clearAttrs.Count -gt 0) { $params.Clear = @($clearAttrs | Sort-Object -Unique) }

          if ($hasUserChanges -or ($clearAttrs.Count -gt 0)) {
            $setParams = Merge-ADParams -params $params -adParams $adp2
            Set-ADUser @setParams
          }

          $lookupIdentity = if (Test-UserFieldEditable -Mode $mode -Field 'SamAccountName') { $effectiveSam } else { [string]$user.SamAccountName }
          $refreshed = Get-ADUser @adp2 -Identity $lookupIdentity -Properties DistinguishedName, Name

          if ($canSetPassword -and (-not [string]::IsNullOrWhiteSpace($pwdPlain))) {
            $pwd = ConvertTo-SecureString $pwdPlain -AsPlainText -Force
            Set-ADAccountPassword @adp2 -Identity $refreshed.DistinguishedName -NewPassword $pwd -Reset
          }
          if ((Test-UserFieldEditable -Mode $mode -Field 'FullName') -and (-not [string]::IsNullOrWhiteSpace($effectiveCn)) -and ($effectiveCn -ne $refreshed.Name)) {
            $renameParams = Merge-ADParams -params @{ Identity = $refreshed.DistinguishedName; NewName = $effectiveCn } -adParams $adp2
            Rename-ADObject @renameParams
            $refreshed = Get-ADUser @adp2 -Identity $lookupIdentity -Properties DistinguishedName, Name
          }

          if ((Test-UserFieldEditable -Mode $mode -Field 'OU') -and (-not [string]::IsNullOrWhiteSpace($selectedOU))) {
            $currentOuAfterUpdate = Get-OUFromDN $refreshed.DistinguishedName
            if ($selectedOU -ne $currentOuAfterUpdate) {
              Move-ADObject @adp2 -Identity $refreshed.DistinguishedName -TargetPath $selectedOU -ErrorAction Stop
            }
          }

          [System.Windows.MessageBox]::Show("Usuario modificado correctamente.", "Éxito", "OK", "Information") | Out-Null
          Set-Status "✅ Usuario '$effectiveSam' modificado." '#a6e3a1'
          Write-AuditLog -Action 'Modificar Usuario' -Result 'OK' -Object $effectiveSam
            Raise-ButtonClick $btnUserSearch
            $dlg.Close()
          }
          catch { Show-Exception "Error al modificar usuario:" $_ }
        }.GetNewClosure())

      $dlg.ShowDialog() | Out-Null
    }
    catch { Show-Exception "Error preparando edición de usuario:" $_ }
  })

# Resetear contraseña
$btnUserResetPwd.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $sel = Get-SelectedUser; if (-not $sel) { return }

    $dlg = New-DialogWindow "Resetear Contraseña" 400 250
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'
    $sp.Children.Add((New-DlgLabel "Usuario: $($sel.SamAccountName)"))
    $sp.Children.Add((New-DlgLabel "Nueva contraseña:"))
    $tbPwd = New-DlgPasswordBox; $sp.Children.Add($tbPwd)
    $sp.Children.Add((New-DlgLabel "Repetir nueva contraseña:"))
    $tbPwdConfirm = New-DlgPasswordBox; $sp.Children.Add($tbPwdConfirm)

    $cbChange = New-Object System.Windows.Controls.CheckBox
    $cbChange.Content = "Cambiar contraseña en próximo inicio"
    $cbChange.IsChecked = $true
    $cbChange.Foreground = ConvertTo-Brush (Get-ActiveTheme).TextPrimary
    $cbChange.Margin = '0,8,0,0'
    $cbChange.IsThreeState = $false
    $sp.Children.Add($cbChange)

    $btnReset = New-DlgButton "Resetear" 'Warn'; $sp.Children.Add($btnReset)
    $dlg.Content = $sp

    $btnReset.Add_Click({
        try {
          if ([string]::IsNullOrWhiteSpace($tbPwd.Password) -or [string]::IsNullOrWhiteSpace($tbPwdConfirm.Password)) {
            Show-Error "Ingresá y repetí la nueva contraseña."
            return
          }
          if ($tbPwd.Password -ne $tbPwdConfirm.Password) {
            Show-Error "Las contraseñas no coinciden."
            return
          }
          $adp = Get-ADParams
          $pwd = ConvertTo-SecureString $tbPwd.Password -AsPlainText -Force
          Set-ADAccountPassword @adp -Identity $sel.SamAccountName -NewPassword $pwd -Reset
          if (Get-Checked $cbChange) { Set-ADUser @adp -Identity $sel.SamAccountName -ChangePasswordAtLogon $true }
          [System.Windows.MessageBox]::Show("Contraseña reseteada.", "Éxito", "OK", "Information") | Out-Null
          Set-Status "✅ Contraseña de '$($sel.SamAccountName)' reseteada." '#a6e3a1'
          Write-AuditLog -Action 'Resetear Contraseña Usuario' -Result 'OK' -Object ([string]$sel.SamAccountName)
          $dlg.Close()
        }
        catch { Show-Exception "Error al resetear contraseña:" $_ }
      }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
  })

# Deshabilitar / Habilitar / Desbloquear / Eliminar
$btnUserDisable.Add_Click({
    Invoke-SelectedUserAction `
      -Action { param($sel) $adp = Get-ADParams; Disable-ADAccount @adp -Identity $sel.SamAccountName } `
      -SuccessMessage "✅ Usuario '{0}' deshabilitado." `
      -ErrorContext "Error deshabilitando usuario:" `
      -ConfirmMessage "¿Deshabilitar al usuario '{0}'?" `
      -AuditAction "Deshabilitar Usuario"
  })
$btnUserEnable.Add_Click({
    Invoke-SelectedUserAction `
      -Action { param($sel) $adp = Get-ADParams; Enable-ADAccount @adp -Identity $sel.SamAccountName } `
      -SuccessMessage "✅ Usuario '{0}' habilitado." `
      -ErrorContext "Error habilitando usuario:" `
      -AuditAction "Habilitar Usuario"
  })
$btnUserUnlock.Add_Click({
    Invoke-SelectedUserAction `
      -Action { param($sel) $adp = Get-ADParams; Unlock-ADAccount @adp -Identity $sel.SamAccountName } `
      -SuccessMessage "✅ Usuario '{0}' desbloqueado." `
      -ErrorContext "Error desbloqueando usuario:" `
      -AuditAction "Desbloquear Usuario"
  })
$btnUserDelete.Add_Click({
    Invoke-SelectedUserAction `
      -Action { param($sel) $adp = Get-ADParams; Remove-ADUser @adp -Identity $sel.SamAccountName -Confirm:$false } `
      -SuccessMessage "✅ Usuario '{0}' eliminado." `
      -ErrorContext "Error eliminando usuario:" `
      -ConfirmMessage "⚠️ ¿Estás seguro de ELIMINAR al usuario '{0}'?`nEsta acción no se puede deshacer." `
      -AuditAction "Eliminar Usuario"
  })

# ── GRUPOS ───────────────────────────────────────────────────────────────────
$btnGroupSearch.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $filter = Get-NormalizedText $txtGroupSearch.Text
    $maxResults = 500
    try {
      [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
      Set-Status "Buscando grupos..." '#89b4fa'
      $adp = Get-ADParams

      if ([string]::IsNullOrWhiteSpace($filter)) {
        $groups = Get-ADGroup @adp `
          -Filter * `
          -ResultSetSize ($maxResults + 1) `
          -Properties Name, SamAccountName, GroupCategory, GroupScope, Description |
        Select-Object Name, SamAccountName, GroupCategory, GroupScope, Description, DistinguishedName
      }
      else {
        $safeFilter = Escape-ADFilterLiteral $filter
        $groups = Get-ADGroup @adp `
          -Filter "Name -like '*$safeFilter*' -or SamAccountName -like '*$safeFilter*'" `
          -ResultSetSize ($maxResults + 1) `
          -Properties Name, SamAccountName, GroupCategory, GroupScope, Description |
        Select-Object Name, SamAccountName, GroupCategory, GroupScope, Description, DistinguishedName
      }

      $totalFound = @($groups).Count
      if ($totalFound -gt $maxResults) {
        $groups = @($groups | Select-Object -First $maxResults)
      }

      $dgGroups.ItemsSource = @($groups)
      if ($totalFound -gt $maxResults) {
        Set-Status "⚠️ Se muestran los primeros $maxResults grupos. Refiná la búsqueda para acotar resultados." '#fab387'
      }
      else {
        Set-Status "✅ Se encontraron $($groups.Count) grupo(s)." '#a6e3a1'
      }
    }
    catch { Show-Exception "Error buscando grupos:" $_ }
    finally { [System.Windows.Input.Mouse]::OverrideCursor = $null }
  })
Bind-EnterKeyToButton -textBox $txtGroupSearch -button $btnGroupSearch

$dgGroups.Add_MouseDoubleClick({
    if ($dgGroups.SelectedItem) {
      $btnGroupMembers.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
  })

$btnGroupNew.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $dlg = New-DialogWindow "Crear Nuevo Grupo" 440 400
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'
    $sp.Children.Add((New-DlgLabel "Nombre del grupo:")); $tbName = New-DlgTextBox; $sp.Children.Add($tbName)
    $sp.Children.Add((New-DlgLabel "Descripción:")); $tbDesc = New-DlgTextBox; $sp.Children.Add($tbDesc)
    $sp.Children.Add((New-DlgLabel "Categoría:")); $cbCat = New-DlgComboBox @('Security', 'Distribution'); $sp.Children.Add($cbCat)
    $sp.Children.Add((New-DlgLabel "Ámbito:")); $cbScope = New-DlgComboBox @('Global', 'Universal', 'DomainLocal'); $sp.Children.Add($cbScope)
    $sp.Children.Add((New-DlgLabel "OU Destino (opcional):")); $tbOU = New-DlgTextBox; $sp.Children.Add($tbOU)
    $btnCreate = New-DlgButton "Crear Grupo" 'Success'; $sp.Children.Add($btnCreate)
    $dlg.Content = $sp

    $btnCreate.Add_Click({
        try {
          $name = Get-RequiredTrimmedText $tbName.Text "Ingresá el nombre del grupo."
          if ($null -eq $name) { return }
          $params = @{
            Name           = $name
            SamAccountName = $name
            GroupCategory  = $cbCat.SelectedItem
            GroupScope     = $cbScope.SelectedItem
          }
          if ($tbDesc.Text.Trim()) { $params.Description = $tbDesc.Text.Trim() }
          if ($tbOU.Text.Trim()) { $params.Path = $tbOU.Text.Trim() }
          $newGroupParams = Merge-ADParams -params $params
          New-ADGroup @newGroupParams
          [System.Windows.MessageBox]::Show("Grupo '$name' creado.", "Éxito", "OK", "Information") | Out-Null
          Set-Status "✅ Grupo creado." '#a6e3a1'
          Write-AuditLog -Action 'Crear Grupo' -Result 'OK' -Object $name
          $dlg.Close()
        }
        catch { Show-Exception "Error al crear grupo:" $_ }
      }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
  })

$btnGroupAddMember.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $sel = Get-SelectedGroup; if (-not $sel) { return }
    $dlg = New-DialogWindow "Agregar Miembro a: $($sel.Name)" 440 200
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'
    $sp.Children.Add((New-DlgLabel "SamAccountName del usuario/grupo a agregar:"))
    $tbMember = New-DlgTextBox; $sp.Children.Add($tbMember)
    $btnAdd = New-DlgButton "Agregar" 'Success'; $sp.Children.Add($btnAdd)
    $dlg.Content = $sp

    $btnAdd.Add_Click({
        try {
          $m = Get-RequiredTrimmedText $tbMember.Text "Ingresá el SamAccountName a agregar."
          if ($null -eq $m) { return }
          $adp = Get-ADParams
          Add-ADGroupMember @adp -Identity $sel.SamAccountName -Members $m
          [System.Windows.MessageBox]::Show("Miembro agregado al grupo '$($sel.Name)'.", "Éxito", "OK", "Information") | Out-Null
          Set-Status "✅ Miembro agregado." '#a6e3a1'
          Write-AuditLog -Action 'Agregar Miembro a Grupo' -Result 'OK' -Object ("Grupo: $($sel.Name) | Miembro: $m")
          $dlg.Close()
        }
        catch { Show-Exception "Error agregando miembro:" $_ }
      }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
  })

$btnGroupRemoveMember.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $sel = Get-SelectedGroup; if (-not $sel) { return }
    $dlg = New-DialogWindow "Quitar Miembro de: $($sel.Name)" 440 200
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'
    $sp.Children.Add((New-DlgLabel "SamAccountName del usuario/grupo a quitar:"))
    $tbMember = New-DlgTextBox; $sp.Children.Add($tbMember)
    $btnRemove = New-DlgButton "Quitar" 'Danger'; $sp.Children.Add($btnRemove)
    $dlg.Content = $sp

    $btnRemove.Add_Click({
        try {
          $m = Get-RequiredTrimmedText $tbMember.Text "Ingresá el SamAccountName a quitar."
          if ($null -eq $m) { return }
          $adp = Get-ADParams
          Remove-ADGroupMember @adp -Identity $sel.SamAccountName -Members $m -Confirm:$false
          [System.Windows.MessageBox]::Show("Miembro quitado del grupo '$($sel.Name)'.", "Éxito", "OK", "Information") | Out-Null
          Set-Status "✅ Miembro quitado." '#a6e3a1'
          Write-AuditLog -Action 'Quitar Miembro de Grupo' -Result 'OK' -Object ("Grupo: $($sel.Name) | Miembro: $m")
          $dlg.Close()
        }
        catch { Show-Exception "Error quitando miembro:" $_ }
      }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
  })

$btnGroupMembers.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $sel = Get-SelectedGroup; if (-not $sel) { return }
    try {
      $theme = Get-ActiveTheme
      $adp = Get-ADParams
      $members = Get-ADGroupMember @adp -Identity $sel.SamAccountName | Select-Object Name, SamAccountName, objectClass
      $dlg = New-DialogWindow "Miembros de: $($sel.Name)" 500 400
      $dg = New-Object System.Windows.Controls.DataGrid
      $dg.AutoGenerateColumns = $true; $dg.IsReadOnly = $true; $dg.CanUserAddRows = $false
      $dg.Background = ConvertTo-Brush $theme.SurfaceBg
      $dg.Foreground = ConvertTo-Brush $theme.TextPrimary
      $dg.BorderBrush = ConvertTo-Brush $theme.Border
      $dg.RowBackground = ConvertTo-Brush $theme.RowBg
      $dg.AlternatingRowBackground = ConvertTo-Brush $theme.RowAltBg
      $dg.HorizontalGridLinesBrush = ConvertTo-Brush $theme.GridLine
      $dg.ItemsSource = @($members)
      $dlg.Content = $dg
      $dlg.ShowDialog() | Out-Null
    }
    catch { Show-Exception "Error obteniendo miembros:" $_ }
  })

$btnGroupDelete.Add_Click({
    Invoke-SelectedGroupAction `
      -Action { param($sel) $adp = Get-ADParams; Remove-ADGroup @adp -Identity $sel.SamAccountName -Confirm:$false } `
      -SuccessMessage "✅ Grupo '{0}' eliminado." `
      -ErrorContext "Error eliminando grupo:" `
      -ConfirmMessage "⚠️ ¿Estás seguro de ELIMINAR el grupo '{0}'?`nEsta acción no se puede deshacer." `
      -AuditAction "Eliminar Grupo"
  })


# ── MOVER ────────────────────────────────────────────────────────────────────
function Load-OUTreeForControl(
  [System.Windows.Controls.TreeView]$treeView,
  [switch]$HideProtectedContainers
) {
  if (-not (Assert-ADModule)) { return }
  try {
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    $treeView.Items.Clear()
    $adp = Get-ADParams
    $domain = Get-ADDomainCached -adParams $adp

    $rootItem = New-Object System.Windows.Controls.TreeViewItem
    $rootItem.Header = $domain.DNSRoot
    $rootItem.Tag = $domain.DistinguishedName
    $rootItem.IsExpanded = $true
    $theme = Get-ActiveTheme
    $rootItem.Foreground = ConvertTo-Brush $theme.TreeItemFg

    # Traer OUs y contenedores para verse como ADUC (Users, Computers, Builtin, etc.).
    $nodes = Get-ADObject @adp `
      -SearchBase $domain.DistinguishedName `
      -SearchScope Subtree `
      -LDAPFilter "(|(objectClass=organizationalUnit)(objectClass=container))" `
      -Properties Name, DistinguishedName, ObjectClass |
    Where-Object {
      ($_.DistinguishedName -ne $domain.DistinguishedName) -and
      ($_.Name -ne 'DomainUpdates')
    } |
    Sort-Object DistinguishedName

    if ($HideProtectedContainers) {
      $nodes = @($nodes | Where-Object { -not (Is-ProtectedContainerDN $_.DistinguishedName) })
    }

    $nodeHash = @{}
    foreach ($ou in $nodes) {
      $item = New-Object System.Windows.Controls.TreeViewItem
      $item.Header = $ou.Name
      $item.Tag = $ou.DistinguishedName
      $item.Foreground = ConvertTo-Brush $theme.TreeItemFg
      $nodeHash[$ou.DistinguishedName] = $item
    }

    foreach ($ou in $nodes) {
      $parentDN = Get-OUFromDN $ou.DistinguishedName
      if ($nodeHash.ContainsKey($parentDN)) {
        $nodeHash[$parentDN].Items.Add($nodeHash[$ou.DistinguishedName]) | Out-Null
      }
      else {
        $rootItem.Items.Add($nodeHash[$ou.DistinguishedName]) | Out-Null
      }
    }

    $treeView.Items.Add($rootItem) | Out-Null
  }
  catch { Show-Exception "Error cargando OUs:" $_ }
  finally { [System.Windows.Input.Mouse]::OverrideCursor = $null }
}
$btnMoveRefreshSourceOU.Add_Click({
    Load-OUTreeForControl $tvMoveSourceOUs
    Set-Status "✅ OUs de origen recargadas." '#a6e3a1'
  })
$btnMoveRefreshTargetOU.Add_Click({
    Load-OUTreeForControl -treeView $tvMoveTargetOUs -HideProtectedContainers
    Set-Status "✅ OUs de destino recargadas." '#a6e3a1'
  })

$btnMoveNewOU.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $parentOU = $tvMoveTargetOUs.SelectedItem
    $parentDN = Get-TreeSelectionDN -treeView $tvMoveTargetOUs `
      -selectionMessage "Seleccioná una OU/contendor padre en destino." `
      -invalidDnMessage "La OU padre no es válida." `
      -DenyProtected `
      -protectedMessage "Por seguridad no se permite crear OUs dentro de 'Domain Controllers'."
    if ($null -eq $parentDN) { return }

    $dlg = New-DialogWindow "Crear OU" 420 210
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = '16'
    $sp.Children.Add((New-DlgLabel "Padre: $($parentOU.Header)"))
    $sp.Children.Add((New-DlgLabel "Nombre de la nueva OU:"))
    $tbName = New-DlgTextBox; $sp.Children.Add($tbName)
    $btnCreateOU = New-DlgButton "Crear OU" 'Success'; $sp.Children.Add($btnCreateOU)
    $dlg.Content = $sp

    $btnCreateOU.Add_Click({
        try {
          $name = Get-RequiredTrimmedText $tbName.Text "Ingresá el nombre de la OU."
          if ($null -eq $name) { return }
          $adp = Get-ADParams
          New-ADOrganizationalUnit @adp -Name $name -Path $parentDN -ProtectedFromAccidentalDeletion $true
          Set-Status "✅ OU '$name' creada." '#a6e3a1'
          Write-AuditLog -Action 'Crear OU' -Result 'OK' -Object $name
          Load-OUTreeForControl -treeView $tvMoveTargetOUs -HideProtectedContainers
          Load-OUTreeForControl $tvMoveSourceOUs
          $dlg.Close()
        }
        catch { Show-Exception "Error creando OU:" $_ }
      }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
  })

$btnMoveDeleteOU.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $selOU = $tvMoveTargetOUs.SelectedItem
    $selDN = Get-TreeSelectionDN -treeView $tvMoveTargetOUs `
      -selectionMessage "Seleccioná la OU a eliminar en destino." `
      -invalidDnMessage "La OU seleccionada no es válida."
    if ($null -eq $selDN) { return }
    if (-not (Test-Guard ($selDN -match '^OU=') "Solo se pueden eliminar objetos tipo OU.")) { return }

    if (Show-Confirm "⚠️ ¿Eliminar OU '$($selOU.Header)'?`nIntentaré quitar la protección por eliminación accidental.") {
      try {
        $adp = Get-ADParams
        Set-ADOrganizationalUnit @adp -Identity $selDN -ProtectedFromAccidentalDeletion $false -ErrorAction SilentlyContinue
        Remove-ADOrganizationalUnit @adp -Identity $selDN -Recursive -Confirm:$false
        Set-Status "✅ OU '$($selOU.Header)' eliminada." '#a6e3a1'
        Write-AuditLog -Action 'Eliminar OU' -Result 'OK' -Object ([string]$selOU.Header)
        Load-OUTreeForControl -treeView $tvMoveTargetOUs -HideProtectedContainers
        Load-OUTreeForControl $tvMoveSourceOUs
      }
      catch { Show-Exception "Error eliminando OU:" $_ }
    }
  })

function Get-ObjectsForOUDN([string]$dn) {
  $adp = Get-ADParams
  $objs = @()
  $objs += Get-ADUser @adp -SearchBase $dn -SearchScope Subtree -Filter * -Properties Name, DistinguishedName |
  Select-Object Name, @{N = 'ObjectClass'; E = { 'user' } }, @{N = 'OUPath'; E = { Get-OUFromDN $_.DistinguishedName } }, DistinguishedName
  $objs += Get-ADGroup @adp -SearchBase $dn -SearchScope Subtree -Filter * -Properties Name, DistinguishedName |
  Select-Object Name, @{N = 'ObjectClass'; E = { 'group' } }, @{N = 'OUPath'; E = { Get-OUFromDN $_.DistinguishedName } }, DistinguishedName
  $objs += Get-ADComputer @adp -SearchBase $dn -SearchScope Subtree -Filter * -Properties Name, DistinguishedName |
  Select-Object Name, @{N = 'ObjectClass'; E = { 'computer' } }, @{N = 'OUPath'; E = { Get-OUFromDN $_.DistinguishedName } }, DistinguishedName
  $objs += Get-ADOrganizationalUnit @adp -SearchBase $dn -SearchScope Subtree -Filter * -Properties Name, DistinguishedName |
  Where-Object { $_.DistinguishedName -ne $dn } |
  Select-Object Name, @{N = 'ObjectClass'; E = { 'organizationalUnit' } }, @{N = 'OUPath'; E = { Get-OUFromDN $_.DistinguishedName } }, DistinguishedName
  return @($objs)
}

$btnMoveLoadFromSource.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $sourceDN = Get-TreeSelectionDN -treeView $tvMoveSourceOUs `
      -selectionMessage "Seleccioná una OU de origen." `
      -invalidDnMessage "La OU de origen no es válida."
    if ($null -eq $sourceDN) { return }

    try {
      [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
      Set-Status "Cargando objetos desde OU origen..." '#89b4fa'
      $objs = Get-ObjectsForOUDN -dn $sourceDN
      $dgMoveObjects.ItemsSource = @($objs)
      Set-Status "✅ Se encontraron $($objs.Count) objeto(s) en origen." '#a6e3a1'
    }
    catch { Show-Exception "Error cargando objetos de OU origen:" $_ }
    finally { [System.Windows.Input.Mouse]::OverrideCursor = $null }
  })

$btnMoveLoadFromTarget.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $targetDN = Get-TreeSelectionDN -treeView $tvMoveTargetOUs `
      -selectionMessage "Seleccioná una OU de destino." `
      -invalidDnMessage "La OU de destino no es válida." `
      -DenyProtected `
      -protectedMessage "Por seguridad no se permite operar sobre 'Domain Controllers' como destino."
    if ($null -eq $targetDN) { return }

    try {
      [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
      Set-Status "Cargando objetos desde OU destino..." '#89b4fa'
      $objs = Get-ObjectsForOUDN -dn $targetDN
      $dgMoveTargetObjects.ItemsSource = @($objs)
      Set-Status "✅ Se encontraron $($objs.Count) objeto(s) en destino." '#a6e3a1'
    }
    catch { Show-Exception "Error cargando objetos de OU destino:" $_ }
    finally { [System.Windows.Input.Mouse]::OverrideCursor = $null }
  })

$btnMoveExecute.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $sourceDN = Get-TreeSelectionDN -treeView $tvMoveSourceOUs `
      -selectionMessage "Seleccioná una OU de origen." `
      -invalidDnMessage "La OU de origen no es válida."
    if ($null -eq $sourceDN) { return }
    $selectedItems = @($dgMoveObjects.SelectedItems)
    if (-not (Test-Guard ($selectedItems.Count -gt 0) "Seleccioná uno o más objetos de la lista.")) { return }
    $targetOU = $tvMoveTargetOUs.SelectedItem
    $targetDN = Get-TreeSelectionDN -treeView $tvMoveTargetOUs `
      -selectionMessage "Seleccioná una OU de destino." `
      -invalidDnMessage "La OU de destino no es válida." `
      -DenyProtected `
      -protectedMessage "Por seguridad no se permite mover objetos a 'Domain Controllers'."
    if ($null -eq $targetDN) { return }

    $confirmMsg = if ($selectedItems.Count -eq 1) {
      "¿Mover '$($selectedItems[0].Name)' a '$($targetOU.Header)'?"
    }
    else {
      "¿Mover $($selectedItems.Count) objetos a '$($targetOU.Header)'?"
    }

    if (Show-Confirm $confirmMsg) {
      try {
        $adp = Get-ADParams
        $moved = @()
        $failed = @()

        foreach ($sel in $selectedItems) {
          try {
            Move-ADObject @adp -Identity $sel.DistinguishedName -TargetPath $targetDN -ErrorAction Stop
            $moved += $sel.Name
          }
          catch {
            $failed += [PSCustomObject]@{
              Name    = $sel.Name
              Message = $_.Exception.Message
            }
          }
        }

        if ($failed.Count -eq 0) {
          if ($moved.Count -eq 1) {
            Set-Status "✅ '$($moved[0])' movido a '$($targetOU.Header)'." '#a6e3a1'
          }
          else {
            Set-Status "✅ $($moved.Count) objeto(s) movidos a '$($targetOU.Header)'." '#a6e3a1'
          }
          Write-AuditLog -Action 'Mover Objetos OU' -Result 'OK' -Object ("Destino: $($targetOU.Header) | Cantidad: $($moved.Count)")
        }
        else {
          Set-Status "⚠️ Movidos: $($moved.Count). Fallidos: $($failed.Count)." '#fab387'
          Write-AuditLog -Action 'Mover Objetos OU' -Result 'WARN' -Object ("Destino: $($targetOU.Header) | Movidos: $($moved.Count) | Fallidos: $($failed.Count)")
          $preview = ($failed | Select-Object -First 8 | ForEach-Object { " - $($_.Name): $($_.Message)" }) -join "`n"
          [System.Windows.MessageBox]::Show(
            "Finalizó con errores.`n`nMovidos: $($moved.Count)`nFallidos: $($failed.Count)`n`n$preview",
            "Mover múltiples objetos",
            "OK",
            "Warning"
          ) | Out-Null
        }

        Raise-ButtonClick $btnMoveLoadFromSource
        Raise-ButtonClick $btnMoveLoadFromTarget
      }
      catch { Show-Exception "Error al mover:" $_ }
    }
  })

# ── LOTE CSV ─────────────────────────────────────────────────────────────────
$script:batchLastResults = @()
$script:batchLastReportPath = $null

function Get-CsvValue($row, [string]$name) {
  if ($null -eq $row) { return '' }
  $prop = $row.PSObject.Properties[$name]
  if ($null -eq $prop -or $null -eq $prop.Value) { return '' }
  return ([string]$prop.Value).Trim()
}

function ConvertTo-BoolOrDefault([string]$value, [bool]$defaultValue = $false) {
  if ([string]::IsNullOrWhiteSpace($value)) { return $defaultValue }
  switch -Regex ($value.Trim().ToLowerInvariant()) {
    '^(1|true|t|si|sí|yes|y|on)$' { return $true }
    '^(0|false|f|no|n|off)$' { return $false }
    default { return $defaultValue }
  }
}

function Get-BatchIdentity($row) {
  $id = Get-CsvValue $row 'Identity'
  if ([string]::IsNullOrWhiteSpace($id)) { $id = Get-CsvValue $row 'SamAccountName' }
  if ([string]::IsNullOrWhiteSpace($id)) { $id = Get-CsvValue $row 'DistinguishedName' }
  return $id
}

function New-BatchResult([int]$rowNumber, [string]$action, [string]$identity, [string]$status, [string]$message) {
  return [PSCustomObject]@{
    Row      = $rowNumber
    Action   = $action
    Identity = $identity
    Status   = $status
    Message  = $message
  }
}

function Invoke-BatchCsvProcess([string]$csvPath, [switch]$Preview) {
  if (-not (Assert-ADModule)) { return @() }

  $path = Get-RequiredTrimmedText $csvPath "Seleccioná un archivo CSV."
  if ($null -eq $path) { return @() }
  if (-not (Test-Path -Path $path -PathType Leaf)) {
    Show-Error "No existe el archivo: $path"
    return @()
  }

  try {
    $rows = Import-Csv -Path $path -ErrorAction Stop
  }
  catch {
    Show-Exception "Error leyendo CSV:" $_
    return @()
  }

  if ($null -eq $rows -or $rows.Count -eq 0) {
    Show-Error "El CSV está vacío."
    return @()
  }

  $simulate = $Preview.IsPresent -or $script:SimulationMode
  $modeLabel = if ($simulate) { "SIMULACIÓN" } else { "EJECUCIÓN" }
  Set-Status "Procesando lote CSV ($modeLabel)..." '#89b4fa'

  $results = @()
  $adp = Get-ADParams
  $domainRoot = ''
  try { $domainRoot = (Get-ADDomainCached -adParams $adp).DNSRoot } catch {}

  for ($idx = 0; $idx -lt $rows.Count; $idx++) {
    $rowNumber = $idx + 1
    $row = $rows[$idx]
    $actionRaw = Get-CsvValue $row 'Action'
    $action = $actionRaw.ToLowerInvariant()
    $identity = Get-BatchIdentity $row

    try {
      if ([string]::IsNullOrWhiteSpace($action)) {
        throw "Falta columna Action en fila $rowNumber (valores: Alta, Modificar, Baja, Mover)."
      }

      switch ($action) {
        'alta' {
          $fn = Get-CsvValue $row 'FirstName'
          $ln = Get-CsvValue $row 'LastName'
          $sam = Get-CsvValue $row 'SamAccountName'
          $upnUser = Get-CsvValue $row 'UpnUser'
          $cnName = Get-CsvValue $row 'FullName'
          $displayName = Get-CsvValue $row 'DisplayName'
          $ou = Get-CsvValue $row 'OU'
          $pwdPlain = Get-CsvValue $row 'Password'

          if ([string]::IsNullOrWhiteSpace($fn) -or [string]::IsNullOrWhiteSpace($ln) -or [string]::IsNullOrWhiteSpace($sam) -or [string]::IsNullOrWhiteSpace($ou)) {
            throw "Alta requiere FirstName, LastName, SamAccountName y OU."
          }
          if ($ou -match '^(?i)(OU|CN)=Domain Controllers,') {
            throw "No se permite crear usuarios en 'Domain Controllers'."
          }
          if ([string]::IsNullOrWhiteSpace($cnName)) { $cnName = "$fn $ln" }
          if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $cnName }
          if ([string]::IsNullOrWhiteSpace($upnUser)) { $upnUser = $sam }
          if ([string]::IsNullOrWhiteSpace($pwdPlain) -and (-not $simulate)) {
            throw "Alta requiere Password en ejecución real."
          }

          $params = @{
            Name              = $cnName
            GivenName         = $fn
            Surname           = $ln
            SamAccountName    = $sam
            UserPrincipalName = "$upnUser@$domainRoot"
            DisplayName       = $displayName
            Enabled           = (ConvertTo-BoolOrDefault (Get-CsvValue $row 'Enabled') $true)
            Path              = $ou
            WhatIf            = $simulate
          }
          $initials = Get-CsvValue $row 'Initials'
          if (-not [string]::IsNullOrWhiteSpace($initials)) { $params.Initials = $initials }
          if (-not [string]::IsNullOrWhiteSpace($pwdPlain)) { $params.AccountPassword = (ConvertTo-SecureString $pwdPlain -AsPlainText -Force) }

          $email = Get-CsvValue $row 'Email'
          if (-not [string]::IsNullOrWhiteSpace($email)) { $params.EmailAddress = $email }
          $office = Get-CsvValue $row 'Office'
          if (-not [string]::IsNullOrWhiteSpace($office)) { $params.Office = $office }

          $createParams = Merge-ADParams -params $params -adParams $adp
          New-ADUser @createParams

          $mustChange = ConvertTo-BoolOrDefault (Get-CsvValue $row 'MustChangePassword') $true
          $cannotChange = ConvertTo-BoolOrDefault (Get-CsvValue $row 'CannotChangePassword') $false
          $neverExpires = ConvertTo-BoolOrDefault (Get-CsvValue $row 'PasswordNeverExpires') $false

          $postParams = Merge-ADParams -params @{
            Identity              = $sam
            ChangePasswordAtLogon = $mustChange
            CannotChangePassword  = $cannotChange
            PasswordNeverExpires  = $neverExpires
            WhatIf                = $simulate
          } -adParams $adp

          Set-ADUser @postParams
          $identity = $sam
        }
        'modificar' {
          if ([string]::IsNullOrWhiteSpace($identity)) { throw "Modificar requiere Identity o SamAccountName." }

          $params = Merge-ADParams -params @{
            Identity = $identity
            WhatIf   = $simulate
          } -adParams $adp

          $fn = Get-CsvValue $row 'FirstName'
          $ini = Get-CsvValue $row 'Initials'
          $ln = Get-CsvValue $row 'LastName'
          $displayName = Get-CsvValue $row 'DisplayName'
          $upnUser = Get-CsvValue $row 'UpnUser'
          $sam = Get-CsvValue $row 'SamAccountName'

          if (-not [string]::IsNullOrWhiteSpace($fn)) { $params.GivenName = $fn }
          if (-not [string]::IsNullOrWhiteSpace($ini)) { $params.Initials = $ini }
          if (-not [string]::IsNullOrWhiteSpace($ln)) { $params.Surname = $ln }
          if (-not [string]::IsNullOrWhiteSpace($displayName)) { $params.DisplayName = $displayName }
          if (-not [string]::IsNullOrWhiteSpace($sam)) { $params.SamAccountName = $sam; $identity = $sam }
          if (-not [string]::IsNullOrWhiteSpace($upnUser)) { $params.UserPrincipalName = "$upnUser@$domainRoot" }

          $email = Get-CsvValue $row 'Email'
          $title = Get-CsvValue $row 'Title'
          $dept = Get-CsvValue $row 'Department'
          $company = Get-CsvValue $row 'Company'
          $office = Get-CsvValue $row 'Office'
          $phone = Get-CsvValue $row 'OfficePhone'
          $desc = Get-CsvValue $row 'Description'
          if (-not [string]::IsNullOrWhiteSpace($email)) { $params.EmailAddress = $email }
          if (-not [string]::IsNullOrWhiteSpace($title)) { $params.Title = $title }
          if (-not [string]::IsNullOrWhiteSpace($dept)) { $params.Department = $dept }
          if (-not [string]::IsNullOrWhiteSpace($company)) { $params.Company = $company }
          if (-not [string]::IsNullOrWhiteSpace($office)) { $params.Office = $office }
          if (-not [string]::IsNullOrWhiteSpace($phone)) { $params.OfficePhone = $phone }
          if (-not [string]::IsNullOrWhiteSpace($desc)) { $params.Description = $desc }

          $enabledRaw = Get-CsvValue $row 'Enabled'
          if (-not [string]::IsNullOrWhiteSpace($enabledRaw)) { $params.Enabled = (ConvertTo-BoolOrDefault $enabledRaw $true) }
          $mustChangeRaw = Get-CsvValue $row 'MustChangePassword'
          if (-not [string]::IsNullOrWhiteSpace($mustChangeRaw)) { $params.ChangePasswordAtLogon = (ConvertTo-BoolOrDefault $mustChangeRaw $false) }
          $cannotChangeRaw = Get-CsvValue $row 'CannotChangePassword'
          if (-not [string]::IsNullOrWhiteSpace($cannotChangeRaw)) { $params.CannotChangePassword = (ConvertTo-BoolOrDefault $cannotChangeRaw $false) }
          $neverExpiresRaw = Get-CsvValue $row 'PasswordNeverExpires'
          if (-not [string]::IsNullOrWhiteSpace($neverExpiresRaw)) { $params.PasswordNeverExpires = (ConvertTo-BoolOrDefault $neverExpiresRaw $false) }

          Set-ADUser @params

          $newPwd = Get-CsvValue $row 'NewPassword'
          if (-not [string]::IsNullOrWhiteSpace($newPwd)) {
            $pwd = ConvertTo-SecureString $newPwd -AsPlainText -Force
            Set-ADAccountPassword @adp -Identity $identity -NewPassword $pwd -Reset -WhatIf:$simulate
          }

          $newCn = Get-CsvValue $row 'FullName'
          if (-not [string]::IsNullOrWhiteSpace($newCn)) {
            $obj = Get-ADUser @adp -Identity $identity -Properties DistinguishedName, Name
            if ($obj.Name -ne $newCn) {
              Rename-ADObject @adp -Identity $obj.DistinguishedName -NewName $newCn -WhatIf:$simulate
            }
          }

          $targetOu = Get-CsvValue $row 'OU'
          if (-not [string]::IsNullOrWhiteSpace($targetOu)) {
            if ($targetOu -match '^(?i)(OU|CN)=Domain Controllers,') { throw "No se permite mover a 'Domain Controllers'." }
            $obj = Get-ADUser @adp -Identity $identity -Properties DistinguishedName
            Move-ADObject @adp -Identity $obj.DistinguishedName -TargetPath $targetOu -WhatIf:$simulate
          }
        }
        'baja' {
          if ([string]::IsNullOrWhiteSpace($identity)) { throw "Baja requiere Identity o SamAccountName." }
          $delete = ConvertTo-BoolOrDefault (Get-CsvValue $row 'Delete') $false
          if ($delete) {
            Remove-ADUser @adp -Identity $identity -Confirm:$false -WhatIf:$simulate
          }
          else {
            Disable-ADAccount @adp -Identity $identity -WhatIf:$simulate
            $targetOu = Get-CsvValue $row 'TargetOU'
            if (-not [string]::IsNullOrWhiteSpace($targetOu)) {
              if ($targetOu -match '^(?i)(OU|CN)=Domain Controllers,') { throw "No se permite mover a 'Domain Controllers'." }
              $obj = Get-ADUser @adp -Identity $identity -Properties DistinguishedName
              Move-ADObject @adp -Identity $obj.DistinguishedName -TargetPath $targetOu -WhatIf:$simulate
            }
          }
        }
        'mover' {
          if ([string]::IsNullOrWhiteSpace($identity)) { throw "Mover requiere Identity o SamAccountName." }
          $targetOu = Get-CsvValue $row 'TargetOU'
          if ([string]::IsNullOrWhiteSpace($targetOu)) { throw "Mover requiere TargetOU." }
          if ($targetOu -match '^(?i)(OU|CN)=Domain Controllers,') { throw "No se permite mover a 'Domain Controllers'." }

          $dn = $identity
          if ($identity -notmatch '^(?i)(CN|OU)=') {
            $obj = Get-ADUser @adp -Identity $identity -Properties DistinguishedName
            $dn = $obj.DistinguishedName
          }
          Move-ADObject @adp -Identity $dn -TargetPath $targetOu -WhatIf:$simulate
        }
        default {
          throw "Acción no soportada: '$actionRaw'. Usá Alta, Modificar, Baja o Mover."
        }
      }

      $msg = if ($simulate) { "Simulado correctamente." } else { "Ejecutado correctamente." }
      $results += New-BatchResult -rowNumber $rowNumber -action $actionRaw -identity $identity -status 'OK' -message $msg
      $auditResult = if ($simulate) { 'SIMULATED' } else { 'OK' }
      Write-AuditLog -Action ("Lote CSV: $actionRaw") -Result $auditResult -Object ([string]$identity) -Detail ("Fila: " + $rowNumber)
    }
    catch {
      $results += New-BatchResult -rowNumber $rowNumber -action $actionRaw -identity $identity -status 'ERROR' -message $_.Exception.Message
      Write-AuditLog -Action ("Lote CSV: $actionRaw") -Result 'ERROR' -Object ([string]$identity) -Detail ("Fila: " + $rowNumber + " | " + $_.Exception.Message)
    }
  }

  return @($results)
}

function Save-BatchReport([object[]]$results) {
  if ($null -eq $results -or $results.Count -eq 0) { return $null }
  $dir = Join-Path $PSScriptRoot 'BatchReports'
  if (-not (Test-Path -Path $dir -PathType Container)) {
    New-Item -Path $dir -ItemType Directory -Force | Out-Null
  }
  $file = Join-Path $dir ("ADBatch_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
  $results | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
  return $file
}

function New-BatchTemplateRows {
  return @(
    [PSCustomObject]@{
      Action                = 'Alta'
      Identity              = ''
      SamAccountName        = 'jperez'
      DistinguishedName     = ''
      FirstName             = 'Juan'
      Initials              = 'P'
      LastName              = 'Perez'
      FullName              = 'Juan Perez'
      DisplayName           = 'Perez, Juan'
      UpnUser               = 'jperez'
      OU                    = 'OU=Usuarios,DC=empresa,DC=local'
      TargetOU              = ''
      Password              = 'Temporal#2026'
      NewPassword           = ''
      Email                 = 'juan.perez@empresa.local'
      Title                 = 'Analista'
      Department            = 'IT'
      Company               = 'Empresa'
      Office                = 'CABA - Piso 3'
      OfficePhone           = '+54 11 5555 0001'
      Description           = 'Alta por lote'
      Enabled               = 'true'
      MustChangePassword    = 'true'
      CannotChangePassword  = 'false'
      PasswordNeverExpires  = 'false'
      Delete                = 'false'
    }
    [PSCustomObject]@{
      Action                = 'Modificar'
      Identity              = 'jperez'
      SamAccountName        = ''
      DistinguishedName     = ''
      FirstName             = ''
      Initials              = ''
      LastName              = ''
      FullName              = ''
      DisplayName           = ''
      UpnUser               = ''
      OU                    = ''
      TargetOU              = ''
      Password              = ''
      NewPassword           = ''
      Email                 = 'juan.perez2@empresa.local'
      Title                 = 'Senior Analyst'
      Department            = 'IT'
      Company               = ''
      Office                = ''
      OfficePhone           = ''
      Description           = 'Actualizacion por lote'
      Enabled               = 'true'
      MustChangePassword    = ''
      CannotChangePassword  = ''
      PasswordNeverExpires  = ''
      Delete                = ''
    }
    [PSCustomObject]@{
      Action                = 'Baja'
      Identity              = 'jperez'
      SamAccountName        = ''
      DistinguishedName     = ''
      FirstName             = ''
      Initials              = ''
      LastName              = ''
      FullName              = ''
      DisplayName           = ''
      UpnUser               = ''
      OU                    = ''
      TargetOU              = 'OU=Bajas,DC=empresa,DC=local'
      Password              = ''
      NewPassword           = ''
      Email                 = ''
      Title                 = ''
      Department            = ''
      Company               = ''
      Office                = ''
      OfficePhone           = ''
      Description           = 'Baja por lote'
      Enabled               = ''
      MustChangePassword    = ''
      CannotChangePassword  = ''
      PasswordNeverExpires  = ''
      Delete                = 'false'
    }
    [PSCustomObject]@{
      Action                = 'Mover'
      Identity              = 'CN=Juan Perez,OU=Usuarios,DC=empresa,DC=local'
      SamAccountName        = ''
      DistinguishedName     = ''
      FirstName             = ''
      Initials              = ''
      LastName              = ''
      FullName              = ''
      DisplayName           = ''
      UpnUser               = ''
      OU                    = ''
      TargetOU              = 'OU=Soporte,DC=empresa,DC=local'
      Password              = ''
      NewPassword           = ''
      Email                 = ''
      Title                 = ''
      Department            = ''
      Company               = ''
      Office                = ''
      OfficePhone           = ''
      Description           = 'Movimiento por lote'
      Enabled               = ''
      MustChangePassword    = ''
      CannotChangePassword  = ''
      PasswordNeverExpires  = ''
      Delete                = ''
    }
  )
}

function Save-BatchTemplate([string]$path) {
  $templateRows = New-BatchTemplateRows
  $templateRows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
  return $path
}

function Get-BatchTemplateHeaders {
  return @(
    'Action', 'Identity', 'SamAccountName', 'DistinguishedName',
    'FirstName', 'Initials', 'LastName', 'FullName', 'DisplayName', 'UpnUser',
    'OU', 'TargetOU',
    'Password', 'NewPassword',
    'Email', 'Title', 'Department', 'Company', 'Office', 'OfficePhone', 'Description',
    'Enabled', 'MustChangePassword', 'CannotChangePassword', 'PasswordNeverExpires',
    'Delete'
  )
}

function Save-BatchTemplateEmpty([string]$path) {
  $headerLine = (Get-BatchTemplateHeaders) -join ','
  Set-Content -Path $path -Value $headerLine -Encoding UTF8
  return $path
}

$btnBatchBrowseCsv.Add_Click({
    try {
      $ofd = New-Object Microsoft.Win32.OpenFileDialog
      $ofd.Filter = "CSV (*.csv)|*.csv|Todos (*.*)|*.*"
      $ofd.Multiselect = $false
      if ($ofd.ShowDialog() -eq $true) {
        $txtBatchCsvPath.Text = $ofd.FileName
      }
    }
    catch { Show-Exception "Error seleccionando CSV:" $_ }
  })

$btnBatchTemplate.Add_Click({
    try {
      $sfd = New-Object Microsoft.Win32.SaveFileDialog
      $sfd.Filter = "CSV (*.csv)|*.csv|Todos (*.*)|*.*"
      $sfd.FileName = "Plantilla_AD_Lote.csv"
      if ($sfd.ShowDialog() -ne $true) { return }

      $out = Save-BatchTemplate -path $sfd.FileName
      $txtBatchCsvPath.Text = $out
      [System.Windows.MessageBox]::Show("Plantilla generada en:`n$out", "Plantilla CSV", "OK", "Information") | Out-Null
      Set-Status "✅ Plantilla CSV generada: $out" '#a6e3a1'
      Write-AuditLog -Action 'Generar Plantilla CSV' -Result 'OK' -Object $out
    }
    catch { Show-Exception "Error generando plantilla CSV:" $_ }
  })

$btnBatchTemplateEmpty.Add_Click({
    try {
      $sfd = New-Object Microsoft.Win32.SaveFileDialog
      $sfd.Filter = "CSV (*.csv)|*.csv|Todos (*.*)|*.*"
      $sfd.FileName = "Plantilla_AD_Lote_Vacia.csv"
      if ($sfd.ShowDialog() -ne $true) { return }

      $out = Save-BatchTemplateEmpty -path $sfd.FileName
      $txtBatchCsvPath.Text = $out
      [System.Windows.MessageBox]::Show("Plantilla vacía generada en:`n$out", "Plantilla CSV", "OK", "Information") | Out-Null
      Set-Status "✅ Plantilla vacía generada: $out" '#a6e3a1'
      Write-AuditLog -Action 'Generar Plantilla CSV Vacia' -Result 'OK' -Object $out
    }
    catch { Show-Exception "Error generando plantilla vacía:" $_ }
  })

$btnBatchPreview.Add_Click({
    $path = Get-RequiredTrimmedText $txtBatchCsvPath.Text "Seleccioná un CSV para preview."
    if ($null -eq $path) { return }

    $results = Invoke-BatchCsvProcess -csvPath $path -Preview
    $script:batchLastResults = @($results)
    $dgBatchResults.ItemsSource = @($results)

    $ok = @($results | Where-Object { $_.Status -eq 'OK' }).Count
    $err = @($results | Where-Object { $_.Status -eq 'ERROR' }).Count
    Set-Status "🧪 Preview CSV finalizado. OK: $ok | ERROR: $err" '#fab387'
    Write-AuditLog -Action 'Preview Lote CSV' -Result 'SIMULATED' -Object $path -Detail ("OK: $ok | ERROR: $err")
  })

$btnBatchExecute.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $path = Get-RequiredTrimmedText $txtBatchCsvPath.Text "Seleccioná un CSV para ejecutar."
    if ($null -eq $path) { return }

    if (-not $script:SimulationMode) {
      if (-not (Show-Confirm "⚠️ Se ejecutarán cambios reales en AD. ¿Continuar?")) { return }
    }

    $results = Invoke-BatchCsvProcess -csvPath $path
    $script:batchLastResults = @($results)
    $dgBatchResults.ItemsSource = @($results)

    $ok = @($results | Where-Object { $_.Status -eq 'OK' }).Count
    $err = @($results | Where-Object { $_.Status -eq 'ERROR' }).Count
    if ($script:SimulationMode) {
      Set-Status "🧪 Ejecución en simulación finalizada. OK: $ok | ERROR: $err" '#fab387'
      Write-AuditLog -Action 'Ejecutar Lote CSV' -Result 'SIMULATED' -Object $path -Detail ("OK: $ok | ERROR: $err")
    }
    else {
      Set-Status "✅ Ejecución CSV finalizada. OK: $ok | ERROR: $err" '#a6e3a1'
      Write-AuditLog -Action 'Ejecutar Lote CSV' -Result 'OK' -Object $path -Detail ("OK: $ok | ERROR: $err")
    }
  })

$btnBatchExportReport.Add_Click({
    if ($null -eq $script:batchLastResults -or $script:batchLastResults.Count -eq 0) {
      Show-Error "No hay resultados para exportar."
      return
    }
    try {
      $out = Save-BatchReport -results $script:batchLastResults
      $script:batchLastReportPath = $out
      [System.Windows.MessageBox]::Show("Reporte exportado en:`n$out", "Reporte CSV", "OK", "Information") | Out-Null
      Set-Status "✅ Reporte exportado: $out" '#a6e3a1'
      Write-AuditLog -Action 'Exportar Reporte Lote CSV' -Result 'OK' -Object $out
    }
    catch { Show-Exception "Error exportando reporte:" $_ }
  })

# ── ATRIBUTOS ────────────────────────────────────────────────────────────────
$script:currentAttrDN = $null

function Clear-AttributeResults([switch]$ClearSearchText) {
  $script:currentAttrDN = $null
  $txtAttrDN.Text = ''
  $dgAttributes.ItemsSource = @()
  if ($ClearSearchText) { $txtAttrSearch.Text = '' }
}

$btnAttrSearch.Add_Click({
    if (-not (Assert-ADModule)) { return }
    $search = Get-NormalizedText $txtAttrSearch.Text
    if ([string]::IsNullOrWhiteSpace($search)) {
      Clear-AttributeResults
      Set-Status "ℹ️ Resultados de atributos limpiados." '#89b4fa'
      return
    }

    try {
      [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
      Set-Status "Cargando atributos..." '#89b4fa'
      $adp = Get-ADParams

      $obj = $null
      try { $obj = Get-ADUser @adp -Identity $search -Properties * } catch {}
      if (-not $obj) { try { $obj = Get-ADGroup @adp -Identity $search -Properties * } catch {} }
      if (-not $obj) {
        Clear-AttributeResults
        Show-Error "No se encontró '$search' como usuario ni como grupo."
        return
      }

      $script:currentAttrDN = $obj.DistinguishedName
      $txtAttrDN.Text = "DN: $($obj.DistinguishedName)"

      $editableAttrs = @(
        'GivenName', 'Surname', 'DisplayName', 'Description', 'EmailAddress', 'Title',
        'Department', 'Company', 'Manager', 'OfficePhone', 'MobilePhone', 'StreetAddress',
        'City', 'State', 'PostalCode', 'Country', 'Office', 'HomePage', 'EmployeeID', 'EmployeeNumber',
        'Info', 'Division', 'wWWHomePage'
      )

      $attrList = foreach ($attr in $editableAttrs) {
        $val = $obj.$attr
        if ($val -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) { $val = ($val -join '; ') }

        [PSCustomObject]@{ Attribute = $attr; Value = [string]$val }
      }

      $dgAttributes.IsReadOnly = $false
      $dgAttributes.ItemsSource = @($attrList)
      Set-Status "✅ Atributos cargados para '$search'." '#a6e3a1'
    }
    catch { Show-Exception "Error cargando atributos:" $_ }
    finally { [System.Windows.Input.Mouse]::OverrideCursor = $null }
  })
Bind-EnterKeyToButton -textBox $txtAttrSearch -button $btnAttrSearch

$btnAttrClear.Add_Click({
    Clear-AttributeResults -ClearSearchText
    Set-Status "ℹ️ Búsqueda y resultados de atributos limpiados." '#89b4fa'
  })

$btnAttrSave.Add_Click({
    if (-not (Assert-ADModule)) { return }
    if (-not (Test-Guard (-not [string]::IsNullOrWhiteSpace($script:currentAttrDN)) "Primero cargá un objeto.")) { return }

    try {
      $adp = Get-ADParams
      $items = $dgAttributes.ItemsSource
      foreach ($item in $items) {
        $attrName = $item.Attribute
        $attrVal = $item.Value
        if ([string]::IsNullOrEmpty($attrVal)) {
          Set-ADObject @adp -Identity $script:currentAttrDN -Clear $attrName -ErrorAction SilentlyContinue
        }
        else {
          Set-ADObject @adp -Identity $script:currentAttrDN -Replace @{ $attrName = $attrVal } -ErrorAction SilentlyContinue
        }
      }
      [System.Windows.MessageBox]::Show("Atributos guardados correctamente.", "Éxito", "OK", "Information") | Out-Null
      Set-Status "✅ Atributos guardados." '#a6e3a1'
      Write-AuditLog -Action 'Guardar Atributos' -Result 'OK' -Object ([string]$script:currentAttrDN)
    }
    catch { Show-Exception "Error guardando atributos:" $_ }
  })

function New-StartupSplash {
  $isDark = [bool]$script:isDarkMode

  $splashBorder = if ($isDark) { '#3b4261' } else { '#b4bdd0' }
  $splashShadow = if ($isDark) { '#000000' } else { '#7d869a' }
  $gradStart = if ($isDark) { '#1f2538' } else { '#f7f9ff' }
  $gradEnd = if ($isDark) { '#14192a' } else { '#e7ecf7' }
  $brandColor = if ($isDark) { '#a9c8ff' } else { '#2d4ea3' }
  $subtitleColor = if ($isDark) { '#d2daf0' } else { '#4a5878' }
  $statusColor = if ($isDark) { '#8ea2c9' } else { '#3f4f72' }
  $progressFg = if ($isDark) { '#6aa9ff' } else { '#4f8eea' }
  $progressBg = if ($isDark) { '#2c334a' } else { '#d5deef' }
  $hintColor = if ($isDark) { '#667391' } else { '#596887' }

  $splash = New-Object System.Windows.Window
  $splash.WindowStyle = 'None'
  $splash.ResizeMode = 'NoResize'
  $splash.AllowsTransparency = $true
  $splash.Background = ConvertTo-Brush '#00000000'
  $splash.ShowInTaskbar = $false
  $splash.Topmost = $true
  $splash.SizeToContent = 'WidthAndHeight'
  $splash.WindowStartupLocation = 'CenterScreen'

  $outer = New-Object System.Windows.Controls.Border
  $outer.CornerRadius = '14'
  $outer.BorderBrush = ConvertTo-Brush $splashBorder
  $outer.BorderThickness = '1'
  $outer.Padding = '26'
  $outer.Effect = New-Object System.Windows.Media.Effects.DropShadowEffect -Property @{
    BlurRadius   = 22
    ShadowDepth  = 0
    Opacity      = 0.35
    Color        = ([System.Windows.Media.ColorConverter]::ConvertFromString($splashShadow))
  }
  $outer.Background = New-Object System.Windows.Media.LinearGradientBrush -Property @{
    StartPoint = '0,0'
    EndPoint   = '1,1'
  }
  $outer.Background.GradientStops.Add((New-Object System.Windows.Media.GradientStop -Property @{ Color = ([System.Windows.Media.ColorConverter]::ConvertFromString($gradStart)); Offset = 0 }))
  $outer.Background.GradientStops.Add((New-Object System.Windows.Media.GradientStop -Property @{ Color = ([System.Windows.Media.ColorConverter]::ConvertFromString($gradEnd)); Offset = 1 }))

  $root = New-Object System.Windows.Controls.Grid
  $root.Width = 560
  $root.Height = 280

  $stack = New-Object System.Windows.Controls.StackPanel
  $stack.VerticalAlignment = 'Center'

  $txtBrand = New-Object System.Windows.Controls.TextBlock
  $txtBrand.Text = 'AD Manager'
  $txtBrand.FontSize = 34
  $txtBrand.FontWeight = 'Bold'
  $txtBrand.Foreground = ConvertTo-Brush $brandColor
  $txtBrand.HorizontalAlignment = 'Center'
  $stack.Children.Add($txtBrand) | Out-Null

  $txtSub = New-Object System.Windows.Controls.TextBlock
  $txtSub.Text = 'Administrador de Active Directory'
  $txtSub.FontSize = 14
  $txtSub.Margin = '0,8,0,18'
  $txtSub.Foreground = ConvertTo-Brush $subtitleColor
  $txtSub.HorizontalAlignment = 'Center'
  $stack.Children.Add($txtSub) | Out-Null

  $txtDev = New-Object System.Windows.Controls.TextBlock
  $txtDev.Text = 'Develop by Matias Vaccari'
  $txtDev.FontSize = 11
  $txtDev.Margin = '0,-10,0,14'
  $txtDev.Foreground = ConvertTo-Brush $hintColor
  $txtDev.HorizontalAlignment = 'Center'
  $stack.Children.Add($txtDev) | Out-Null

  $statusText = New-Object System.Windows.Controls.TextBlock
  $statusText.Text = 'Inicializando...'
  $statusText.FontSize = 13
  $statusText.Foreground = ConvertTo-Brush $statusColor
  $statusText.HorizontalAlignment = 'Center'
  $statusText.Margin = '0,0,0,10'
  $stack.Children.Add($statusText) | Out-Null

  $pb = New-Object System.Windows.Controls.ProgressBar
  $pb.Width = 420
  $pb.Height = 10
  $pb.IsIndeterminate = $false
  $pb.Minimum = 0
  $pb.Maximum = 100
  $pb.Value = 0
  $pb.Foreground = ConvertTo-Brush $progressFg
  $pb.Background = ConvertTo-Brush $progressBg
  $pb.BorderBrush = ConvertTo-Brush $splashBorder
  $pb.BorderThickness = '1'
  $stack.Children.Add($pb) | Out-Null

  $txtHint = New-Object System.Windows.Controls.TextBlock
  $txtHint.Text = 'Conectando servicios y validando dominio'
  $txtHint.FontSize = 11
  $txtHint.Margin = '0,10,0,0'
  $txtHint.Foreground = ConvertTo-Brush $hintColor
  $txtHint.HorizontalAlignment = 'Center'
  $stack.Children.Add($txtHint) | Out-Null

  $root.Children.Add($stack) | Out-Null
  $outer.Child = $root
  $splash.Content = $outer

  return [PSCustomObject]@{
    Window = $splash
    Status = $statusText
    Progress = $pb
  }
}

function Set-StartupSplashProgress([object]$splashUi, [int]$percent, [string]$message = '') {
  if ($null -eq $splashUi) { return }
  if ($splashUi.Status -and (-not [string]::IsNullOrWhiteSpace($message))) {
    $splashUi.Status.Text = $message
  }
  if ($splashUi.Progress) {
    $clamped = [Math]::Max(0, [Math]::Min(100, $percent))
    $splashUi.Progress.Value = $clamped
  }
  [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
}

function Step-StartupSplashProgress([object]$splashUi, [int]$targetPercent, [string]$message = '') {
  if ($null -eq $splashUi -or $null -eq $splashUi.Progress) { return }
  $target = [Math]::Max(0, [Math]::Min(100, $targetPercent))
  $current = [int][Math]::Floor($splashUi.Progress.Value)
  if ($current -ge $target) {
    Set-StartupSplashProgress -splashUi $splashUi -percent $current -message $message
    return
  }
  for ($i = ($current + 1); $i -le $target; $i++) {
    Set-StartupSplashProgress -splashUi $splashUi -percent $i -message $message
    Start-Sleep -Milliseconds 9
  }
}

function Update-StartupSplashStatus([object]$splashUi, [string]$message) {
  if ($null -eq $splashUi -or $null -eq $splashUi.Status) { return }
  Set-StartupSplashProgress -splashUi $splashUi -percent ([int][Math]::Floor($splashUi.Progress.Value)) -message $message
}

function Complete-StartupSplash([object]$splashUi, [string]$message = 'Listo.') {
  if ($null -eq $splashUi) { return }
  Step-StartupSplashProgress -splashUi $splashUi -targetPercent 100 -message $message
  [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
  Start-Sleep -Milliseconds 350
}

function Invoke-StartupInitializationWithSplash([scriptblock]$StartupAction) {
  $minimumVisibleMs = 1400
  $splashUi = New-StartupSplash
  $startupTick = [Environment]::TickCount
  try {
    $splashUi.Window.Show()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
    & $StartupAction $splashUi

    if ($splashUi -and $splashUi.Progress) {
      $finalMessage = if ($splashUi.Status -and (-not [string]::IsNullOrWhiteSpace([string]$splashUi.Status.Text))) {
        [string]$splashUi.Status.Text
      }
      else {
        'Listo.'
      }
      Step-StartupSplashProgress -splashUi $splashUi -targetPercent 100 -message $finalMessage
    }

    $elapsed = [Math]::Abs([Environment]::TickCount - $startupTick)
    if ($elapsed -lt $minimumVisibleMs) {
      Start-Sleep -Milliseconds ($minimumVisibleMs - $elapsed)
    }
  }
  finally {
    if ($splashUi -and $splashUi.Window -and $splashUi.Window.IsVisible) {
      $splashUi.Window.Close()
    }
  }
}

# ── MOSTRAR ──────────────────────────────────────────────────────────────────
Invoke-StartupInitializationWithSplash {
  param($splashUi)

  Step-StartupSplashProgress -splashUi $splashUi -targetPercent 10 -message 'Inicializando componentes...'
  Step-StartupSplashProgress -splashUi $splashUi -targetPercent 22 -message 'Cargando configuracion de auditoria...'
  try { [void](Load-AuditSettingsFromFile) } catch {}
  Update-AuditGuiFromSettings
  Update-AuditModeUiState
  $diagInit = Get-AuditValidationResult
  Show-AuditDiagnostics -lines $diagInit.Lines

  Step-StartupSplashProgress -splashUi $splashUi -targetPercent 36 -message 'Aplicando tema visual...'
  Apply-Theme

  if (-not $script:ADModuleAvailable) {
    Complete-StartupSplash -splashUi $splashUi -message 'Finalizando...'
    Set-Status "⚠️ Módulo ActiveDirectory no disponible. Instalá RSAT para usar las funciones AD." '#fab387'
  }
  else {
    Step-StartupSplashProgress -splashUi $splashUi -targetPercent 50 -message 'Preparando validacion multidominio...'
    Set-Status "Verificando conectividad de dominios..." '#89b4fa'
    if (Test-DCConnectivityAtStartup -TimeoutSeconds 4 -SplashUi $splashUi) {
      Step-StartupSplashProgress -splashUi $splashUi -targetPercent 95 -message 'Cargando estructura de dominio...'
      Refresh-DomainCombo
      if (($cbDomainSelect.SelectedIndex -lt 0) -and ($cbDomainSelect.Items.Count -gt 0)) { $cbDomainSelect.SelectedIndex = 0 }
      Load-OUTreeForControl $tvMoveSourceOUs
      Load-OUTreeForControl -treeView $tvMoveTargetOUs -HideProtectedContainers
      $domainLabel = if ($global:activeDomain) { [string]$global:activeDomain } else { [string]$script:LocalDomainDns }
      Set-Status "✅ Conexión a AD verificada. Dominio activo: $domainLabel" '#a6e3a1'
      Complete-StartupSplash -splashUi $splashUi -message 'Listo.'
      Show-StartupDomainFailuresSummary
    }
    else {
      Complete-StartupSplash -splashUi $splashUi -message 'Sin conexion al DC. Iniciando en modo limitado...'
      Show-FriendlyDcUnavailableMessage
      Show-StartupDomainFailuresSummary
    }
  }
}

$window.ShowDialog() | Out-Null






