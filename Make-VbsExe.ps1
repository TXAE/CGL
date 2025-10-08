# Make-VbsExe.ps1
# Creates a standalone EXE that runs your VBScript via cscript.exe, forwards all output,
# embeds ALL .vbs files (main + helpers), extracts them to a temp folder **as raw bytes** (no re-encoding),
# and sets the working directory so relative opens (e.g., "SAP Login.vbs") work.
# Includes auto-pause on double-click and --pause switch when running the final EXE.
# Handles locked EXE copies with retry/fallback.
# Accepts and ignores a stray "-and" switch to avoid invocation typos.
# NOTE: This script intentionally avoids the PowerShell "-and" operator to prevent confusion.
#
# EXAMPLE USAGE:
# .\Make-VbsExe.ps1 `
#   -VbsPath "C:\Users\d363973\OneDrive - Cargill Inc\Documents\Cargill scripting\SAP\read Schichtbuch excel etc.vbs" `
#   -SupportVbs "C:\Users\d363973\OneDrive - Cargill Inc\Documents\Cargill scripting\SAP\SAP Login.vbs" `
#   -Runtime win-x86 `
#   -OutDir "C:\Users\d363973\OneDrive - Cargill Inc\2Do\SAP_Script\"

[CmdletBinding()]
param(
  # Path to the main VBScript (will be copied into the project and run as -MainScriptName)
  [Parameter(Mandatory = $true)]
  [string] $VbsPath,

  # Additional .vbs files to include (e.g., "SAP Login.vbs"); copied into the project root.
  [string[]] $SupportVbs = @(),

  # Name to give the main script inside the EXE payload (must be a file name only).
  [string] $MainScriptName = "script.vbs",

  # Target runtime; use win-x86 for legacy 32-bit COM (common in SAP GUI), win-x64 otherwise.
  [ValidateSet("win-x64","win-x86")]
  [string] $Runtime = "win-x86",

  # Output folder where the final EXE is copied.
  [string] $OutDir = ".\dist",

  # Safety: ignore a stray "-and" in invocation (e.g., user typed -and between params)
  [switch] $and
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Test-FileLocked {
  param([Parameter(Mandatory=$true)][string]$Path)
  try {
    $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
    $fs.Close()
    return $false
  } catch [System.IO.IOException] {
    return $true
  }
}

function Copy-WithRetryOrTimestamp {
  param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Dest,
    [int]$Retries = 10,
    [int]$DelayMs = 500
  )
  for ($i = 0; $i -lt $Retries; $i++) {
    try {
      Copy-Item -LiteralPath $Source -Destination $Dest -Force
      return $Dest
    } catch [System.IO.IOException] {
      Start-Sleep -Milliseconds $DelayMs
    }
  }
  # Fallback: copy to a timestamped file to avoid lock contention
  $dir  = Split-Path -Parent $Dest
  $name = [System.IO.Path]::GetFileNameWithoutExtension($Dest)
  $ext  = [System.IO.Path]::GetExtension($Dest)
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $alt = Join-Path $dir "$name-$stamp$ext"
  Copy-Item -LiteralPath $Source -Destination $alt -Force
  Write-Warning "Could not overwrite '$Dest' (file in use). Copied to '$alt' instead."
  return $alt
}

# --- [0] Validate inputs -----------------------------------------------------
Write-Host "[0/8] Validating inputs..."
if (-not (Test-Path -LiteralPath $VbsPath)) {
  throw "Main VBScript not found: $VbsPath"
}
foreach ($s in $SupportVbs) {
  if (-not (Test-Path -LiteralPath $s)) {
    throw "Support VBScript not found: $s"
  }
}
# Require a pure filename (no path) for the main script inside the payload.
if ($MainScriptName -match '[\\\/]' -or [string]::IsNullOrWhiteSpace($MainScriptName)) {
  throw "-MainScriptName must be a file name without path (e.g., 'script.vbs'). Provided: '$MainScriptName'"
}

# Normalize paths
$VbsPath    = (Resolve-Path -LiteralPath $VbsPath).Path
$SupportVbs = $SupportVbs | ForEach-Object { (Resolve-Path -LiteralPath $_).Path }

# Derive ProjectName from the last segment of VbsPath (filename without extension), sanitized for .NET project name
$ProjectNameRaw = [System.IO.Path]::GetFileNameWithoutExtension($VbsPath)
if ([string]::IsNullOrWhiteSpace($ProjectNameRaw)) {
  throw "-VbsPath must point to a file."
}
# Replace any char not A-Za-z0-9_.- with underscore to keep it safe for project/exe names
$ProjectName = ($ProjectNameRaw -replace '[^A-Za-z0-9_.-]', '_').Trim('_')

if ([string]::IsNullOrWhiteSpace($ProjectName)) { $ProjectName = 'VbsRunner' }

$ProjectDir = Join-Path -Path (Get-Location) -ChildPath $ProjectName

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
  throw "dotnet CLI not found in PATH. Install .NET SDK 8+ and re-open your terminal. (Try 'dotnet --info')"
}

# --- [1] Create project ------------------------------------------------------
Write-Host "[1/8] Creating project '$ProjectName'..."
if (Test-Path $ProjectDir) { Remove-Item $ProjectDir -Recurse -Force }
dotnet new console -n $ProjectName | Out-Null

# --- [2] Write .csproj embedding *all* .vbs files with exact file names -----
Write-Host "[2/8] Writing project file (.csproj) to embed *.vbs..."
@"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <!-- Embed all .vbs files that are in the project root.
       LogicalName ensures the manifest resource name equals the actual file name. -->
  <ItemGroup>
    <EmbeddedResource Include="*.vbs">
      <LogicalName>%(Filename)%(Extension)</LogicalName>
    </EmbeddedResource>
  </ItemGroup>
</Project>
"@ | Set-Content -Path (Join-Path $ProjectDir "$ProjectName.csproj") -Encoding UTF8

# --- [3] Write Program.cs (multi-VBS extraction, WorkingDirectory=temp, auto-pause) ---
Write-Host "[3/8] Writing Program.cs (multi-VBS extraction + auto-pause, RAW BYTES extraction)..."
$csharp = @'
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

internal static class Program
{
    [DllImport("kernel32.dll")]
    private static extern uint GetConsoleProcessList(uint[] processList, uint processCount);

    static async Task<int> Main(string[] args)
    {
        bool forcePause = args.Any(a => string.Equals(a, "--pause", StringComparison.OrdinalIgnoreCase));
        args = args.Where(a => !string.Equals(a, "--pause", StringComparison.OrdinalIgnoreCase)).ToArray();

        var asm = Assembly.GetExecutingAssembly();
        var tempDir = Path.Combine(Path.GetTempPath(), $"VbsRunner_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);

        string mainScriptName = "__MAIN_SCRIPT_NAME__";
        string mainScriptPath = Path.Combine(tempDir, mainScriptName);

        int exitCode = 0;
        try
        {
            // 1) Extract ALL embedded .vbs resources to temp, preserving their file names **as raw bytes**.
            var resources = asm.GetManifestResourceNames()
                               .Where(n => n.EndsWith(".vbs", StringComparison.OrdinalIgnoreCase))
                               .ToArray();
            if (resources.Length == 0)
            {
                Console.Error.WriteLine("No embedded .vbs resources were found. Ensure your .csproj embeds *.vbs.");
                return 2;
            }

            foreach (var res in resources)
            {
                string fileName = res; // LogicalName set to the file name, so res == "SomeFile.vbs"
                string outPath = Path.Combine(tempDir, fileName);
                using var inStream = asm.GetManifestResourceStream(res)!;
                using var outStream = File.Create(outPath);
                await inStream.CopyToAsync(outStream).ConfigureAwait(false);
            }

            if (!File.Exists(mainScriptPath))
            {
                Console.Error.WriteLine($"Main script '{mainScriptName}' was not embedded or copied.");
                return 2;
            }

            // Helpful for debugging where files landed:
            Console.WriteLine($"[VbsRunner] Payload folder: {tempDir}");

            // 2) Find cscript.exe
            var cscriptPath = FindCScript();
            if (cscriptPath is null)
            {
                Console.Error.WriteLine("Could not locate cscript.exe on this system.");
                return 3;
            }

            // 3) Run cscript with the temp folder as WorkingDirectory so relative paths resolve
            var psi = new ProcessStartInfo
            {
                FileName = cscriptPath,
                Arguments = BuildArguments(mainScriptPath, args),
                WorkingDirectory = tempDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = false,
                WindowStyle = ProcessWindowStyle.Normal
            };

            using var proc = new Process { StartInfo = psi, EnableRaisingEvents = true };
            proc.OutputDataReceived += (_, e) => { if (e.Data != null) Console.Out.WriteLine(e.Data); };
            proc.ErrorDataReceived  += (_, e) => { if (e.Data != null) Console.Error.WriteLine(e.Data); };

            Console.CancelKeyPress += (_, e) =>
            {
                try { if (!proc.HasExited) proc.Kill(entireProcessTree: true); } catch { }
                e.Cancel = true;
            };

            if (!proc.Start())
            {
                Console.Error.WriteLine("Failed to start cscript.exe.");
                exitCode = 4;
            }
            else
            {
                proc.BeginOutputReadLine();
                proc.BeginErrorReadLine();
                await proc.WaitForExitAsync();
                exitCode = proc.ExitCode;
            }
            return exitCode;
        }
        finally
        {
            try
            {
                if (Environment.GetEnvironmentVariable("KEEP_VBSRUNNER_TEMP") != "1" && Directory.Exists(tempDir))
                    Directory.Delete(tempDir, recursive: true);
            }
            catch { /* ignore cleanup issues */ }

            // Keep window open if double-clicked (no parent console), or if --pause was passed.
            if (forcePause || StartedInOwnConsole())
            {
                Console.WriteLine();
                Console.Write("Press Enter to exit...");
                try { Console.ReadLine(); } catch { }
            }
        }
    }

    private static string BuildArguments(string vbsPath, string[] args)
        => $"//nologo \"{vbsPath}\" {string.Join(" ", args.Select(Quote))}";

    private static string Quote(string s)
        => string.IsNullOrEmpty(s) ? "\"\"" :
           (s.Contains(' ') || s.Contains('\t') || s.Contains('\"'))
           ? "\"" + s.Replace("\"", "\\\"") + "\""
           : s;

    private static string? FindCScript()
    {
        var sysDir = Environment.SystemDirectory; // 64-bit self: System32 (64-bit); 32-bit self: SysWOW64 (32-bit)
        var candidate = Path.Combine(sysDir, "cscript.exe");
        if (File.Exists(candidate)) return candidate;
        return "cscript.exe"; // let PATH resolve if needed
    }

    private static bool StartedInOwnConsole()
    {
        try
        {
            uint[] list = new uint[16];
            uint count = GetConsoleProcessList(list, (uint)list.Length);
            return count <= 1;
        }
        catch { return false; }
    }
}
'@

# Replace placeholder with the requested main script file name (must be a plain file name).
$csharp = $csharp.Replace('__MAIN_SCRIPT_NAME__', $MainScriptName)
$csharp | Set-Content -Path (Join-Path $ProjectDir "Program.cs") -Encoding UTF8

# --- [4] Copy VBS files into the project root --------------------------------
Write-Host "[4/8] Copying your VBScript files into the project..."
# Copy main VBScript under the desired name (e.g., script.vbs)
Copy-Item -LiteralPath $VbsPath -Destination (Join-Path $ProjectDir $MainScriptName) -Force

# Copy any support .vbs files, preserving their names
foreach ($s in $SupportVbs) {
  $leaf = Split-Path -Leaf $s
  if ($leaf -ieq $MainScriptName) { continue } # don't overwrite main
  Copy-Item -LiteralPath $s -Destination (Join-Path $ProjectDir $leaf) -Force
}

# --- [5] Restore --------------------------------------------------------------
Write-Host "[5/8] Restoring packages (if needed)..."
Push-Location $ProjectDir
dotnet restore -v minimal

# --- [6] Publish --------------------------------------------------------------
Write-Host "[6/8] Publishing single-file, self-contained ($Runtime)..."
dotnet publish -v minimal -c Release -r $Runtime -p:PublishSingleFile=true --self-contained true
Pop-Location

# --- [7] Collect output with retry/fallback ----------------------------------
Write-Host "[7/8] Collecting output..."
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$publishExe = Join-Path $ProjectDir "bin\Release\net8.0\$Runtime\publish\$ProjectName.exe"
$destExe    = Join-Path $OutDir "$ProjectName.exe"

# Avoid using "-and" here; use nested if for clarity and compat.
if (Test-Path $destExe) {
  if (Test-FileLocked -Path $destExe) {
    Write-Warning "'$destExe' appears to be in use. Will retry for a short while, then fall back to a timestamped copy."
  }
}
$finalCopied = Copy-WithRetryOrTimestamp -Source $publishExe -Dest $destExe

# --- [8] Done ----------------------------------------------------------------
$finalPath = (Resolve-Path $finalCopied).Path
Write-Host "[8/8] Done. Your EXE:" $finalPath
Write-Host "Tip: Run with --pause to force a pause at the end, or set KEEP_VBSRUNNER_TEMP=1 to keep extracted files."
