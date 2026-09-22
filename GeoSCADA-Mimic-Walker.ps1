# GeoSCADA-Mimic-Walker.ps1
# Copyright (c) 2026  Adam Woodland
#
# Licensed under the MIT License. You may obtain a copy of the License in the LICENSE file
# distributed with this software, or at <https://opensource.org/licenses/MIT>.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
# BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT.

<#
.SYNOPSIS
    Walks one or more Geo SCADA mimics via the web server (https://<host>/db/<Mimic>?content),
    recursing through every embedded mimic, and generates an interactive HTML file showing
    the mimic -> layer -> object hierarchy together with a complexity/cost score per mimic.

.DESCRIPTION
    Requesting a mimic from the Geo SCADA web server with the "?content" query returns the
    mimic definition as XML: <Mimic> containing <Param>, <Layer>, <Script> and mimic-level
    <Animation> elements; each <Layer> contains drawing objects (Polyline, Text, Button,
    Bitmap, Group, EmbeddedMimic, Graph, List, AlarmList ...) and each object carries its own
    <Animation> elements and optional <PickMenu>.

    The script:
      1. Logs on to the web server (application-level logon form) if required.
      2. Fetches the root mimic(s) and every <EmbeddedMimic> reachable from them (relative
         references such as SCX:////CMimic/..Pumps.Display.Symbol are resolved against the
         parent mimic's full name; each leading '.' strips one level).
      3. Records layers, objects (recursing into Groups), animations (classified as simple /
         object-method / historic / indirect / SQL), pick-menu actions, scripts and parameters.
      4. Scores each mimic using the "Guide to Mimic Complexity" table (editable at the top
         of this file) - both the mimic's OWN score and its TOTAL score including the
         complexity of every embedded mimic instance.
      5. Writes a self-contained HTML report: a summary table of every mimic reached, a
         lazily-rendered collapsible tree, search, and a details panel per mimic / object.

    Only built-in PowerShell / .NET features are used - no third-party modules. HTTP is done by a
    small C# helper compiled in-process so certificate handling is per request and per host.

    Security notes: https is required unless -AllowHttp; certificate validation is on unless
    -SkipCertificateCheck (scoped to the -BaseUrl host); the report omits the server address and
    user name unless -IncludeServerLinks; -OmitExpressions strips expression/script text; the
    script exits with code 2 if any mimic could not be fetched (or stops at once with
    -FailOnFetchError); response size, mimic count, object count and nesting depth are capped.

.PARAMETER Mimic
    Full name(s) of the mimic(s) to start from, e.g. 'Sewage.Site.Display.Overview'.

.PARAMETER Group
    Instead of (or as well as) -Mimic, walk every mimic found under this database group
    (recursing into sub-groups via the web server's /db/<Group> listing pages).

.PARAMETER BaseUrl
    Root URL of the Geo SCADA web server. Defaults to https://localhost/

.PARAMETER Credential
    PSCredential for the Geo SCADA application logon. If omitted, -Username/-Password are
    used; if those are also omitted the script tries anonymous (guest) access first and prompts
    for credentials only if the server demands a logon.

.PARAMETER Username
.PARAMETER Password
    Plain-text alternative to -Credential (handy for scripted runs; prefer -Credential).

.PARAMETER Guest
    Do not log on up front: try the web server anonymously and only prompt for credentials
    if it demands a logon. Without this switch the script ALWAYS logs on explicitly (prompting
    if no credential was given) and logs off again when it finishes.

    Why: the Geo SCADA web server keeps a short-lived session keyed on the client address
    after any logon, so for roughly a minute a request with no cookies is still accepted. In
    -Guest mode a run started soon after another can therefore proceed under the previous
    run's session without prompting; the script warns when that happens.

.PARAMETER OutputPath
    Where to write the HTML. Defaults to .\MimicTree.html next to this script.

.PARAMETER CsvPath
    Optional: also write a CSV of per-mimic statistics and scores.

.PARAMETER MaxDepth
    Maximum embedded-mimic nesting depth to follow (default 25).

.PARAMETER SkipCertificateCheck
    Accept an untrusted / self-signed certificate FROM THE -BaseUrl HOST ONLY. Off by default;
    certificates are validated normally. When set, the exception is scoped to requests this
    script makes to that host (nothing process-wide is changed) - but it still removes
    protection against a man-in-the-middle on that connection, so prefer installing the
    server's certificate as trusted.

.PARAMETER AllowHttp
    Permit a plain http:// -BaseUrl. Refused by default because the logon posts credentials.

.PARAMETER IncludeServerLinks
    Embed the server address in the report so mimic and object names link back to the web
    server. Off by default so the report does not disclose the server address.

.PARAMETER OmitExpressions
    Leave animation expressions, pick-menu action text and script library names out of the
    report (counts, classes and scores are kept). Use when the report will leave your team.

.PARAMETER FailOnFetchError
    Stop at the first mimic that cannot be fetched instead of continuing with a partial
    report. Without it the script finishes, warns that the report is incomplete and exits
    with code 2.

.PARAMETER MaxMimics
    Maximum number of distinct mimics to fetch (default 5000). Further mimics are reported
    as unavailable.

.PARAMETER MaxResponseMB
    Maximum size of a single web response in MB (default 50). Larger responses are discarded.

.PARAMETER MaxObjects
    Maximum total number of drawing objects to parse across the run (default 1000000).

.PARAMETER MaxGroupDepth
    Maximum nesting depth of Groups within a mimic to descend (default 50).

.PARAMETER AcceptDisclaimer
    Accept the usage disclaimer non-interactively (required in non-interactive sessions).

.EXAMPLE
    .\GeoSCADA-Mimic-Walker.ps1 -BaseUrl https://192.168.229.139/ -Mimic 'Sewage.CLM.CLM.SPS.081.Display.Overview' -Username adamwoodland -Password 'secret'

.EXAMPLE
    .\GeoSCADA-Mimic-Walker.ps1 -BaseUrl https://scada/ -Group 'Sewage.CLM.CLM.SPS.081.Display' -Credential (Get-Credential) -CsvPath .\mimics.csv -AcceptDisclaimer

.NOTES
    DISCLAIMER
    This script is provided "AS IS", without warranty of any kind, express or implied. In no
    event shall the author be liable for any claim, damages or other liability arising from,
    out of or in connection with the script or its use. It connects to the Geo SCADA web server
    to read mimic definitions and writes an HTML file; with -SkipCertificateCheck it accepts an
    untrusted certificate from that one host. It is NOT certified for production SCADA environments. Do
    NOT run it against a production or safety-critical system without first reviewing the code
    and testing it on a representative non-production system. You run it at your own risk and
    are responsible for compliance with your own change-control and security policies.

    Complexity scoring follows the Schneider Electric community article "Guide to Mimic
    Complexity" (https://community.se.com/t5/Geo-SCADA-Knowledge-Base/Guide-to-Mimic-Complexity/ba-p/278933).
    Animation classes are inferred from the expression text, so treat scores as a guide.
#>
[CmdletBinding()]
param(
    [string[]]$Mimic,
    [string]$Group,
    [string]$BaseUrl = 'https://localhost/',
    [System.Management.Automation.PSCredential]$Credential,
    [string]$Username,
    [string]$Password,
    [string]$OutputPath,
    [string]$CsvPath,
    [int]$MaxDepth = 25,
    [switch]$SkipCertificateCheck,
    [switch]$AllowHttp,
    [switch]$Guest,
    [switch]$IncludeServerLinks,
    [switch]$OmitExpressions,
    [switch]$FailOnFetchError,
    [int]$MaxMimics = 5000,
    [int]$MaxResponseMB = 50,
    [int]$MaxObjects = 1000000,
    [int]$MaxGroupDepth = 50,
    [switch]$AcceptDisclaimer
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web

# =============================================================================================
#  COSTING TABLE  -  edit these values to change how complexity is scored
#  Source: "Guide to Mimic Complexity" (Geo SCADA Knowledge Base, community.se.com)
# =============================================================================================
#
#  Keyed by the XML element name of the object as it appears in the ?content XML.
#    Base            - points for the element itself
#    PerExtraSegment - points per line/pipe segment beyond the first (Polyline / Pipe)
#    PerKB           - points per KB of embedded image data (Bitmap)
#    AddAnimations   - $true  => add the animation scores below to this element
#                      $false => the guide gives a flat figure with no animation uplift
#    PerParameter    - points per parameter passed to an embedded mimic
#    IfSqlAnimated   - extra points if the List's SQL query itself is animated
#
$CostTable = [ordered]@{
    #  XML element              Guide row                                         Points
    'Polyline'      = @{ Base = 1;   PerExtraSegment = 0.1; AddAnimations = $true  }  # Primitive graphic element  1 (+0.1 per additional line segment)
    'Text'          = @{ Base = 1;   AddAnimations = $true  }                         # Primitive graphic element  1
    'Button'        = @{ Base = 1;   AddAnimations = $true  }                         # Primitive graphic element  1
    'Rectangle'     = @{ Base = 1;   AddAnimations = $true  }                         # Primitive graphic element  1
    'Ellipse'       = @{ Base = 1;   AddAnimations = $true  }                         # Primitive graphic element  1
    'Arc'           = @{ Base = 1;   AddAnimations = $true  }                         # Primitive graphic element  1
    'Symbol'        = @{ Base = 1;   AddAnimations = $true  }                         # Primitive graphic element  1
    'Pipe'          = @{ Base = 2;   PerExtraSegment = 0.2; AddAnimations = $true  }  # Pipe                        2 (+0.2 per additional pipe segment)
    'Bitmap'        = @{ Base = 0;   PerKB = 5;             AddAnimations = $true  }  # Bitmap image                5 * file size in KB
    'Image'         = @{ Base = 0;   PerKB = 5;             AddAnimations = $true  }  # Bitmap image (alt. name)    5 * file size in KB
    'Pie'           = @{ Base = 20;  AddAnimations = $true  }                         # Pie chart                  20
    'PieChart'      = @{ Base = 20;  AddAnimations = $true  }                         # Pie chart (alt. name)      20
    'List'          = @{ Base = 30;  IfSqlAnimated = 50;    AddAnimations = $true  }  # SQL list                   30 (+50 if SQL query is animated)
    'EmbeddedMimic' = @{ Base = 30;  PerParameter = 5;      AddAnimations = $true; ParametersAreAnimations = $false }  # Embedded mimic  30 + complexity of embedded mimic + 5 per parameter (+ each binding as an animation too if ParametersAreAnimations)
    'AlarmList'     = @{ Base = 50;  AddAnimations = $true  }                         # Alarm list                 50
    'Graph'         = @{ Base = 100; AddAnimations = $false }                         # Embedded trend            100
    'Trend'         = @{ Base = 100; AddAnimations = $false }                         # Embedded trend (alt.)     100
    'XYPlot'        = @{ Base = 100; AddAnimations = $false }                         # Embedded XY plot          100
    'XYZPlot'       = @{ Base = 100; AddAnimations = $false }                         # Embedded XYZ plot         100
    'Dynagraph'     = @{ Base = 100; AddAnimations = $false }                         # Embedded dynagraph        100
    'RemoteImage'   = @{ Base = 100; AddAnimations = $false }                         # Remote image              100
    'Group'         = @{ Base = 0;   AddAnimations = $true  }                         # Group - container only; its members are scored individually
    # Non-object rows
    'Layer'         = @{ Base = 10;  AddAnimations = $true  }                         # Layer                      10
    'Mimic'         = @{ Base = 0;   AddAnimations = $true  }                         # Mimic-level animations (not in the guide; animations only)
    # Anything not listed above
    '_Default'      = @{ Base = 1;   AddAnimations = $true  }                         # Unknown element type - treated as a primitive
}

# Points per animation, by inferred class (see Get-AnimationClass for the inference rules).
$AnimationCosts = [ordered]@{
    'Simple'   = 5     # Simple tag animation      - "Point.CurrentValue"
    'Method'   = 15    # Object method animation   - "Point.Method( args )"
    'Historic' = 20    # Historic animation        - "Point;Aggregate;Time;Interval" or "Point:Offset:Period:Max"
    'Indirect' = 50    # Indirect animation        - [ expression ] tag
    'Sql'      = 200   # SQL animation             - SELECT ... FROM ... in the expression
}

# Extras NOT in the guide (all default to 0 so the guide's numbers are reproduced exactly).
$ExtraCosts = [ordered]@{
    'ScriptPresent' = 0    # points if the mimic has a <Script> block
    'ScriptLibrary' = 0    # points per script library referenced by the mimic
    'ScriptLine'    = 0    # points per line of script (libraries included)
    'PickMenuAction'= 0    # points per pick-menu action on an object
}

# Score bands (checked in order; a mimic falls into the first band whose limit it is below).
$ComplexityBands = @(
    @{ Name = 'Simple';  Below = 1000  }
    @{ Name = 'Medium';  Below = 5000  }
    @{ Name = 'Complex'; Below = [double]::PositiveInfinity }
)
# ======================================  end of costing table  ================================

# Resolve the output path here (not in the param default) so it works even when
# $PSScriptRoot is empty - e.g. running a selection (F8) or dot-sourcing.
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot }
             elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
             else { (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $scriptDir 'MimicTree.html' }

if (-not $Mimic -and -not $Group) {
    throw 'Specify at least one -Mimic full name and/or a -Group to walk.'
}

# --- Usage disclaimer -----------------------------------------------------
function Confirm-Disclaimer {
    param([switch]$Accepted)

    $disclaimer = @'
------------------------------------------------------------------------------
 GeoSCADA-Mimic-Walker.ps1  Copyright (c) 2026  Adam Woodland
 Licensed under the MIT License. This is free software, and you are welcome to
 redistribute it under those conditions; it comes with ABSOLUTELY NO WARRANTY.

 DISCLAIMER
 This script is provided "AS IS", WITHOUT WARRANTY OF ANY KIND, express or
 implied. The author accepts no liability for any damages arising from its use.
 It connects to the Geo SCADA web server to read mimic definitions and writes
 an HTML file; with -SkipCertificateCheck it accepts an untrusted certificate
 from that one host. It is NOT certified for production SCADA systems. Do NOT run it
 against a production or safety-critical system without first reviewing the
 code and testing on a representative non-production system. You run it at
 your own risk and remain responsible for your own change-control and security
 policies.
------------------------------------------------------------------------------
'@

    if ($Accepted) { Write-Verbose 'Disclaimer accepted via -AcceptDisclaimer.'; return $true }

    $canPrompt = [Environment]::UserInteractive
    try { if ([Console]::IsInputRedirected) { $canPrompt = $false } } catch { }

    if (-not $canPrompt) {
        Write-Host $disclaimer -ForegroundColor Yellow
        throw 'Disclaimer not accepted (non-interactive session). Re-run with -AcceptDisclaimer to confirm acceptance.'
    }

    Write-Host ''
    Write-Host $disclaimer -ForegroundColor Yellow
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'I accept the terms and have tested appropriately.'
    $no  = New-Object System.Management.Automation.Host.ChoiceDescription '&No',  'Do not run.'
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@($yes, $no)
    try {
        $decision = $Host.UI.PromptForChoice('Disclaimer', 'Do you accept these terms and confirm you have tested appropriately?', $choices, 1)
    } catch {
        throw 'Disclaimer not accepted (host could not prompt). Re-run with -AcceptDisclaimer to confirm acceptance.'
    }
    if ($decision -ne 0) { Write-Warning 'Disclaimer not accepted. Exiting without running.'; return $false }
    return $true
}

# --- URL sanity ---------------------------------------------------------------
$BaseUrl = $BaseUrl.TrimEnd('/') + '/'
$baseUri = $null
if (-not [Uri]::TryCreate($BaseUrl, [UriKind]::Absolute, [ref]$baseUri)) { throw "-BaseUrl '$BaseUrl' is not a valid absolute URL." }
if ($baseUri.Scheme -eq 'http' -and -not $AllowHttp) {
    throw "-BaseUrl uses plain http, which would send the logon credentials unencrypted. Use https, or pass -AllowHttp if you accept that."
}
if ($baseUri.Scheme -notin @('http', 'https')) { throw "-BaseUrl must be http(s)." }

# --- HTTP layer -----------------------------------------------------------
# A small .NET helper (compiled in-process) is used instead of Invoke-WebRequest so that:
#   * certificate validation can be relaxed for ONE host, per request, via
#     HttpWebRequest.ServerCertificateValidationCallback - nothing process-wide is changed;
#   * response bodies are read with a hard size cap;
#   * behaviour is identical on Windows PowerShell 5.1 and PowerShell 7+.
$script:IsCore = $PSVersionTable.PSVersion.Major -ge 6
if (-not ([System.Management.Automation.PSTypeName]'GeoScadaHttp').Type) {
    $src = @"
using System;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
public class GeoScadaHttp {
    public int Status; public string Text = ""; public string Error; public string ContentType;
    public static CookieContainer Cookies = new CookieContainer();
    public static string TrustedHost = null;          // host whose certificate errors are tolerated (null = none)
    public static long MaxBytes = 50L * 1024 * 1024;
    public static int TimeoutMs = 60000;
    public static GeoScadaHttp Send(string url, string method, string body) {
        var r = new GeoScadaHttp();
        try {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = method; req.CookieContainer = Cookies; req.Timeout = TimeoutMs; req.ReadWriteTimeout = TimeoutMs;
            req.AllowAutoRedirect = true; req.UserAgent = "GeoSCADA-Mimic-Walker";
            req.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(Validate);
            if (!string.IsNullOrEmpty(body)) {
                var b = Encoding.UTF8.GetBytes(body);
                req.ContentType = "application/x-www-form-urlencoded"; req.ContentLength = b.Length;
                using (var rs = req.GetRequestStream()) { rs.Write(b, 0, b.Length); }
            }
            HttpWebResponse resp;
            try { resp = (HttpWebResponse)req.GetResponse(); }
            catch (WebException we) {
                if (we.Response == null) { r.Error = we.Message; return r; }
                resp = (HttpWebResponse)we.Response;
            }
            using (resp) {
                r.Status = (int)resp.StatusCode; r.ContentType = resp.ContentType;
                r.Text = ReadCapped(resp, r);
            }
        } catch (Exception ex) { r.Error = ex.Message; }
        return r;
    }
    static string ReadCapped(HttpWebResponse resp, GeoScadaHttp r) {
        if (resp.ContentLength > MaxBytes) { r.Error = "Response larger than the " + MaxBytes + " byte limit"; return ""; }
        using (var st = resp.GetResponseStream())
        using (var ms = new MemoryStream()) {
            var buf = new byte[65536]; int n; long total = 0;
            while ((n = st.Read(buf, 0, buf.Length)) > 0) {
                total += n;
                if (total > MaxBytes) { r.Error = "Response larger than the " + MaxBytes + " byte limit"; return ""; }
                ms.Write(buf, 0, n);
            }
            Encoding enc = Encoding.UTF8;
            try { if (!string.IsNullOrEmpty(resp.CharacterSet)) enc = Encoding.GetEncoding(resp.CharacterSet); } catch { }
            return enc.GetString(ms.ToArray()).TrimStart('\uFEFF');
        }
    }
    public static bool Validate(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors errors) {
        if (errors == SslPolicyErrors.None) return true;
        var req = sender as HttpWebRequest;
        return TrustedHost != null && req != null &&
               string.Equals(req.RequestUri.Host, TrustedHost, StringComparison.OrdinalIgnoreCase);
    }
}
"@
    if ($script:IsCore) { $src = "#pragma warning disable SYSLIB0014`n" + $src }   # HttpWebRequest is 'obsolete' on .NET 6+ but fully supported
    $addType = @{ TypeDefinition = $src }
    if ($script:IsCore) {
        # Resolve the assemblies from already-loaded types so the list is right for every .NET version.
        $addType.ReferencedAssemblies = @(
            [System.Net.HttpWebRequest], [System.Net.CookieContainer], [System.Net.WebHeaderCollection],
            [System.Net.Security.RemoteCertificateValidationCallback], [System.Net.Security.SslPolicyErrors],
            [System.Security.Cryptography.X509Certificates.X509Certificate], [System.Security.Cryptography.X509Certificates.X509Chain],
            [System.Text.Encoding], [System.IO.MemoryStream], [System.Uri]
        ) | ForEach-Object { $_.Assembly.Location } | Where-Object { $_ } | Select-Object -Unique
    }
    Add-Type @addType
}
[GeoScadaHttp]::Cookies     = New-Object System.Net.CookieContainer
[GeoScadaHttp]::MaxBytes    = [long]$MaxResponseMB * 1024 * 1024
[GeoScadaHttp]::TrustedHost = if ($SkipCertificateCheck) { $baseUri.Host } else { $null }
if ($SkipCertificateCheck) { Write-Warning "Certificate validation relaxed for host '$($baseUri.Host)' only (-SkipCertificateCheck)." }

# TLS 1.2 is added (never substituted) only on old .NET Frameworks that do not use the OS
# default protocol list; the previous value is restored when the script finishes.
$script:PrevSecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol
if ($script:PrevSecurityProtocol -ne [System.Net.SecurityProtocolType]::SystemDefault -and
    -not ($script:PrevSecurityProtocol -band [System.Net.SecurityProtocolType]::Tls12)) {
    [System.Net.ServicePointManager]::SecurityProtocol = $script:PrevSecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
}

# --- State ----------------------------------------------------------------
$script:LoggedOn    = $false
$script:ServedWithoutLogon = 0   # /db requests answered although this run had not logged on
$script:Mimics      = [ordered]@{}   # fullName -> definition (ordered hashtable, serialised to JSON)
$script:Fetched     = 0
$script:FetchErrors = 0
$script:ObjectsSeen = 0
$script:Walked      = New-Object System.Collections.Generic.HashSet[string]

if (-not $Credential -and $Username) {
    $sec = if ($Password) { ConvertTo-SecureString $Password -AsPlainText -Force } else { New-Object System.Security.SecureString }
    $Credential = New-Object System.Management.Automation.PSCredential($Username, $sec)
}

function HE([string]$s) { [System.Web.HttpUtility]::HtmlEncode($s) }

function Invoke-GeoRequest {
    <# Returns @{ Status=[int]; Xml=[xml]|$null; Text=[string]; Error=[string] } and never throws on HTTP errors. #>
    param([string]$Url, [string]$Method = 'GET', $Body)
    $r = [GeoScadaHttp]::Send($Url, $Method, $Body)
    if ($r.Error -and $r.Status -eq 0) { throw "Request to $Url failed: $($r.Error)" }
    $text = if ($r.Error) { '' } else { [string]$r.Text }
    $xml = $null
    if ($text -and $text.TrimStart().StartsWith('<')) { try { $xml = [xml]$text } catch { $xml = $null } }
    return @{ Status = [int]$r.Status; Xml = $xml; Text = $text; Error = $r.Error }
}

function Test-LogonPage($r) {
    return ($null -ne $r.Xml -and $null -ne $r.Xml.Page -and $null -ne $r.Xml.Page.SelectSingleNode('Logon'))
}

function Invoke-GeoLogon {
    if (-not $script:Credential) {
        $canPrompt = [Environment]::UserInteractive
        try { if ([Console]::IsInputRedirected) { $canPrompt = $false } } catch { }
        if (-not $canPrompt) { throw 'The web server requires a logon. Supply -Credential or -Username/-Password.' }
        $script:Credential = Get-Credential -Message "Geo SCADA logon for $BaseUrl"
        if (-not $script:Credential) { throw 'No credentials supplied.' }
    }
    $cred  = $script:Credential
    $plain = $cred.GetNetworkCredential().Password
    Write-Host "Logging on to $BaseUrl as $($cred.UserName) ..." -ForegroundColor Cyan
    $body = 'user=' + [Uri]::EscapeDataString($cred.UserName) + '&password=' + [Uri]::EscapeDataString($plain) + '&redir=%2F'
    $r = Invoke-GeoRequest -Url ($BaseUrl + 'logon') -Method 'POST' -Body $body
    if ($r.Xml -and $r.Xml.Page) {
        $logon = $r.Xml.Page.SelectSingleNode('Logon')
        if ($logon) { throw "Logon failed: $($logon.GetAttribute('status'))" }
        $msg = $r.Xml.Page.SelectSingleNode('Message')
        if ($msg -and $msg.InnerText -match 'successful') {
            $script:LoggedOn = $true
            Write-Host "Logged on as $($r.Xml.Page.GetAttribute('user'))." -ForegroundColor Green
            return
        }
    }
    # Fail closed: only a session cookie from the server counts as success from here on.
    $cookieNames = @([GeoScadaHttp]::Cookies.GetCookies([Uri]$BaseUrl) | ForEach-Object { $_.Name })
    if ($r.Status -eq 200 -and ($cookieNames | Where-Object { $_ -like 'CLEARSCADA*' })) {
        $script:LoggedOn = $true
        Write-Host 'Logged on (session cookie received).' -ForegroundColor Green
        return
    }
    throw "Logon failed: HTTP $($r.Status) without a logon success message or session cookie."
}

function Invoke-GeoLogoff {
    <# Best-effort POST /logoff (what the web UI's Log Off button does) so no session is left behind. #>
    if (-not $script:LoggedOn) { return }
    try {
        $r = [GeoScadaHttp]::Send($BaseUrl + 'logoff', 'POST', 'dummy=')
        if ($r.Error -or ($r.Status -ne 204 -and $r.Status -ne 200)) { Write-Verbose "Logoff returned HTTP $($r.Status) $($r.Error)" }
        else { Write-Host 'Logged off.' -ForegroundColor Cyan }
    } catch { Write-Verbose "Logoff failed: $($_.Exception.Message)" }
    $script:LoggedOn = $false
}

function Get-GeoXml {
    <# Fetch a /db/ URL as XML, logging on if the server redirects to the logon page. Returns @{ Xml; Status; Error } #>
    param([string]$RelativeUrl)
    $url = $BaseUrl + $RelativeUrl
    $r = Invoke-GeoRequest -Url $url
    if (Test-LogonPage $r) {
        if ($script:LoggedOn) { return @{ Xml = $null; Status = 401; Error = 'Not authorised (logon page returned)' } }
        Invoke-GeoLogon
        $r = Invoke-GeoRequest -Url $url
        if (Test-LogonPage $r) { return @{ Xml = $null; Status = 401; Error = 'Not authorised (logon page returned after logon)' } }
    } elseif (-not $script:LoggedOn -and $r.Status -eq 200) {
        $script:ServedWithoutLogon++
    }
    if ($r.Error) { return @{ Xml = $null; Status = $r.Status; Error = $r.Error } }
    if ($r.Status -ne 200) {
        $why = "HTTP $($r.Status)"
        if ($r.Xml -and $r.Xml.Page) {
            $err = $r.Xml.Page.SelectSingleNode('Error'); if ($err) { $why += ': ' + $err.InnerText.Trim() }
            elseif ($r.Xml.Page.GetAttribute('title')) { $why += ': ' + $r.Xml.Page.GetAttribute('title') }
        }
        return @{ Xml = $null; Status = $r.Status; Error = $why }
    }
    if (-not $r.Xml) { return @{ Xml = $null; Status = $r.Status; Error = 'Response was not XML' } }
    return @{ Xml = $r.Xml; Status = 200; Error = $null }
}

function Get-DbUrl([string]$FullName, [string]$Query) {
    return 'db/' + [Uri]::EscapeDataString($FullName) + $Query
}

# --- Name resolution --------------------------------------------------------
function Resolve-MimicRef {
    <# "SCX:////CMimic/..Pumps.Display.Symbol" relative to "A.B.C.Display.Overview" -> "A.B.C.Pumps.Display.Symbol" #>
    param([string]$Ref, [string]$ParentFullName)
    $name = $Ref
    if ($name -match '^[A-Za-z]+:/+[^/]*/(.*)$') { $name = $Matches[1] }   # strip SCX:////CMimic/
    elseif ($name.Contains('/')) { $name = $name.Substring($name.LastIndexOf('/') + 1) }
    $dots = 0
    while ($dots -lt $name.Length -and $name[$dots] -eq '.') { $dots++ }
    if ($dots -eq 0) { return $name }
    $rest  = $name.Substring($dots)
    $parts = $ParentFullName.Split('.')
    $keep  = $parts.Count - $dots
    if ($keep -le 0) { return $rest }
    return (($parts[0..($keep - 1)] -join '.') + '.' + $rest)
}

# --- Animation classification ------------------------------------------------
function Get-AnimationClass {
    <# Infers the guide's animation class from the expression text. Priority: Sql > Indirect > Historic > Method > Simple #>
    param([string]$Expr)
    if ([string]::IsNullOrWhiteSpace($Expr)) { return 'Simple' }
    $noStrings = [regex]::Replace($Expr, "'[^']*'", "''")             # drop string literals
    if ($noStrings -match '(?is)\bSELECT\b.*\bFROM\b') { return 'Sql' }
    $tags = [regex]::Matches($noStrings, '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value }
    $noTags = [regex]::Replace($noStrings, '"[^"]*"', '""')
    if ($noTags -match '\[') { return 'Indirect' }
    $method = $false
    foreach ($t in $tags) {
        if ($t.StartsWith('Parameter:')) { continue }
        if ($t.Contains(';')) { return 'Historic' }
        if ($t -match ':[^:]*:[^:]*:[^:]*$') { return 'Historic' }
        if ($t.Contains('(')) { $method = $true }
    }
    if ($method) { return 'Method' }
    return 'Simple'
}

function Get-CostEntry([string]$Type) {
    if ($CostTable.Contains($Type)) { return $CostTable[$Type] }
    return $CostTable['_Default']
}

function Get-AnimationScore($Anims) {
    $s = 0.0
    foreach ($a in $Anims) { if ($AnimationCosts.Contains($a.c)) { $s += [double]$AnimationCosts[$a.c] } }
    return $s
}

function Get-Band([double]$Score) {
    foreach ($b in $ComplexityBands) { if ($Score -lt [double]$b.Below) { return [string]$b.Name } }
    return [string]$ComplexityBands[-1].Name
}

# --- XML -> data ------------------------------------------------------------
function Convert-Animations($Node, [bool]$DirectOnly) {
    $xpath = if ($DirectOnly) { 'Animation' } else { './/Animation' }
    $out = @()
    foreach ($a in $Node.SelectNodes($xpath)) {
        $expr = $a.InnerText.Trim()
        $out += [ordered]@{ p = $a.GetAttribute('property'); c = (Get-AnimationClass $expr); x = $(if ($OmitExpressions) { '' } else { $expr }) }
    }
    return ,$out
}

function Convert-Actions($Node) {
    $out = @()
    foreach ($act in $Node.SelectNodes('PickMenu/Action')) {
        $txt = ($act.InnerText -replace '\s+', ' ').Trim()
        if ($txt.Length -gt 160) { $txt = $txt.Substring(0, 160) + '...' }
        if ($OmitExpressions) { $txt = '' }
        $out += [ordered]@{ n = $act.GetAttribute('name'); t = $act.GetAttribute('type'); x = $txt }
    }
    return ,$out
}

function Test-IsObjectNode($Node) {
    if ($Node.NodeType -ne 'Element') { return $false }
    if ($Node.HasAttribute('name')) { return $true }
    return $CostTable.Contains($Node.LocalName) -and $Node.LocalName -notin @('Layer', 'Mimic', '_Default')
}

function Convert-Object {
    <# Converts one drawing object element (recursing into Groups). $Ctx accumulates per-mimic counters. #>
    param($Node, [string]$MimicName, $Ctx, [int]$GroupDepth = 0)

    $script:ObjectsSeen++
    if ($script:ObjectsSeen -gt $MaxObjects) { throw "Object limit of $MaxObjects exceeded (raise -MaxObjects if this is expected)." }
    $type = $Node.LocalName
    $o = [ordered]@{ n = $Node.GetAttribute('name'); t = $type }
    if ($Node.HasAttribute('hidden'))   { $o.h = $true }
    if ($Node.HasAttribute('disabled')) { $o.d = $true }
    if ($Node.HasAttribute('noCache'))  { $o.nc = $true }
    $tip = $Node.GetAttribute('tooltipText'); if ($tip) { $o.tip = $tip }

    $isGroup = ($type -eq 'Group')
    # For a group only its own direct (and pick-menu) animations belong to it; nested objects own
    # theirs. For everything else include nested <Animation>s (pick-menu argument animations etc.).
    $anims = Convert-Animations -Node $Node -DirectOnly $isGroup
    if ($isGroup) {
        foreach ($a in $Node.SelectNodes('PickMenu//Animation')) {
            $expr = $a.InnerText.Trim()
            $anims += [ordered]@{ p = $a.GetAttribute('property'); c = (Get-AnimationClass $expr); x = $(if ($OmitExpressions) { '' } else { $expr }) }
        }
    }
    $acts = Convert-Actions $Node
    if ($anims.Count) { $o.anim = @($anims) }
    if ($acts.Count)  { $o.acts = @($acts) }

    $entry = Get-CostEntry $type
    $cost  = [double]$entry.Base
    $detail = [ordered]@{}

    switch ($type) {
        'EmbeddedMimic' {
            $raw = $Node.GetAttribute('ref')
            $o.raw = $raw
            $o.ref = Resolve-MimicRef -Ref $raw -ParentFullName $MimicName
            # Parameter bindings vs. animations: an EmbeddedMimic <Animation> whose property is
            # not a standard drawing property is a parameter being passed to the child mimic.
            $stdProps = @('Visible','Hidden','Disabled','PickDisabled','PosVal','X','Y','Width','Height','Rotation','Tooltip','TooltipText','Pos','Size','Enabled')
            $params = @(); $real = @()
            foreach ($a in $anims) {
                if ($stdProps -contains $a.p) { $real += $a } else { $params += $a }
            }
            $o.params = $params.Count
            if ($params.Count) { $o.panim = @($params) }
            if ($real.Count) { $o.anim = @($real) } else { $o.Remove('anim') }
            # Parameter bindings are scored at PerParameter each; set ParametersAreAnimations = $true
            # in the cost table to ALSO score each binding as an animation of its inferred class.
            if (-not $entry.ParametersAreAnimations) { $anims = $real }
            if ($entry.PerParameter) { $cost += [double]$entry.PerParameter * $params.Count }
            $Ctx.Embedded++
            [void]$Ctx.EmbeddedRefs.Add($o.ref)
        }
        'Graph' { $o.raw = $Node.GetAttribute('ref') }
        'List' {
            $sqlAnimated = $false
            foreach ($a in $anims) { if ($a.p -eq 'Sql') { $sqlAnimated = $true } }
            if ($sqlAnimated) { $o.sqlAnim = $true; if ($entry.IfSqlAnimated) { $cost += [double]$entry.IfSqlAnimated } }
        }
        default {
            if ($entry.PerKB) {
                $data = $Node.SelectSingleNode('Data')
                if ($data) {
                    $kb = [math]::Round(($data.InnerText.Trim().Length * 0.75) / 1024, 1)
                    $o.kb = $kb
                    $cost += [double]$entry.PerKB * $kb
                }
            }
            if ($entry.PerExtraSegment) {
                $pts = $Node.SelectNodes('.//Figure//Pos').Count
                if ($pts -eq 0) { $pts = $Node.SelectNodes('Pos').Count }
                $extra = [math]::Max(0, $pts - 2)
                $o.pts = $pts
                $cost += [double]$entry.PerExtraSegment * $extra
            }
        }
    }

    $animScore = 0.0
    if ($entry.AddAnimations -ne $false) { $animScore = Get-AnimationScore $anims }
    $cost += $animScore
    if ($ExtraCosts.PickMenuAction) { $cost += [double]$ExtraCosts.PickMenuAction * $acts.Count }
    $o.cost = [math]::Round($cost, 2)

    # per-mimic counters
    $Ctx.Objects++
    if (-not $Ctx.ByType.Contains($type)) { $Ctx.ByType[$type] = 0 }
    $Ctx.ByType[$type]++
    foreach ($a in $anims) { $Ctx.Anims++; $Ctx.AnimClass[$a.c]++ }
    $Ctx.ElementScore += ($cost - $animScore)
    $Ctx.AnimScore    += $animScore
    $Ctx.Actions      += $acts.Count

    if ($isGroup) {
        $kids = @()
        if ($GroupDepth -ge $MaxGroupDepth) {
            $o.truncated = $true
            Write-Warning "Group nesting deeper than -MaxGroupDepth ($MaxGroupDepth) in '$MimicName'; members of '$($o.n)' not descended."
        } else {
            foreach ($c in $Node.ChildNodes) {
                if (Test-IsObjectNode $c) { $kids += (Convert-Object -Node $c -MimicName $MimicName -Ctx $Ctx -GroupDepth ($GroupDepth + 1)) }
            }
        }
        $o.kids = @($kids)
    }
    return $o
}

function Convert-Mimic {
    <# Parses a <Mimic> document into the definition stored in $script:Mimics #>
    param([string]$FullName, [xml]$Xml)

    $m   = $Xml.DocumentElement
    $ctx = @{ Objects = 0; Embedded = 0; Anims = 0; Actions = 0; ElementScore = 0.0; AnimScore = 0.0
              ByType = [ordered]@{}; AnimClass = [ordered]@{ Simple = 0; Method = 0; Historic = 0; Indirect = 0; Sql = 0 }
              EmbeddedRefs = New-Object System.Collections.ArrayList }

    $def = [ordered]@{
        name = $FullName
        w    = $m.GetAttribute('w'); h = $m.GetAttribute('h')
        ver  = $m.GetAttribute('configVersion')
        web  = ($m.GetAttribute('webEnabled') -eq 'true')
        bg   = $m.GetAttribute('backColour')
    }

    # Parameters (skip group headers)
    $params = @()
    foreach ($p in $m.SelectNodes('Param')) { if ($p.GetAttribute('group') -ne 'true') { $params += $p.GetAttribute('name') } }
    $def.params = @($params)

    # Script
    $scriptNode = $m.SelectSingleNode('Script')
    if ($scriptNode) {
        $libs = @(); foreach ($l in $scriptNode.SelectNodes('Library')) { $libs += $(if ($OmitExpressions) { '(library)' } else { $l.GetAttribute('name') }) }
        $lines = ($scriptNode.InnerText -split "`n").Count
        $def.script = [ordered]@{ libs = @($libs); lines = $lines }
    }

    # Mimic-level animations and pick menu
    $mAnims = Convert-Animations -Node $m -DirectOnly $true
    foreach ($a in $m.SelectNodes('PickMenu//Animation')) {
        $expr = $a.InnerText.Trim(); $mAnims += [ordered]@{ p = $a.GetAttribute('property'); c = (Get-AnimationClass $expr); x = $(if ($OmitExpressions) { '' } else { $expr }) }
    }
    $mActs = Convert-Actions $m
    if ($mAnims.Count) { $def.anim = @($mAnims) }
    if ($mActs.Count)  { $def.acts = @($mActs) }
    $mimicAnimScore = 0.0
    if ((Get-CostEntry 'Mimic').AddAnimations -ne $false) { $mimicAnimScore = Get-AnimationScore $mAnims }
    foreach ($a in $mAnims) { $ctx.Anims++; $ctx.AnimClass[$a.c]++ }
    $ctx.AnimScore += $mimicAnimScore
    $ctx.Actions   += $mActs.Count

    # Layers
    $layerEntry = Get-CostEntry 'Layer'
    $layerScore = 0.0
    $layers = @()
    foreach ($ln in $m.SelectNodes('Layer')) {
        $L = [ordered]@{ n = $ln.GetAttribute('name') }
        if ($ln.HasAttribute('disabled')) { $L.d = $true }
        if ($ln.HasAttribute('fixed'))    { $L.f = $true }
        if ($ln.HasAttribute('hidden'))   { $L.h = $true }
        $lAnims = Convert-Animations -Node $ln -DirectOnly $true
        if ($lAnims.Count) { $L.anim = @($lAnims) }
        $lScore = [double]$layerEntry.Base
        if ($layerEntry.AddAnimations -ne $false) { $lScore += Get-AnimationScore $lAnims }
        foreach ($a in $lAnims) { $ctx.Anims++; $ctx.AnimClass[$a.c]++ }
        $objs = @()
        foreach ($c in $ln.ChildNodes) {
            if (Test-IsObjectNode $c) { $objs += (Convert-Object -Node $c -MimicName $FullName -Ctx $ctx) }
        }
        $L.objs = @($objs)
        $L.cost = [math]::Round($lScore, 2)
        $layerScore += $lScore
        $layers += $L
    }
    $def.layers = @($layers)

    # Extras (default 0)
    $extra = 0.0
    if ($def.script) {
        $extra += [double]$ExtraCosts.ScriptPresent
        $extra += [double]$ExtraCosts.ScriptLibrary * $def.script.libs.Count
        $extra += [double]$ExtraCosts.ScriptLine    * $def.script.lines
    }
    $extra += [double]$ExtraCosts.PickMenuAction * $mActs.Count

    $own = $ctx.ElementScore + $ctx.AnimScore + $layerScore + $extra
    $def.counts = [ordered]@{
        layers = $layers.Count; objs = $ctx.Objects; emb = $ctx.Embedded; anim = $ctx.Anims; acts = $ctx.Actions
        byType = $ctx.ByType; animClass = $ctx.AnimClass
    }
    $def.cost = [ordered]@{
        own = [math]::Round($own, 2); total = $null; band = $null
        bd  = [ordered]@{ elements = [math]::Round($ctx.ElementScore, 2); animations = [math]::Round($ctx.AnimScore, 2)
                          layers = [math]::Round($layerScore, 2); extras = [math]::Round($extra, 2); embedded = $null }
    }
    $def.emb     = @($ctx.EmbeddedRefs | Select-Object -Unique)
    $def.parents = @()
    return $def
}

# --- Crawl --------------------------------------------------------------------
function New-FailedMimicDef([string]$FullName, [string]$Why) {
    $script:FetchErrors++
    if ($FailOnFetchError) { throw "Failed to fetch mimic '$FullName': $Why (-FailOnFetchError)" }
    Write-Warning "Failed to fetch mimic '$FullName': $Why"
    return [ordered]@{ name = $FullName; err = $Why; layers = @(); params = @(); emb = @(); parents = @()
                       counts = [ordered]@{ layers = 0; objs = 0; emb = 0; anim = 0; acts = 0; byType = [ordered]@{}; animClass = [ordered]@{} }
                       cost = [ordered]@{ own = 0; total = 0; band = 'n/a'; bd = [ordered]@{} } }
}

function Get-MimicDef {
    param([string]$FullName)
    if ($script:Mimics.Contains($FullName)) { return $script:Mimics[$FullName] }

    if ($script:Fetched -ge $MaxMimics) {
        $def = New-FailedMimicDef $FullName "not fetched: -MaxMimics ($MaxMimics) reached"
        $script:Mimics[$FullName] = $def
        return $def
    }
    $script:Fetched++
    Write-Progress -Activity 'Walking Geo SCADA mimics' -Status "Fetched $($script:Fetched) mimics (current: $FullName)"
    $r = Get-GeoXml -RelativeUrl (Get-DbUrl $FullName '?content')
    if (-not $r.Xml -or $r.Xml.DocumentElement.LocalName -ne 'Mimic') {
        $why = if ($r.Error) { $r.Error } else { "Unexpected root element '$($r.Xml.DocumentElement.LocalName)'" }
        $def = New-FailedMimicDef $FullName $why
        $script:Mimics[$FullName] = $def
        return $def
    }
    try {
        $def = Convert-Mimic -FullName $FullName -Xml $r.Xml
    } catch {
        if ($FailOnFetchError) { throw }
        $def = New-FailedMimicDef $FullName "could not be parsed: $($_.Exception.Message)"
    }
    $script:Mimics[$FullName] = $def
    return $def
}

function Walk-Mimic {
    <# Depth-first walk. Each mimic's children are expanded once (global $script:Walked set);
       the per-path $Ancestors set only serves cycle detection and the depth limit. #>
    param([string]$FullName, [System.Collections.Generic.HashSet[string]]$Ancestors, [int]$Depth)
    $def = Get-MimicDef -FullName $FullName
    if ($def.err) { return }
    if ($Ancestors.Contains($FullName)) { return }
    if ($Depth -ge $MaxDepth) { $def.depthLimited = $true; return }
    if (-not $script:Walked.Add($FullName)) { return }
    [void]$Ancestors.Add($FullName)
    foreach ($child in $def.emb) {
        if ($Ancestors.Contains($child)) { continue }   # cycle - flagged per instance in Get-TotalCost
        Walk-Mimic -FullName $child -Ancestors $Ancestors -Depth ($Depth + 1)
        $cdef = $script:Mimics[$child]
        if ($cdef -and ($cdef.parents -notcontains $FullName)) { $cdef.parents = @($cdef.parents + $FullName) }
    }
    [void]$Ancestors.Remove($FullName)
}

$script:TotalMemo = @{}
function Get-TotalCost {
    <# own score + (child total) for every embedded instance; cyclic edges contribute 0 #>
    param([string]$FullName, [System.Collections.Generic.HashSet[string]]$Ancestors)
    if ($script:TotalMemo.ContainsKey($FullName)) { return $script:TotalMemo[$FullName] }
    $def = $script:Mimics[$FullName]
    if (-not $def -or $def.err) { return 0.0 }
    if ($Ancestors.Contains($FullName)) { return 0.0 }
    [void]$Ancestors.Add($FullName)
    $embedded = 0.0
    foreach ($L in $def.layers) {
        $stack = New-Object System.Collections.Stack
        foreach ($o in $L.objs) { $stack.Push($o) }
        while ($stack.Count) {
            $o = $stack.Pop()
            if ($o.t -eq 'EmbeddedMimic' -and $o.ref) {
                if ($Ancestors.Contains($o.ref)) { $o.cyc = $true }
                else {
                    $ct = Get-TotalCost -FullName $o.ref -Ancestors $Ancestors
                    $o.ctotal = [math]::Round($ct, 2)
                    $embedded += $ct
                    $cd = $script:Mimics[$o.ref]
                    if (-not $cd -or $cd.err) { $o.miss = $true }
                }
            }
            if ($o.kids) { foreach ($k in $o.kids) { $stack.Push($k) } }
        }
    }
    [void]$Ancestors.Remove($FullName)
    $total = [double]$def.cost.own + $embedded
    $def.cost.total = [math]::Round($total, 2)
    $def.cost.band  = Get-Band $total
    $def.cost.bd.embedded = [math]::Round($embedded, 2)
    $script:TotalMemo[$FullName] = $total
    return $total
}

function Get-GroupMimics {
    <# Lists every mimic under a group (recursively) using the /db/<Group> listing page. #>
    param([string]$GroupName)
    $found = New-Object System.Collections.ArrayList
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($GroupName)
    $seen = New-Object System.Collections.Generic.HashSet[string]
    while ($queue.Count) {
        $g = $queue.Dequeue()
        if (-not $seen.Add($g)) { continue }
        Write-Progress -Activity 'Listing mimics in group' -Status $g
        $rel = if ($g) { Get-DbUrl $g '' } else { 'db/' }
        $r = Get-GeoXml -RelativeUrl $rel
        if (-not $r.Xml) { Write-Warning "Could not list group '$g': $($r.Error)"; continue }
        foreach ($c in $r.Xml.SelectNodes('//ViewInfo/Children/Child')) {
            $childName = if ($g) { "$g.$($c.InnerText.Trim())" } else { $c.InnerText.Trim() }
            if ($c.GetAttribute('class') -eq 'Mimic') { [void]$found.Add($childName) }
            elseif ($c.GetAttribute('isGroup') -eq 'True') { $queue.Enqueue($childName) }
        }
    }
    Write-Progress -Activity 'Listing mimics in group' -Completed
    return @($found)
}

# --- Main -----------------------------------------------------------------
if (-not (Confirm-Disclaimer -Accepted:$AcceptDisclaimer)) { return }

$exitCode = 0
try {
if ($Guest) {
    Write-Host 'Guest mode: not logging on unless the server asks.' -ForegroundColor Cyan
} else {
    Invoke-GeoLogon   # always an explicit logon for this run (prompts if no credential was given)
}

$roots = @()
if ($Mimic) { $roots += $Mimic }
if ($Group) {
    Write-Host "Listing mimics under group '$Group' ..." -ForegroundColor Cyan
    $gm = Get-GroupMimics -GroupName $Group.Trim('.')
    Write-Host "  found $($gm.Count) mimics." -ForegroundColor Cyan
    $roots += $gm
}
$roots = @($roots | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
if (-not $roots.Count) { throw 'Nothing to walk - no mimics specified or found.' }

Write-Host "Walking $($roots.Count) root mimic(s) from $BaseUrl ..." -ForegroundColor Cyan
foreach ($rootName in $roots) {
    Walk-Mimic -FullName $rootName -Ancestors ([System.Collections.Generic.HashSet[string]]::new()) -Depth 0
}
Write-Progress -Activity 'Walking Geo SCADA mimics' -Completed

foreach ($name in @($script:Mimics.Keys)) {
    [void](Get-TotalCost -FullName $name -Ancestors ([System.Collections.Generic.HashSet[string]]::new()))
}

# Instance counts (how many times each mimic is embedded, across all fetched mimics)
$instances = @{}
foreach ($d in $script:Mimics.Values) {
    if ($d.err) { continue }
    foreach ($L in $d.layers) {
        $stack = New-Object System.Collections.Stack
        foreach ($o in $L.objs) { $stack.Push($o) }
        while ($stack.Count) {
            $o = $stack.Pop()
            if ($o.t -eq 'EmbeddedMimic' -and $o.ref) { if (-not $instances.ContainsKey($o.ref)) { $instances[$o.ref] = 0 }; $instances[$o.ref]++ }
            if ($o.kids) { foreach ($k in $o.kids) { $stack.Push($k) } }
        }
    }
}
foreach ($name in @($script:Mimics.Keys)) {
    $d = $script:Mimics[$name]
    $d.inst = if ($instances.ContainsKey($name)) { $instances[$name] } else { 0 }
    $d.root = ($roots -contains $name)
}

# --- Console summary / CSV -----------------------------------------------------
$rows = foreach ($d in $script:Mimics.Values) {
    [pscustomobject]@{
        Mimic      = $d.name
        Root       = $d.root
        Instances  = $d.inst
        Layers     = $d.counts.layers
        Objects    = $d.counts.objs
        Embedded   = $d.counts.emb
        Animations = $d.counts.anim
        Script     = [bool]$d.script
        OwnScore   = $d.cost.own
        TotalScore = $d.cost.total
        Band       = $d.cost.band
        Error      = $d.err
    }
}
$rows | Where-Object { $_.Root } | Sort-Object Mimic | Format-Table Mimic, Layers, Objects, Embedded, Animations, Script, OwnScore, TotalScore, Band -AutoSize | Out-String -Width 220 | Write-Host
if ($CsvPath) {
    $rows | Sort-Object Mimic | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV written to: $CsvPath" -ForegroundColor Green
}

# --- HTML output ----------------------------------------------------------------
$generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$errCount  = @($script:Mimics.Values | Where-Object { $_.err }).Count
# The report deliberately omits the server address and user name unless -IncludeServerLinks.
$meta = "$($roots.Count) root mimic(s) &bull; $($script:Mimics.Count) mimics fetched" +
        $(if ($errCount) { " ($errCount unavailable - report incomplete)" } else { '' }) +
        $(if ($IncludeServerLinks) { " &bull; Source: $(HE $BaseUrl)" } else { '' }) +
        $(if ($OmitExpressions) { ' &bull; expressions omitted' } else { '' }) +
        " &bull; Generated $generated"

$costsForPage = [ordered]@{
    elements   = $CostTable
    animations = $AnimationCosts
    extras     = $ExtraCosts
    bands      = @($ComplexityBands | ForEach-Object { [ordered]@{ Name = $_.Name; Below = $(if ([double]::IsInfinity($_.Below)) { $null } else { $_.Below }) } })
}

$ltEscape = [char]0x5C + 'u003c'
$json      = ($script:Mimics | ConvertTo-Json -Depth 60 -Compress).Replace('<', $ltEscape)
$rootsJson = (ConvertTo-Json @($roots) -Compress).Replace('<', $ltEscape)
if (-not $rootsJson.StartsWith('[')) { $rootsJson = "[$rootsJson]" }
$costsJson = ($costsForPage | ConvertTo-Json -Depth 10 -Compress).Replace('<', $ltEscape)
$baseUrlJson = if ($IncludeServerLinks) { (ConvertTo-Json $BaseUrl -Compress).Replace('<', $ltEscape) } else { 'null' }
$titleRoot = if ($roots.Count -eq 1) { $roots[0] } elseif ($Group) { $Group } else { "$($roots.Count) mimics" }

$template = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Geo SCADA Mimic Tree - {{ROOT}}</title>
<style>
  :root { color-scheme: light; --accent:#0f6cbd; --muted:#6b7280; --line:#e2e5e9; --bg:#f6f7f9; --panel:#ffffff;
          --mimic:#0f4c9c; --layer:#8b1a1a; --obj:#1f7a3a; }
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; margin:0; background:var(--bg); color:#1c1e21; }
  header { background:var(--accent); color:#fff; padding:14px 24px; box-shadow:0 1px 4px rgba(0,0,0,.2); }
  header h1 { margin:0; font-size:19px; font-weight:600; }
  header .meta { font-size:12.5px; opacity:.92; margin-top:4px; }
  .toolbar { padding:10px 24px; background:#fff; border-bottom:1px solid var(--line); position:sticky; top:0; z-index:5;
             display:flex; gap:10px; align-items:center; flex-wrap:wrap; }
  .toolbar button { font:inherit; padding:5px 12px; border:1px solid var(--accent); background:#fff; color:var(--accent);
                    border-radius:5px; cursor:pointer; }
  .toolbar button:hover { background:var(--accent); color:#fff; }
  .toolbar input[type=search] { font:inherit; padding:6px 10px; border:1px solid #ccc; border-radius:5px; min-width:260px; background:#fff; color:#1c1e21; }
  .toolbar label { font-size:12.5px; color:var(--muted); display:inline-flex; align-items:center; gap:4px; cursor:pointer; }
  .toolbar label input { margin:0; }
  main { padding:14px 24px 60px; }
  h2.sec { font-size:15px; margin:18px 0 8px; color:#374151; }
  ul { list-style:none; margin:0; padding-left:20px; }
  main > ul { padding-left:0; }
  li { margin:2px 0; }
  li.leaf { padding-left:18px; position:relative; }
  li.leaf::before { content:'\2022'; position:absolute; left:2px; color:#9aa0a6; }
  details > summary { cursor:pointer; list-style:none; padding:1px 4px 1px 0; border-radius:4px; }
  details > summary:hover { background:#eef4fb; }
  details > summary::-webkit-details-marker { display:none; }
  details > summary::before { content:'\25B6'; display:inline-block; width:14px; color:var(--accent); font-size:10px;
                              transition:transform .12s; }
  details[open] > summary::before { transform:rotate(90deg); }
  .name { font-weight:600; cursor:pointer; padding:1px 3px; border-radius:4px; }
  .name:hover { background:#dbeafe; text-decoration:underline; }
  .k-mimic > .name, .k-mimic > details > summary > .name { color:var(--mimic); }
  .k-layer > details > summary > .lname, .k-layer > .lname { color:var(--layer); font-weight:600; }
  .k-obj > .name, .k-obj > details > summary > .name { color:var(--obj); }
  li.sel > .name, li.sel > details > summary > .name { background:var(--accent); color:#fff; }
  .sub { color:var(--muted); font-size:12.5px; font-weight:normal; }
  .sub.italic { font-style:italic; }
  .count { color:var(--muted); font-size:12px; font-weight:normal; }
  .lnk { text-decoration:none; font-size:12px; opacity:.5; }
  .lnk:hover { opacity:1; }
  .tag { font-size:11px; padding:1px 6px; border-radius:8px; margin-left:6px; white-space:nowrap; }
  .tag.err { background:#fde7e7; color:#b91c1c; }
  .tag.cycle { background:#fff3cd; color:#92600a; }
  .tag.hid { background:#eef2f7; color:#475569; }
  .tag.dis { background:#f3e8ff; color:#6b21a8; }
  .tag.scr { background:#e0f2fe; color:#075985; }
  .tag.score { background:#eef2f7; color:#334155; font-variant-numeric:tabular-nums; }
  .band { font-size:11px; padding:1px 7px; border-radius:8px; margin-left:6px; font-weight:600; white-space:nowrap; }
  .band-simple  { background:#dcfce7; color:#166534; }
  .band-medium  { background:#fef3c7; color:#92400e; }
  .band-complex { background:#fee2e2; color:#991b1b; }
  .band-na      { background:#e5e7eb; color:#374151; }
  .hidden { display:none !important; }
  mark { background:#ffe58a; color:inherit; padding:0 1px; border-radius:2px; }
  .notice { font-size:12.5px; color:#92600a; background:#fff3cd; padding:6px 10px; border-radius:6px; margin:6px 0; }
  footer.disclaimer { margin:0 24px 40px; padding:12px 16px; border-top:1px solid var(--line);
                      font-size:11.5px; line-height:1.5; color:var(--muted); max-width:900px; }
  footer.disclaimer a { color:var(--accent); }

  /* ---- Summary table ---- */
  .tblwrap { overflow-x:auto; background:#fff; border:1px solid var(--line); border-radius:6px; max-height:60vh; overflow-y:auto; }
  table.sum { border-collapse:collapse; font-size:12.5px; min-width:100%; }
  table.sum th, table.sum td { padding:5px 9px; border-bottom:1px solid #f0f2f4; text-align:left; white-space:nowrap; }
  table.sum th { background:#f8fafc; cursor:pointer; user-select:none; position:sticky; top:0; }
  table.sum th.num, table.sum td.num { text-align:right; font-variant-numeric:tabular-nums; }
  table.sum tr:hover td { background:#f5f9ff; }
  table.sum td.mn { font-weight:600; color:var(--mimic); cursor:pointer; white-space:normal; }
  table.sum td.mn:hover { text-decoration:underline; }
  table.sum th.sorted::after { content:' \25BE'; color:var(--accent); }
  table.sum th.sorted.asc::after { content:' \25B4'; }
  .star { color:#d97706; }

  /* ---- Details panel ---- */
  #panel { position:fixed; top:0; right:0; height:100vh; width:min(520px,94vw); background:var(--panel);
           box-shadow:-3px 0 16px rgba(0,0,0,.25); transform:translateX(100%); transition:transform .18s ease;
           overflow-y:auto; z-index:20; }
  #panel.open { transform:none; }
  #panel .phead { position:sticky; top:0; background:var(--accent); color:#fff; padding:14px 16px; z-index:2; }
  #panel .phead h2 { margin:0; font-size:17px; font-weight:600; word-break:break-word; padding-right:28px; }
  #panel .phead .psub { font-size:13px; opacity:.95; font-style:italic; margin-top:2px; }
  #panel .phead .psub a { color:#fff; }
  #panel .phead .pbadges { margin-top:8px; display:flex; gap:6px; flex-wrap:wrap; align-items:center; font-size:12px; }
  #panel .phead .pill { background:rgba(255,255,255,.22); padding:2px 8px; border-radius:10px; }
  #panel .phead a { color:#fff; }
  #panel .pclose { position:absolute; top:10px; right:12px; background:none; border:none; color:#fff; font-size:22px;
                   cursor:pointer; line-height:1; }
  #panel .pbody { padding:4px 0 40px; }
  section.grp { border-bottom:1px solid var(--line); }
  section.grp > h3 { margin:0; padding:10px 16px; font-size:13px; text-transform:uppercase; letter-spacing:.04em;
                     color:#374151; cursor:pointer; display:flex; justify-content:space-between; align-items:center;
                     background:#f8fafc; user-select:none; }
  section.grp > h3:hover { background:#eef2f7; }
  section.grp > h3 .gc { color:var(--muted); font-weight:normal; text-transform:none; }
  section.grp.collapsed .items { display:none; }
  section.grp > h3::after { content:'\25BC'; font-size:9px; color:var(--muted); margin-left:8px; }
  section.grp.collapsed > h3::after { content:'\25B6'; }
  .items { padding:2px 0; }
  .row { padding:7px 16px; border-top:1px solid #f0f2f4; font-size:12.5px; }
  .row:first-child { border-top:none; }
  .row .rname { font-weight:600; font-family:'Cascadia Code',Consolas,monospace; font-size:12.5px; }
  .row .rtype { display:inline-block; background:#eef2f7; color:#334155; border-radius:4px; padding:0 6px;
                font-size:11.5px; margin-left:6px; }
  .row .rflags, .rflags { font-size:11px; color:var(--muted); margin-left:6px; }
  .row .rexpr { color:#4b5563; font-size:12px; margin-top:3px; white-space:pre-wrap; word-break:break-word;
                font-family:'Cascadia Code',Consolas,monospace; }
  .chip { display:inline-block; background:#eef4fb; color:#1e40af; border-radius:10px; padding:1px 8px; font-size:11.5px;
          margin:2px 4px 0 0; }
  .chip.c-simple { background:#f1f5f9; color:#334155; } .chip.c-method { background:#e0f2fe; color:#075985; }
  .chip.c-historic { background:#fef3c7; color:#92400e; } .chip.c-indirect { background:#f3e8ff; color:#6b21a8; }
  .chip.c-sql { background:#fee2e2; color:#991b1b; }
  a.mlink { color:var(--accent); cursor:pointer; text-decoration:none; }
  a.mlink:hover { text-decoration:underline; }
  table.kv { border-collapse:collapse; font-size:12.5px; width:100%; }
  table.kv td { padding:4px 16px; border-top:1px solid #f0f2f4; }
  table.kv td.num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
  table.kv tr.tot td { font-weight:600; border-top:2px solid var(--line); }
  .empty { padding:20px 16px; color:var(--muted); font-style:italic; }
  details.rules { margin:20px 0 0; font-size:12.5px; }
  details.rules summary { cursor:pointer; color:var(--accent); font-weight:600; }
  details.rules table { border-collapse:collapse; margin:8px 0 14px; background:#fff; }
  details.rules th, details.rules td { border:1px solid var(--line); padding:3px 8px; text-align:left; }
</style>
</head>
<body>
<header>
  <h1>Geo SCADA Mimic Tree</h1>
  <div class="meta">{{META}}</div>
</header>
<div class="toolbar">
  <button id="btnExpand" type="button">Expand all</button>
  <button id="btnCollapse" type="button">Collapse all</button>
  <input id="filter" type="search" placeholder="Filter: mimic, layer, object name/type, animation text...">
  <span id="filterInfo" class="count"></span>
  <label><input type="checkbox" id="showSummary" checked> Summary table</label>
</div>
<main>
  <div id="summary">
    <h2 class="sec">Mimics reached <span class="count" id="sumCount"></span></h2>
    <div class="tblwrap"><table class="sum" id="sumTable"><thead></thead><tbody></tbody></table></div>
  </div>
  <h2 class="sec">Hierarchy</h2>
  <div id="treeNotice" class="notice hidden"></div>
  <ul id="tree"></ul>
  <details class="rules"><summary>Scoring rules used for this report</summary><div id="rules"></div></details>
</main>

<footer class="disclaimer">
  <strong>GeoSCADA-Mimic-Walker</strong> &mdash; Copyright &copy; 2026 Adam Woodland. Licensed under the
  <a href="https://opensource.org/licenses/MIT" target="_blank">MIT License</a>.
  Complexity scoring based on the Schneider Electric community article
  <a href="https://community.se.com/t5/Geo-SCADA-Knowledge-Base/Guide-to-Mimic-Complexity/ba-p/278933" target="_blank">Guide to Mimic Complexity</a>;
  animation classes are inferred from expression text, so treat scores as indicative.
  <br>
  <strong>Disclaimer:</strong> This report is generated by a tool provided "AS IS", without warranty of any kind,
  express or implied. The author accepts no liability for any damages arising from its use. The generator reads
  mimic definitions over the Geo SCADA web server. It is <strong>NOT</strong> certified for production SCADA
  systems &mdash; review the code and test on a
  representative non-production system before use. You use it at your own risk and remain responsible for your own
  change-control and security policies.
</footer>

<aside id="panel" aria-hidden="true">
  <div class="phead">
    <button class="pclose" id="pClose" type="button" title="Close">&times;</button>
    <h2 id="pTitle"></h2>
    <div id="pSub" class="psub"></div>
    <div id="pBadges" class="pbadges"></div>
  </div>
  <div id="pBody" class="pbody"></div>
</aside>

<script>
var MIMICS = {{JSON}};
var ROOTS  = {{ROOTS}};
var COSTS  = {{COSTS}};
var BASEURL = {{BASEURL}};   // null unless the report was generated with -IncludeServerLinks
var NODE_CAP = 40000;

function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g,function(m){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m];}); }
function fmt(n){ if (n==null) return ''; return (Math.round(n*100)/100).toLocaleString(); }
function bandCls(b){ return 'band-'+String(b||'na').toLowerCase(); }
function bandPill(b){ return '<span class="band '+bandCls(b)+'">'+esc(b||'n/a')+'</span>'; }
function dbUrl(name,q){ return BASEURL ? BASEURL+'db/'+encodeURIComponent(name)+(q||'') : null; }
function rawLink(name){ var u = dbUrl(name,'?content'); return u ? ' <a class="lnk" href="'+esc(u)+'" target="_blank" rel="noopener" title="Open raw mimic XML">&#128279;</a>' : ''; }
function hl(txt, term){
  txt = String(txt==null?'':txt);
  if (!term) return esc(txt);
  var i = txt.toLowerCase().indexOf(term); if (i===-1) return esc(txt);
  return esc(txt.slice(0,i))+'<mark>'+esc(txt.slice(i,i+term.length))+'</mark>'+esc(txt.slice(i+term.length));
}

// ---- Search ----
function objBlob(o){
  if (o._s) return o._s;
  var p = [o.n, o.t, o.ref, o.raw, o.tip];
  (o.anim||[]).forEach(function(a){ p.push(a.p, a.x); });
  (o.panim||[]).forEach(function(a){ p.push(a.p, a.x); });
  (o.acts||[]).forEach(function(a){ p.push(a.n, a.t, a.x); });
  o._s = p.join(' ').toLowerCase(); return o._s;
}
function objMatches(o, term, visiting){
  if (objBlob(o).indexOf(term)!==-1) return true;
  if (o.t==='EmbeddedMimic' && o.ref && !o.cyc && mimicMatches(o.ref, term, visiting)) return true;
  if (o.kids) for (var i=0;i<o.kids.length;i++) if (objMatches(o.kids[i], term, visiting)) return true;
  return false;
}
function layerMatches(L, term, visiting){
  if ((L.n||'').toLowerCase().indexOf(term)!==-1) return true;
  for (var i=0;i<(L.anim||[]).length;i++) if ((L.anim[i].p+' '+L.anim[i].x).toLowerCase().indexOf(term)!==-1) return true;
  for (var j=0;j<L.objs.length;j++) if (objMatches(L.objs[j], term, visiting)) return true;
  return false;
}
var MATCH = {}, MATCH_TERM = null;
function mimicMatches(name, term, visiting){
  if (MATCH_TERM!==term){ MATCH = {}; MATCH_TERM = term; }
  if (name in MATCH) return MATCH[name];
  var m = MIMICS[name];
  if (!m || m.err) return MATCH[name] = (name.toLowerCase().indexOf(term)!==-1);
  visiting = visiting || {};
  if (visiting[name]) return false;
  visiting[name] = true;
  var r = [name].concat(m.params||[], (m.script&&m.script.libs)||[]).join(' ').toLowerCase().indexOf(term)!==-1;
  if (!r) for (var i=0;i<m.layers.length;i++) if (layerMatches(m.layers[i], term, visiting)) { r = true; break; }
  delete visiting[name];
  MATCH[name] = r; return r;
}

// ---- Tree ----
var OBJREF = {}, objSeq = 0;
function mkDetails(li, summaryHtml, fill, ctx){
  var d = document.createElement('details'), s = document.createElement('summary'), ul = document.createElement('ul');
  s.innerHTML = summaryHtml; d.appendChild(s); d.appendChild(ul); li.appendChild(d);
  if (ctx.eager){ d.open = true; fill(ul, ctx); }
  else d.addEventListener('toggle', function(){ if (d.open && !ul.hasChildNodes()) fill(ul, {eager:false, term:null, count:0}); });
  return d;
}
function capReached(ctx){ if (ctx.count++ < NODE_CAP) return false; ctx.capped = true; return true; }
function mimicBadges(m){
  if (!m) return ' <span class="tag err">not fetched</span>';
  if (m.err) return ' <span class="tag err" title="'+esc(m.err)+'">unavailable</span>';
  var h = ' ' + bandPill(m.cost.band) + ' <span class="tag score" title="Total score (own '+fmt(m.cost.own)+')">'+fmt(m.cost.total)+'</span>';
  if (m.script) h += ' <span class="tag scr" title="'+esc((m.script.libs||[]).join(', '))+'">script</span>';
  if (m.depthLimited) h += ' <span class="tag cycle">depth limit</span>';
  return h;
}
function mimicNameHtml(name, term){
  return '<span class="name" data-mimic="'+esc(name)+'">'+hl(name, term)+'</span>' + rawLink(name);
}
function fillMimic(ul, ctx, name, ancestors){
  var m = MIMICS[name]; if (!m || m.err) return;
  var anc = {}; for (var k in ancestors) anc[k] = true; anc[name] = true;
  m.layers.forEach(function(L){
    if (ctx.term && !layerMatches(L, ctx.term, {})) return;
    if (capReached(ctx)) return;
    var li = document.createElement('li'); li.className = 'k-layer';
    var flags = '';
    if (L.d) flags += ' <span class="tag dis">disabled</span>';
    if (L.f) flags += ' <span class="tag hid">fixed</span>';
    if (L.h) flags += ' <span class="tag hid">hidden</span>';
    var nAnim = (L.anim||[]).length;
    var sum = '<span class="lname">'+hl(L.n||'(unnamed layer)', ctx.term)+'</span> <span class="sub">layer</span>'
            + ' <span class="count">'+L.objs.length+' object'+(L.objs.length===1?'':'s')+(nAnim?', '+nAnim+' anim':'')+'</span>'
            + ' <span class="tag score" title="Layer score">'+fmt(L.cost)+'</span>'+flags;
    if (L.objs.length) mkDetails(li, sum, function(u, c){ fillObjects(u, c, L.objs, name, anc); }, ctx);
    else { li.className += ' leaf'; li.innerHTML = sum; }
    ul.appendChild(li);
  });
}
function fillObjects(ul, ctx, objs, mName, ancestors){
  objs.forEach(function(o){
    if (ctx.term && !objMatches(o, ctx.term, {})) return;
    if (capReached(ctx)) return;
    var id = 'o'+(++objSeq); OBJREF[id] = {m:mName, o:o};
    var li = document.createElement('li'); li.className = 'k-obj';
    var flags = '';
    if (o.h) flags += ' <span class="tag hid">hidden</span>';
    if (o.d) flags += ' <span class="tag dis">disabled</span>';
    var nAnim = (o.anim||[]).length;
    var nameHtml = '<span class="name" data-oid="'+id+'">'+hl(o.n||'(unnamed)', ctx.term)+'</span>';
    if (o.t==='EmbeddedMimic'){
      var child = MIMICS[o.ref];
      var sum = nameHtml + ' <span class="sub">embedded mimic &rarr;</span> ' + mimicNameHtml(o.ref, ctx.term)
              + (o.params?' <span class="count">'+o.params+' param'+(o.params===1?'':'s')+'</span>':'')
              + (nAnim?' <span class="count">'+nAnim+' anim</span>':'')
              + ' <span class="tag score" title="Instance score = own '+fmt(o.cost)+' + child total '+fmt(o.ctotal||0)+'">'+fmt(o.cost+(o.ctotal||0))+'</span>'
              + flags;
      if (!child || child.err){ li.className += ' leaf'; li.innerHTML = sum + mimicBadges(child); }
      else if (o.cyc || ancestors[o.ref]){ li.className += ' leaf'; li.innerHTML = sum + ' <span class="tag cycle">cycle - already on this branch</span>'; }
      else {
        sum += ' ' + bandPill(child.cost.band) + (child.script?' <span class="tag scr">script</span>':'');
        mkDetails(li, sum, function(u,c){ fillMimic(u, c, o.ref, ancestors); }, ctx);
      }
    } else if (o.t==='Group'){
      var kids = o.kids||[];
      var gs = nameHtml + ' <span class="sub">group</span> <span class="count">'+kids.length+' object'+(kids.length===1?'':'s')+(nAnim?', '+nAnim+' anim':'')+'</span>'
             + ' <span class="tag score">'+fmt(o.cost)+'</span>' + flags;
      if (kids.length) mkDetails(li, gs, function(u,c){ fillObjects(u, c, kids, mName, ancestors); }, ctx);
      else { li.className += ' leaf'; li.innerHTML = gs; }
    } else {
      var extra = '';
      if (o.t==='Graph' && o.raw) extra = ' <span class="sub italic">'+esc(o.raw.replace(/^SCX:\/+[^\/]*\//,''))+'</span>';
      if (o.kb!=null) extra += ' <span class="count">'+o.kb+' KB</span>';
      if (o.sqlAnim) extra += ' <span class="tag dis">SQL animated</span>';
      li.className += ' leaf';
      li.innerHTML = nameHtml + ' <span class="sub">'+esc(o.t)+'</span>' + extra
                   + (nAnim?' <span class="count">'+nAnim+' anim</span>':'')
                   + ((o.acts&&o.acts.length)?' <span class="count">'+o.acts.length+' action'+(o.acts.length===1?'':'s')+'</span>':'')
                   + ' <span class="tag score">'+fmt(o.cost)+'</span>' + flags;
    }
    ul.appendChild(li);
  });
}
function buildTree(opts){
  var tree = document.getElementById('tree'), notice = document.getElementById('treeNotice');
  tree.innerHTML = ''; OBJREF = {}; objSeq = 0;
  var ctx = {eager: !!opts.eager, term: opts.term||null, count:0, capped:false};
  var shown = 0;
  ROOTS.forEach(function(name){
    var m = MIMICS[name];
    if (ctx.term && !mimicMatches(name, ctx.term, {})) return;
    shown++;
    var li = document.createElement('li'); li.className = 'k-mimic';
    var sum = mimicNameHtml(name, ctx.term) + ' <span class="sub">mimic</span>' + mimicBadges(m)
            + (m && !m.err ? ' <span class="count">'+m.counts.layers+' layers, '+m.counts.objs+' objects, '+m.counts.emb+' embedded, '+m.counts.anim+' anim</span>' : '');
    if (!m || m.err){ li.className += ' leaf'; li.innerHTML = sum; }
    else {
      var d = mkDetails(li, sum, function(u,c){ fillMimic(u, c, name, {}); }, ctx);
      if (!ctx.eager && opts.openRoots) d.open = true;
    }
    tree.appendChild(li);
  });
  notice.classList.toggle('hidden', !ctx.capped);
  if (ctx.capped) notice.textContent = 'Tree truncated at '+NODE_CAP.toLocaleString()+' nodes - narrow the filter or expand branches individually.';
  return shown;
}
function currentTerm(){ var t = document.getElementById('filter').value.trim().toLowerCase(); return t||null; }
function expandAll(){ buildTree({eager:true, term:currentTerm()}); }
function collapseAll(){ document.getElementById('filter').value=''; document.getElementById('filterInfo').textContent=''; buildTree({eager:false, openRoots:false}); }
var filterTimer = null;
function doFilter(v){
  clearTimeout(filterTimer);
  filterTimer = setTimeout(function(){
    var term = v.trim().toLowerCase(), info = document.getElementById('filterInfo');
    if (!term){ buildTree({eager:false, openRoots:true}); info.textContent=''; return; }
    var shown = buildTree({eager:true, term:term});
    var total = 0; Object.keys(MIMICS).forEach(function(k){ if (mimicMatches(k, term, {})) total++; });
    info.textContent = shown+' root'+(shown===1?'':'s')+' shown, '+total+' matching mimic'+(total===1?'':'s');
  }, 200);
}

// ---- Summary table ----
var SUMCOLS = [
  {k:'name', t:'Mimic', cls:'mn', f:function(m){ return (m.root?'<span class="star" title="root mimic">&#9733;</span> ':'')+esc(m.name); }},
  {k:'inst', t:'Instances', num:true},
  {k:'layers', t:'Layers', num:true, g:function(m){ return m.counts.layers; }},
  {k:'objs', t:'Objects', num:true, g:function(m){ return m.counts.objs; }},
  {k:'emb', t:'Embedded', num:true, g:function(m){ return m.counts.emb; }},
  {k:'anim', t:'Animations', num:true, g:function(m){ return m.counts.anim; }},
  {k:'script', t:'Script', g:function(m){ return m.script?1:0; }, f:function(m){ return m.script?'yes':''; }},
  {k:'own', t:'Own score', num:true, g:function(m){ return m.cost.own; }, f:function(m){ return fmt(m.cost.own); }},
  {k:'total', t:'Total score', num:true, g:function(m){ return m.cost.total; }, f:function(m){ return fmt(m.cost.total); }},
  {k:'band', t:'Band', g:function(m){ return m.cost.total||0; }, f:function(m){ return m.err?'<span class="tag err" title="'+esc(m.err)+'">unavailable</span>':bandPill(m.cost.band); }}
];
var sumSort = {k:'total', asc:false};
function colVal(c, m){ return c.g ? c.g(m) : m[c.k]; }
function renderSummary(){
  var rows = Object.keys(MIMICS).map(function(k){ return MIMICS[k]; });
  var col = SUMCOLS.filter(function(c){ return c.k===sumSort.k; })[0];
  rows.sort(function(a,b){
    var x = colVal(col,a), y = colVal(col,b);
    if (typeof x==='string' || typeof y==='string'){ x=String(x||'').toLowerCase(); y=String(y||'').toLowerCase(); }
    else { x = x||0; y = y||0; }
    var r = x<y?-1:(x>y?1:0); if (r===0) r = a.name<b.name?-1:1;
    return sumSort.asc ? r : -r;
  });
  document.querySelector('#sumTable thead').innerHTML = '<tr>'+SUMCOLS.map(function(c){
    return '<th class="'+(c.num?'num ':'')+(c.k===sumSort.k?'sorted '+(sumSort.asc?'asc':''):'')+'" data-sort="'+esc(c.k)+'">'+c.t+'</th>'; }).join('')+'</tr>';
  document.querySelector('#sumTable tbody').innerHTML = rows.map(function(m){
    return '<tr>'+SUMCOLS.map(function(c){
      var v = c.f ? c.f(m) : esc(colVal(c,m));
      var on = c.cls==='mn' ? ' data-mimic="'+esc(m.name)+'"' : '';
      return '<td class="'+(c.num?'num':'')+(c.cls?' '+c.cls:'')+'"'+on+'>'+v+'</td>';
    }).join('')+'</tr>';
  }).join('');
  var e = rows.filter(function(m){ return m.err; }).length;
  document.getElementById('sumCount').textContent = '('+rows.length+(e?', '+e+' unavailable':'')+')';
}
function sortSummary(k){ if (sumSort.k===k) sumSort.asc=!sumSort.asc; else { sumSort.k=k; sumSort.asc=(k==='name'); } renderSummary(); }

// ---- Panel ----
function pickMimic(ev, name){ ev.preventDefault(); ev.stopPropagation(); selectLi(ev.target.closest('li')); openMimicPanel(name); }
function pickObject(ev, id){ ev.preventDefault(); ev.stopPropagation(); selectLi(ev.target.closest('li')); openObjectPanel(id); }
function selectLi(li){ document.querySelectorAll('#tree li.sel').forEach(function(x){ x.classList.remove('sel'); }); if (li) li.classList.add('sel'); }
function mlink(name){
  var m = MIMICS[name];
  return '<a class="mlink" data-mimic="'+esc(name)+'">'+esc(name)+'</a>'
       + (m&&!m.err ? ' '+bandPill(m.cost.band)+' <span class="tag score">'+fmt(m.cost.total)+'</span>' : (m&&m.err ? ' <span class="tag err">unavailable</span>' : ''));
}
function section(title, count, inner, collapsed){
  return '<section class="grp'+(collapsed?' collapsed':'')+'"><h3 class="gh">'+esc(title)+'<span class="gc">'+(count==null?'':count)+'</span></h3><div class="items">'+inner+'</div></section>';
}
function animRow(a){
  return '<div class="row"><span class="chip c-'+String(a.c).toLowerCase()+'" title="'+esc(a.c)+' animation: '+(COSTS.animations[a.c]||0)+' points">'+esc(a.c)+'</span> <span class="rname">'+esc(a.p)+'</span><div class="rexpr">'+esc(a.x)+'</div></div>';
}
function actRow(a){ return '<div class="row"><span class="rname">'+esc(a.n)+'</span><span class="rtype">'+esc(a.t)+'</span><div class="rexpr">'+esc(a.x)+'</div></div>'; }
function showPanel(title, sub, badges, body){
  document.getElementById('pTitle').textContent = title;
  var s = document.getElementById('pSub'); s.innerHTML = sub||''; s.style.display = sub?'':'none';
  document.getElementById('pBadges').innerHTML = badges;
  document.getElementById('pBody').innerHTML = body;
  var p = document.getElementById('panel'); p.classList.add('open'); p.setAttribute('aria-hidden','false'); p.scrollTop = 0;
}
function openMimicPanel(name){
  var m = MIMICS[name];
  if (!m){ showPanel(name, '', '<span class="pill">not fetched</span>', '<div class="empty">This mimic was not fetched during the crawl.</div>'); return; }
  var badges = [];
  if (m.err) badges.push('<span class="pill">unavailable: '+esc(m.err)+'</span>');
  else {
    badges.push(bandPill(m.cost.band), '<span class="pill">total '+fmt(m.cost.total)+'</span>', '<span class="pill">own '+fmt(m.cost.own)+'</span>');
    if (m.root) badges.push('<span class="pill">&#9733; root</span>');
    badges.push('<span class="pill">embedded '+m.inst+'&times;</span>');
    if (m.w) badges.push('<span class="pill">'+esc(m.w)+' &times; '+esc(m.h)+'</span>');
    if (m.ver) badges.push('<span class="pill">v'+esc(m.ver)+'</span>');
    if (m.web) badges.push('<span class="pill">web enabled</span>');
  }
  if (BASEURL) badges.push('<a href="'+esc(dbUrl(name,'?content'))+'" target="_blank" rel="noopener" title="Raw mimic XML">&#128279; xml</a>', '<a href="'+esc(dbUrl(name,''))+'" target="_blank" rel="noopener" title="Object page">&#128279; object</a>');
  var html = '';
  if (m.parents && m.parents.length) html += section('Embedded by', m.parents.length, m.parents.map(function(p){ return '<div class="row">'+mlink(p)+'</div>'; }).join(''));
  if (!m.err){
    var bd = m.cost.bd||{}, ac = m.counts.animClass||{};
    var acTxt = Object.keys(ac).filter(function(k){ return ac[k]; }).map(function(k){ return ac[k]+' '+k.toLowerCase(); }).join(', ');
    html += section('Score breakdown', fmt(m.cost.total),
      '<table class="kv">'
      + '<tr><td>Elements (objects incl. segment / KB / parameter uplifts)</td><td class="num">'+fmt(bd.elements)+'</td></tr>'
      + '<tr><td>Animations'+(acTxt?' ('+acTxt+')':'')+'</td><td class="num">'+fmt(bd.animations)+'</td></tr>'
      + '<tr><td>Layers ('+m.counts.layers+')</td><td class="num">'+fmt(bd.layers)+'</td></tr>'
      + (bd.extras?'<tr><td>Extras (scripts / actions)</td><td class="num">'+fmt(bd.extras)+'</td></tr>':'')
      + '<tr class="tot"><td>Own score</td><td class="num">'+fmt(m.cost.own)+'</td></tr>'
      + '<tr><td>Embedded mimics (sum of child totals, every instance)</td><td class="num">'+fmt(bd.embedded)+'</td></tr>'
      + '<tr class="tot"><td>Total score &rarr; '+esc(m.cost.band)+'</td><td class="num">'+fmt(m.cost.total)+'</td></tr>'
      + '</table>');
    html += section('Layers', m.layers.length, '<table class="kv">'+m.layers.map(function(L){
      return '<tr><td>'+esc(L.n)+(L.d?' <span class="tag dis">disabled</span>':'')+(L.f?' <span class="tag hid">fixed</span>':'')+'</td><td class="num">'+L.objs.length+' obj</td><td class="num">'+(L.anim||[]).length+' anim</td><td class="num">'+fmt(L.cost)+'</td></tr>'; }).join('')+'</table>');
    var bt = m.counts.byType||{};
    html += section('Object types', m.counts.objs, '<table class="kv">'+Object.keys(bt).sort(function(a,b){ return bt[b]-bt[a]; }).map(function(t){
      var e = COSTS.elements[t]||COSTS.elements._Default;
      return '<tr><td>'+esc(t)+(COSTS.elements[t]?'':' <span class="rflags">(not in cost table - default used)</span>')+'</td><td class="num">'+bt[t]+'</td><td class="num rflags">base '+e.Base+'</td></tr>'; }).join('')+'</table>');
    var embNames = m.emb||[];
    if (embNames.length){
      var counts = {};
      (function walk(objs){ objs.forEach(function(o){ if (o.t==='EmbeddedMimic') counts[o.ref]=(counts[o.ref]||0)+1; if (o.kids) walk(o.kids); }); })([].concat.apply([], m.layers.map(function(L){ return L.objs; })));
      html += section('Embedded mimics', embNames.length, embNames.slice().sort().map(function(n){ return '<div class="row">'+mlink(n)+' <span class="rflags">&times;'+(counts[n]||0)+'</span></div>'; }).join(''));
    }
    if (m.params && m.params.length) html += section('Parameters', m.params.length, '<div class="row">'+m.params.map(function(p){ return '<span class="chip">'+esc(p)+'</span>'; }).join('')+'</div>', true);
    if (m.script) html += section('Script', m.script.libs.length+' lib, '+m.script.lines+' lines', '<div class="row">'+(m.script.libs.length?m.script.libs.map(function(l){ return '<span class="chip">'+esc(l)+'</span>'; }).join(''):'<span class="rflags">inline script only</span>')+'</div>');
    if (m.anim && m.anim.length) html += section('Mimic-level animations', m.anim.length, m.anim.map(animRow).join(''), true);
    if (m.acts && m.acts.length) html += section('Mimic pick-menu actions', m.acts.length, m.acts.map(actRow).join(''), true);
  }
  showPanel(name, m.err?'':('Mimic'+(m.script?' &bull; scripted':'')), badges.join(' '), html||'<div class="empty">No details.</div>');
}
function openObjectPanel(id){
  var r = OBJREF[id]; if (!r) return;
  var o = r.o, e = COSTS.elements[o.t]||COSTS.elements._Default;
  var badges = ['<span class="pill">'+esc(o.t)+'</span>', '<span class="pill">score '+fmt(o.cost)+'</span>'];
  if (o.h) badges.push('<span class="pill">hidden</span>');
  if (o.d) badges.push('<span class="pill">disabled</span>');
  if (o.nc) badges.push('<span class="pill">no cache</span>');
  if (o.pts!=null) badges.push('<span class="pill">'+o.pts+' points</span>');
  if (o.kb!=null) badges.push('<span class="pill">'+o.kb+' KB</span>');
  if (o.tip) badges.push('<span class="pill">tooltip: '+esc(o.tip)+'</span>');
  var animPts = (o.anim||[]).reduce(function(s,a){ return s+(COSTS.animations[a.c]||0); },0);
  var html = section('Score', fmt(o.cost), '<table class="kv"><tr><td>Base ('+esc(o.t)+')</td><td class="num">'+e.Base+'</td></tr>'
        + (o.params?'<tr><td>Parameters &times; '+o.params+' @ '+(e.PerParameter||0)+'</td><td class="num">'+fmt(o.params*(e.PerParameter||0))+'</td></tr>':'')
        + (o.pts!=null&&e.PerExtraSegment?'<tr><td>Additional segments &times; '+Math.max(0,o.pts-2)+' @ '+e.PerExtraSegment+'</td><td class="num">'+fmt(Math.max(0,o.pts-2)*e.PerExtraSegment)+'</td></tr>':'')
        + (o.kb!=null&&e.PerKB?'<tr><td>Image '+o.kb+' KB @ '+e.PerKB+'</td><td class="num">'+fmt(o.kb*e.PerKB)+'</td></tr>':'')
        + (o.sqlAnim?'<tr><td>SQL query animated</td><td class="num">'+(e.IfSqlAnimated||0)+'</td></tr>':'')
        + ((o.anim||[]).length?'<tr><td>Animations'+(e.AddAnimations===false?' (not counted for this type)':'')+'</td><td class="num">'+(e.AddAnimations===false?0:fmt(animPts))+'</td></tr>':'')
        + (o.t==='EmbeddedMimic'?'<tr class="tot"><td>Own</td><td class="num">'+fmt(o.cost)+'</td></tr><tr><td>Child mimic total'+(o.cyc?' (cycle - excluded)':'')+'</td><td class="num">'+fmt(o.ctotal||0)+'</td></tr><tr class="tot"><td>Instance total</td><td class="num">'+fmt(o.cost+(o.ctotal||0))+'</td></tr>':'')
        + '</table>');
  if (o.t==='EmbeddedMimic') html += section('Embedded mimic', '', '<div class="row">'+mlink(o.ref)+'<div class="rexpr">'+esc(o.raw)+'</div></div>');
  if (o.t==='Graph' && o.raw) html += section('Trend reference', '', '<div class="row"><div class="rexpr">'+esc(o.raw)+'</div></div>');
  if (o.panim && o.panim.length) html += section('Parameters passed', o.panim.length, o.panim.map(function(a){ return '<div class="row"><span class="rname">'+esc(a.p)+'</span><div class="rexpr">'+esc(a.x)+'</div></div>'; }).join(''));
  if (o.anim && o.anim.length) html += section('Animations', o.anim.length, o.anim.map(animRow).join(''));
  if (o.acts && o.acts.length) html += section('Pick-menu actions', o.acts.length, o.acts.map(actRow).join(''));
  if (o.kids && o.kids.length) html += section('Group members', o.kids.length, o.kids.map(function(k){ return '<div class="row"><span class="rname">'+esc(k.n)+'</span><span class="rtype">'+esc(k.t)+'</span> <span class="rflags">'+fmt(k.cost)+'</span></div>'; }).join(''), true);
  showPanel(o.n||'(unnamed)', 'in '+mlink(r.m), badges.join(' '), html);
}
function closePanel(){ var p=document.getElementById('panel'); p.classList.remove('open'); p.setAttribute('aria-hidden','true'); selectLi(null); }
document.addEventListener('keydown', function(e){ if (e.key==='Escape') closePanel(); });

// ---- Event wiring (no inline handlers: names from the server are only ever placed in data attributes / text) ----
document.getElementById('btnExpand').addEventListener('click', expandAll);
document.getElementById('btnCollapse').addEventListener('click', collapseAll);
document.getElementById('pClose').addEventListener('click', closePanel);
document.getElementById('filter').addEventListener('input', function(){ doFilter(this.value); });
document.getElementById('showSummary').addEventListener('change', function(){ document.getElementById('summary').classList.toggle('hidden', !this.checked); });
document.addEventListener('click', function(ev){
  var t = ev.target;
  if (t.closest('a.lnk')) { ev.stopPropagation(); return; }                 // raw-XML link: follow it, don't toggle the tree
  var el = t.closest('#tree .name[data-mimic]');  if (el) { pickMimic(ev, el.getAttribute('data-mimic')); return; }
  el = t.closest('#tree .name[data-oid]');        if (el) { pickObject(ev, el.getAttribute('data-oid')); return; }
  el = t.closest('a.mlink[data-mimic], td.mn[data-mimic]'); if (el) { openMimicPanel(el.getAttribute('data-mimic')); return; }
  el = t.closest('#sumTable th[data-sort]');      if (el) { sortSummary(el.getAttribute('data-sort')); return; }
  el = t.closest('section.grp > h3.gh');          if (el) { el.parentNode.classList.toggle('collapsed'); return; }
});

// ---- Rules ----
function renderRules(){
  var h = '<table><tr><th>XML element</th><th>Base</th><th>Per extra segment</th><th>Per KB</th><th>Per parameter</th><th>If SQL animated</th><th>+ animations</th></tr>';
  Object.keys(COSTS.elements).forEach(function(k){ var e = COSTS.elements[k];
    h += '<tr><td>'+esc(k)+'</td><td>'+e.Base+'</td><td>'+(e.PerExtraSegment||'')+'</td><td>'+(e.PerKB||'')+'</td><td>'+(e.PerParameter||'')+'</td><td>'+(e.IfSqlAnimated||'')+'</td><td>'+(e.AddAnimations===false?'no':'yes')+'</td></tr>'; });
  h += '</table><table><tr><th>Animation class</th><th>Points</th></tr>';
  Object.keys(COSTS.animations).forEach(function(k){ h += '<tr><td>'+esc(k)+'</td><td>'+COSTS.animations[k]+'</td></tr>'; });
  h += '</table><table><tr><th>Extra (not in guide)</th><th>Points</th></tr>';
  Object.keys(COSTS.extras).forEach(function(k){ h += '<tr><td>'+esc(k)+'</td><td>'+COSTS.extras[k]+'</td></tr>'; });
  h += '</table><table><tr><th>Band</th><th>Total score</th></tr>';
  COSTS.bands.forEach(function(b){ h += '<tr><td>'+bandPill(b.Name)+'</td><td>'+(b.Below==null?'everything else':'&lt; '+b.Below)+'</td></tr>'; });
  document.getElementById('rules').innerHTML = h + '</table>';
}

renderSummary();
renderRules();
buildTree({eager:false, openRoots:true});
</script>
</body>
</html>
'@

$html = $template.
    Replace('{{ROOT}}',    (HE $titleRoot)).
    Replace('{{META}}',    $meta).
    Replace('{{JSON}}',    $json).
    Replace('{{ROOTS}}',   $rootsJson).
    Replace('{{COSTS}}',   $costsJson).
    Replace('{{BASEURL}}', $baseUrlJson)

$html | Out-File -FilePath $OutputPath -Encoding utf8
Write-Host "Done. $($script:Mimics.Count) mimics fetched from $($roots.Count) root(s)." -ForegroundColor Green
Write-Host "HTML written to: $OutputPath" -ForegroundColor Green
Write-Host ''
Write-Host 'NOTICE: The report describes SCADA system topology (mimic, layer and object names' -ForegroundColor Yellow
Write-Host '        and, unless -OmitExpressions was used, animation expressions and script names).' -ForegroundColor Yellow
Write-Host '        It may contain sensitive information and should be stored, shared and disposed' -ForegroundColor Yellow
Write-Host '        of in line with your relevant data protection and security policies.' -ForegroundColor Yellow
if ($script:ServedWithoutLogon) {
    Write-Warning ("The server answered $($script:ServedWithoutLogon) request(s) although this run had not logged on. " +
                   'That is either guest access or a still-live session for this client address left by an earlier logon ' +
                   '(the Geo SCADA web server keeps one for about a minute). Run without -Guest to log on explicitly.')
}
if ($script:FetchErrors) {
    Write-Warning "REPORT INCOMPLETE: $($script:FetchErrors) mimic(s) could not be fetched or parsed (see warnings above). Exit code 2."
    $exitCode = 2
}
}
finally {
    Invoke-GeoLogoff
    # Undo the only process-wide change this script may make.
    [System.Net.ServicePointManager]::SecurityProtocol = $script:PrevSecurityProtocol
    [GeoScadaHttp]::TrustedHost = $null
    [GeoScadaHttp]::Cookies = New-Object System.Net.CookieContainer
}
if ($exitCode) { exit $exitCode }
