# GeoSCADA-Mimic-Walker

Walk one or more **Geo SCADA** mimics through the web server, follow every embedded mimic, and
generate an interactive, self-contained HTML view of the **mimic &rarr; layer &rarr; object**
hierarchy together with a complexity score per mimic based on the
[Guide to Mimic Complexity](https://community.se.com/t5/Geo-SCADA-Knowledge-Base/Guide-to-Mimic-Complexity/ba-p/278933).

A companion to [GeoSCADA-Schema-Tree](https://github.com/adamwoodland2/GeoSCADA-Schema-Tree) and a
web-based replacement for the old ViewX-automation "Mimic Hierarchy" tool.

## How it works

Requesting a mimic from the Geo SCADA web server with the `?content` query string
(for example `https://<host>/db/Site.Display.Overview?content`) returns the mimic definition as XML
rather than the ActiveX viewer. The script:

1. **Logs on** using the web server's application-level logon form (`POST /logon`), prompting
   for credentials if none were given, keeps the session cookies for the run and logs off
   (`POST /logoff`) at the end. `-Guest` skips the up-front logon and only logs on if the
   server asks.
2. **Fetches the root mimic(s)** and every `<EmbeddedMimic>` reachable from them. Relative
   references such as `SCX:////CMimic/..Pumps.Display.Symbol` are resolved against the parent
   mimic's full name (each leading `.` strips one level), exactly as ViewX does.
3. **Records** layers, drawing objects (recursing into groups), animations, pick-menu actions,
   parameters and scripts for every mimic. Each mimic is fetched once however many times it is
   embedded; cycles and a maximum depth are guarded.
4. **Scores** every mimic: an *own* score for its own layers/objects/animations and a *total*
   score that adds the complexity of every embedded mimic instance, then grades it
   Simple / Medium / Complex.
5. **Writes one HTML file** with a sortable summary table, a lazily-rendered collapsible tree,
   search, and a details panel for every mimic and object.

Only built-in PowerShell / .NET features are used. No third-party modules or libraries. HTTP is
handled by a small C# helper compiled in-process (`HttpWebRequest`), which is what allows the
per-host certificate handling and response size cap described under *Security*.

## Features

- **Summary table** of every mimic reached: instance count, layers, objects, embedded mimics,
  animations, script, own score, total score and band. Click a column to sort, a name to open
  its details.
- **Hierarchy tree** &mdash; mimic (blue) &rarr; layer (dark red) &rarr; object (green), with
  embedded mimics expandable in place to show the child's layers, and groups expandable to their
  members. Hidden / disabled / fixed flags, per-node scores and band pills are shown inline.
  Branches render on demand, so very large trees stay responsive.
- **Search** across mimic names, layer names, object names and types, embedded references,
  animation expressions and pick-menu actions. Matching branches are expanded and highlighted.
- **Details panel** &mdash; for a mimic: where it is embedded, score breakdown, layers, object
  type counts, embedded mimics, parameters, script libraries and mimic-level animations, with
  links to the raw XML and the object page. For an object: score breakdown, animations (with
  the inferred class), parameters passed, pick-menu actions and group members.
- **Scoring rules** used for the report are embedded at the bottom of the page.
- **Optional CSV** of per-mimic statistics via `-CsvPath`.
- **Group mode** &mdash; `-Group` walks every mimic under a database group (recursively) using
  the web server's `/db/<Group>` listing pages.

## Complexity scoring

Scores follow the Geo SCADA Knowledge Base article *Guide to Mimic Complexity*:

| Mimic element | Score |
| ------------- | ----- |
| Additional line segment | 0.1 |
| Additional pipe segment | 0.2 |
| Primitive graphic element (text, line, shape) | 1 + animations |
| Pipe | 2 + animations |
| Bitmap image | 5 &times; file size in KB |
| Layer | 10 + animations |
| Pie chart | 20 + animations |
| SQL list | 30 + animations (+50 if the SQL query is animated) |
| Embedded mimic | 30 + complexity of the embedded mimic + 5 per parameter + animations |
| Alarm list | 50 + animations |
| Embedded trend / XY / XYZ / dynagraph | 100 |
| Remote image | 100 |

| Animation type | Score |
| -------------- | ----- |
| Simple tag animation | 5 |
| Object method animation | 15 |
| Historic animation | 20 |
| Indirect animation | 50 |
| SQL animation | 200 |

| Total | Band |
| ----- | ---- |
| &lt; 1000 | Simple |
| &lt; 5000 | Medium |
| &ge; 5000 | Complex |

**All of these values live in a clearly marked block at the top of the script** (`$CostTable`,
`$AnimationCosts`, `$ExtraCosts`, `$ComplexityBands`) and can be changed freely. The table is
keyed by the XML element name of the object (`Polyline`, `Text`, `Button`, `Bitmap`, `Group`,
`EmbeddedMimic`, `Graph`, `List`, `AlarmList` ...); unknown element names fall back to the
`_Default` row and are flagged in the report. `$ExtraCosts` lets you add points for scripts
and pick-menu actions, which the guide does not score (all default to 0).

### How the XML maps onto the guide

- **Animation class** is inferred from each expression's text, in priority order:
  SQL (`SELECT ... FROM` outside string literals) &rarr; Indirect (`[ ... ]` indirect tag) &rarr;
  Historic (tag containing `;` or the `:offset:period:max` calculation syntax) &rarr; Object
  method (tag containing a method call such as `"Point.CheckAccess('CTL')"`) &rarr; Simple.
  Treat the result as indicative.
- **Embedded mimic parameters** are the `<Animation>` elements on an `<EmbeddedMimic>` whose
  property is not a standard drawing property (Visible, PosVal, X, Y ...). They score 5 each and
  are *not* also scored as animations unless `ParametersAreAnimations` is set to `$true` in the
  `EmbeddedMimic` row.
- **Line segments** are counted from the `<Pos>` points of a polyline's figure; segments beyond
  the first add 0.1 each. **Bitmap size** is derived from the embedded base64 image data.
- **Total score** adds the child mimic's total for *every* instance embedded, as the guide
  describes, so a heavily-nested display can reach a very large total. An embedded mimic that
  is already on the current branch (a cycle) contributes 0 and is flagged.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+.
- Network access to the Geo SCADA web server, and a user with permission to view the mimics.

## Security

- **Transport.** `https` is required; a plain `http` `-BaseUrl` is refused unless you pass
  `-AllowHttp`, because the logon posts your credentials.
- **Certificates.** Validation is **on by default**. `-SkipCertificateCheck` accepts an untrusted
  (for example self-signed) certificate from the `-BaseUrl` host only, and only for the requests
  this script makes: the check is done per request through
  `HttpWebRequest.ServerCertificateValidationCallback`, so nothing process-wide is changed and
  other HTTPS traffic in the same PowerShell session is unaffected. It still removes
  man-in-the-middle protection for that connection, so prefer trusting the server certificate.
  TLS 1.2 is only *added* on old .NET Frameworks that lack it, and the previous protocol setting
  is restored when the script ends.
- **Logon.** The script fails closed: only the server's "Logon successful" page or a
  `CLEARSCADA*` session cookie counts as a logon. Every run logs on explicitly and logs off at
  the end; nothing is cached between runs.
- **Address-based sessions on the server.** After any logon the Geo SCADA web server keeps a
  short-lived session keyed on the *client address*: for roughly a minute a request with no
  cookies at all is still accepted (this is visible with `curl` and survives `/logoff`). That is
  why, with `-Guest`, a run started soon after another one may proceed without prompting. The
  script warns when the server answered requests although the run had not logged on. Without
  `-Guest` this cannot happen because the run always logs on as the user you give.
- **Report contents.** The HTML deliberately omits the server address and user name; pass
  `-IncludeServerLinks` to embed them so names link back to the web server. Object names,
  animation expressions, pick-menu action text and script library names *are* included because
  they are the point of the report; `-OmitExpressions` strips the expression and script text
  (keeping counts, classes and scores) if the report will leave your team. Treat the report as
  containing SCADA topology either way: the script prints a notice at the end of every run
  reminding you that the output may contain sensitive information and should be handled in line
  with your data protection and security policies.
- **Output safety.** Everything from the server is rendered as text or `data-` attributes; the
  page has no inline event handlers, so a hostile object or mimic name cannot execute script.
- **Resource limits.** Responses are read with a hard size cap (`-MaxResponseMB`, default 50),
  and the walk is bounded by `-MaxMimics` (5000), `-MaxObjects` (1,000,000), `-MaxDepth` (25
  embedded levels) and `-MaxGroupDepth` (50). Each mimic's children are expanded once even when
  it is embedded on many paths.
- **Partial reports.** If any mimic cannot be fetched or parsed the script still writes the
  report, prints `REPORT INCOMPLETE` and exits with code **2**; `-FailOnFetchError` stops at the
  first failure instead.

## Usage

```powershell
# Walk one mimic, prompting for credentials if the server requires a logon
.\GeoSCADA-Mimic-Walker.ps1 -BaseUrl https://scada/ -Mimic 'Sewage.Site.Display.Overview'

# Test server with a self-signed certificate: accept it for that host only
.\GeoSCADA-Mimic-Walker.ps1 -BaseUrl https://192.168.1.10/ -Mimic 'Sewage.Site.Display.Overview' -SkipCertificateCheck

# Non-interactive, with credentials and a CSV as well as the HTML
.\GeoSCADA-Mimic-Walker.ps1 -BaseUrl https://scada/ -Mimic 'Sewage.Site.Display.Overview' `
    -Username operator -Password 'secret' -OutputPath C:\temp\overview.html -CsvPath C:\temp\overview.csv -AcceptDisclaimer

# Several roots, or every mimic under a group
.\GeoSCADA-Mimic-Walker.ps1 -BaseUrl https://scada/ -Mimic 'A.Display.Overview','B.Display.Overview' -Credential (Get-Credential)
.\GeoSCADA-Mimic-Walker.ps1 -BaseUrl https://scada/ -Group 'Sewage.Site.Display' -Credential (Get-Credential)
```

Open the resulting `MimicTree.html` in any modern browser.

### Parameters

| Parameter               | Default                   | Description |
| ----------------------- | ------------------------- | ----------- |
| `-Mimic`                |                           | One or more mimic full names to start from. |
| `-Group`                |                           | Walk every mimic under this group (recursive). Can be combined with `-Mimic`. |
| `-BaseUrl`              | `https://localhost/`      | Root URL of the Geo SCADA web server. |
| `-Credential`           |                           | PSCredential for the application logon. |
| `-Username`/`-Password` |                           | Plain-text alternative to `-Credential`. If neither is given you are prompted (unless `-Guest`). |
| `-Guest`                | *(off)*                   | Try anonymously; only log on if the server demands it. |
| `-OutputPath`           | `MimicTree.html` (by script) | Where to write the HTML. |
| `-CsvPath`              |                           | Optional CSV of per-mimic statistics and scores. |
| `-MaxDepth`             | `25`                      | Maximum embedded-mimic nesting depth to follow. |
| `-SkipCertificateCheck` | *(off)*                   | Accept an untrusted certificate from the `-BaseUrl` host only. |
| `-AllowHttp`            | *(off)*                   | Permit a plain `http://` base URL. |
| `-IncludeServerLinks`   | *(off)*                   | Embed the server address so the report links back to the web server. |
| `-OmitExpressions`      | *(off)*                   | Leave animation expressions, action text and script library names out of the report. |
| `-FailOnFetchError`     | *(off)*                   | Stop at the first mimic that cannot be fetched (otherwise finish and exit 2). |
| `-MaxMimics`            | `5000`                    | Maximum distinct mimics to fetch. |
| `-MaxResponseMB`        | `50`                      | Maximum size of one web response. |
| `-MaxObjects`           | `1000000`                 | Maximum drawing objects parsed in total. |
| `-MaxGroupDepth`        | `50`                      | Maximum Group nesting depth within a mimic. |
| `-AcceptDisclaimer`     | *(off)*                   | Accept the disclaimer non-interactively. |

## Disclaimer

This script is provided **"AS IS"**, without warranty of any kind, express or implied. The author
accepts no liability for any damages arising from its use. It connects to the Geo SCADA web server
to read mimic definitions and writes an HTML file; with `-SkipCertificateCheck` it accepts an untrusted
certificate from that one host. It is **NOT** certified for production SCADA environments. Do **NOT** run it
against a production or safety-critical system without first reviewing the code and testing it on a
representative non-production system. You run it at your own risk and are responsible for compliance
with your own change-control and security policies.

You are prompted to accept this disclaimer before the script runs (or pass `-AcceptDisclaimer`).

## License

Copyright (c) 2026 Adam Woodland.

Licensed under the **MIT License** &mdash; see the [LICENSE](LICENSE) file, or
<https://opensource.org/licenses/MIT>. This is free software, and you are welcome to
redistribute it under those conditions; it comes with **ABSOLUTELY NO WARRANTY**.
