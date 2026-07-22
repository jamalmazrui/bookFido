; =====================================================================
; bookFido installer script for Inno Setup 6.x
;
; Compile with the Inno Setup IDE (ISCC.exe) to produce
; bookFido_setup.exe.  The resulting installer:
;   - Targets 64-bit Windows 10 (and later) only.
;   - Requires administrator privileges.
;   - Prompts the user for the installation directory; default is
;     C:\Program Files\bookFido.
;   - Shows a brief MIT license summary on the welcome page (no extra
;     wizard screen).  The full license text is installed alongside
;     the program as License.htm.
;   - Registers the product for "Apps & Features" uninstall.
;   - Creates a desktop shortcut with hotkey Alt+Control+Shift+B (B
;     for bookFido).  Chosen after research: plain Ctrl+Alt+B is
;     claimed by Figma (detach instance), JetBrains IDEs (go to
;     implementation), and Oracle web applications (bookmark, bold),
;     and a global desktop hotkey would steal the keystroke from all
;     of them; Ctrl+Alt+Shift+B shows only one documented claimant
;     anywhere, a benchmark trigger in Chrome's Canary developer
;     build, and is unassigned by Windows itself.
;   - On the final wizard page, offers two PostInstall checkboxes
;     (both checked by default): launch bookFido (with a hotkey
;     reminder) and read the HTML documentation.
;
; This installer ships only the runtime distribution (the .exe, the
; documentation in HTML form, and the license).  The C# source, the
; build script, and this .iss script live in the GitHub repository.
; ReadMe.htm can be produced from README.md with 2htm.
;
; When installed under Program Files, the program detects that its
; folder refuses writes and keeps its working files (the log, the
; state snapshot, and diagnostic pages) in
; %LOCALAPPDATA%\bookFido instead; run as a portable exe from a
; writable folder, it keeps them beside itself.
; =====================================================================

#define sAppName       "bookFido"
#define sAppVersion    "1.1.0"
#define sAppPublisher  "Jamal Mazrui"
#define sAppUrl        "https://github.com/jamalmazrui/bookFido"
#define sAppExeName    "bookFido.exe"
#define sAppCopyright  "Copyright (c) 2026 Jamal Mazrui. MIT License."
#define sHotKey        "Alt+Ctrl+Shift+B"

[Setup]
AppId={{59FCCDDD-693F-4D26-9CC1-CB7092597915}

AppName={#sAppName}
AppVersion={#sAppVersion}
AppVerName={#sAppName} {#sAppVersion}
AppPublisher={#sAppPublisher}
AppPublisherURL={#sAppUrl}
AppSupportURL={#sAppUrl}
AppUpdatesURL={#sAppUrl}/releases
AppCopyright={#sAppCopyright}
VersionInfoVersion={#sAppVersion}

; Install under Program Files. {autopf} resolves to "Program Files"
; on 64-bit Windows when the installer runs in 64-bit mode (see
; ArchitecturesInstallIn64BitMode below).  The user can override this
; default on the wizard's directory page.
DefaultDirName={autopf}\{#sAppName}
DefaultGroupName={#sAppName}
DisableProgramGroupPage=yes
UsePreviousAppDir=yes

; Force the "Select Destination Location" page to always be shown,
; even on reinstall, so the install location is reviewable and
; obviously editable; UsePreviousAppDir pre-fills it on reinstall.
DisableDirPage=no
UsePreviousGroup=yes

OutputDir=.
OutputBaseFilename={#sAppName}_setup
Compression=lzma2
SolidCompression=yes
SetupIconFile={#sAppName}.ico
WizardStyle=modern

; Installer requires admin to write to Program Files.
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=

; 64-bit Windows only.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

Uninstallable=yes
UninstallDisplayIcon={app}\{#sAppExeName}
UninstallDisplayName={#sAppName} {#sAppVersion}

MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
; Replace the default welcome-page body text with one that includes a
; brief MIT license notice, keeping the license summary on an existing
; wizard screen rather than adding a dedicated page.  The full license
; text is installed alongside the program.
WelcomeLabel2=This will install [name/ver] on your computer.%n%n[name] catalogs your Audible library into accessible HTML, Markdown, and Excel, downloading companion PDFs and converting them to accessible HTML.%n%n[name] is licensed under the MIT License: free to use, copy, modify, and distribute; provided "as is" with no warranty. The full license text will be installed as License.htm in the program folder.%n%nIt is recommended that you close all other applications before continuing.

[Files]
; The runtime distribution: the executable, the HTML docs, and the
; license.  The icon is embedded in bookFido.exe at build time
; (csc /win32icon flag), so the .ico does not need to ship in the
; install directory.
Source: "{#sAppName}.exe";    DestDir: "{app}"; Flags: ignoreversion
Source: "ReadMe.htm";         DestDir: "{app}"; Flags: ignoreversion
Source: "License.htm";        DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu group.
Name: "{group}\{#sAppName}"; \
  Filename: "{app}\{#sAppExeName}"; \
  WorkingDir: "{app}"; \
  Comment: "Catalog an Audible library into accessible HTML, Markdown, and Excel"

Name: "{group}\{#sAppName} ReadMe"; \
  Filename: "{app}\ReadMe.htm"; \
  WorkingDir: "{app}"; \
  Comment: "Documentation for {#sAppName}"

Name: "{group}\Uninstall {#sAppName}"; \
  Filename: "{uninstallexe}"; \
  Comment: "Remove {#sAppName} from this computer"

; Desktop shortcut with the Alt+Ctrl+Shift+B hotkey.
Name: "{userdesktop}\{#sAppName}"; \
  Filename: "{app}\{#sAppExeName}"; \
  WorkingDir: "{app}"; \
  HotKey: {#sHotKey}; \
  Comment: "Catalog your Audible library ({#sHotKey})"

[Run]
; Post-install checkboxes shown on the final wizard page.  Both
; default to checked; the user can uncheck either to skip.  The launch
; checkbox label includes a reminder of the desktop hotkey.

FileName: "{app}\{#sAppExeName}"; \
  WorkingDir: "{app}"; \
  Description: "Launch {#sAppName} now (desktop hotkey: {#sHotKey})"; \
  Flags: nowait postinstall skipifsilent

FileName: "{app}\ReadMe.htm"; \
  Description: "Read documentation for {#sAppName}"; \
  Flags: postinstall shellexec skipifsilent
