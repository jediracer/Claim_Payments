; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{E6891CD3-E9FA-4A74-96EA-6E505F2A390F}
AppName=Claim Payments
AppVersion=1.0
;AppVerName=Claim Payments 1.0
AppPublisher=Jared Behler
AppPublisherURL=https://github.com/jediracer/Claim_Payments
AppSupportURL=https://github.com/jediracer/Claim_Payments
AppUpdatesURL=https://github.com/jediracer/Claim_Payments
DefaultDirName={autopf}\Claim Payments
DisableDirPage=yes
DisableProgramGroupPage=yes
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
OutputDir=C:\Users\jbehler\python\Claim_Payments\installer
OutputBaseFilename=Claim Payment Setup
SetupIconFile=C:\Users\jbehler\python\Claim_Payments\images\claim_payments.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ChangesEnvironment=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\Claim_Payments.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_asyncio.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_bz2.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_cffi_backend.cp38-win32.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_ctypes.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_decimal.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_elementtree.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_hashlib.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_lzma.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_multiprocessing.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_overlapped.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_queue.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_socket.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_sqlite3.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_ssl.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_tkinter.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\_win32sysloader.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-console-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-datetime-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-debug-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-errorhandling-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-file-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-file-l1-2-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-file-l2-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-handle-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-heap-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-interlocked-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-libraryloader-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-localization-l1-2-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-memory-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-namedpipe-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-processenvironment-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-processthreads-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-processthreads-l1-1-1.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-profile-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-rtlsupport-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-string-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-synch-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-synch-l1-2-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-sysinfo-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-timezone-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-core-util-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-conio-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-convert-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-environment-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-filesystem-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-heap-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-locale-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-math-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-multibyte-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-process-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-runtime-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-stdio-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-string-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-time-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\api-ms-win-crt-utility-l1-1-0.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\base_library.zip"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\libcrypto-1_1.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\libffi-8.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\libopenblas.VTYUM5MXKVFE4PZZER3L7PNO6YB4XFF3.gfortran-win32.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\libsodium-1a96dce1.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\libssl-1_1.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\libzmq-v141-mt-4_3_4-809e0775.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\mfc140u.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\msvcp140.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\pyexpat.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\pyodbc.cp38-win32.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\python3.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\python38.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\pythoncom38.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\pywintypes38.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\qpdf28.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\select.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\sqlite3.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\tcl86t.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\tk86t.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\ucrtbase.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\unicodedata.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\VCRUNTIME140.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\win32api.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\win32clipboard.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\win32evtlog.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\win32pdh.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\win32security.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\win32trace.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\win32ui.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\win32wnet.pyd"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\altgraph-0.17.2.dist-info\*"; DestDir: "{app}\altgraph-0.17.2.dist-info"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\bcrypt\*"; DestDir: "{app}\bcrypt"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\cryptography\*"; DestDir: "{app}\cryptography"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\cryptography-36.0.1.dist-info\*"; DestDir: "{app}\cryptography-36.0.1.dist-info"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\IPython\*"; DestDir: "{app}\IPython"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\jedi\*"; DestDir: "{app}\jedi"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\lxml\*"; DestDir: "{app}\lxml"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\markupsafe\*"; DestDir: "{app}\markupsafe"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\nacl\*"; DestDir: "{app}\nacl"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\numpy\*"; DestDir: "{app}\numpy"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\pandas\*"; DestDir: "{app}\pandas"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\parso\*"; DestDir: "{app}\parso"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\pikepdf\*"; DestDir: "{app}\pikepdf"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\PIL\*"; DestDir: "{app}\PIL"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\pyinstaller-4.9.dist-info\*"; DestDir: "{app}\pyinstaller-4.9.dist-info"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\pytz\*"; DestDir: "{app}\pytz"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\pyzmq.libs\*"; DestDir: "{app}\pyzmq.libs"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\setuptools-58.0.4-py3.8.egg-info\*"; DestDir: "{app}\setuptools-58.0.4-py3.8.egg-info"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\tcl\*"; DestDir: "{app}\tcl"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\tcl8\*"; DestDir: "{app}\tcl8"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\tk\*"; DestDir: "{app}\tk"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\tornado\*"; DestDir: "{app}\tornado"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\wcwidth\*"; DestDir: "{app}\wcwidth"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\wheel-0.37.0-py3.9.egg-info\*"; DestDir: "{app}\wheel-0.37.0-py3.9.egg-info"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\win32com\*"; DestDir: "{app}\win32com"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\dist\Claim_Payments\zmq\*"; DestDir: "{app}\zmq"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\html\*"; DestDir: "{app}\html"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\images\*"; DestDir: "{app}\images"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\packages\poppler-21.03.0\*"; DestDir: "{app}\packages"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\jbehler\python\Claim_Payments\letters\*"; DestDir: "{app}\letters"; Flags: ignoreversion recursesubdirs createallsubdirs; Permissions: users-modify
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Registry]
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; \
    ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}\packages\poppler-21.03.0\Library\bin"; \
    Check: NeedsAddPath('{app}\packages\poppler-21.03.0\Library\bin')

[Code]
function NeedsAddPath(Param: string): boolean;
var
  OrigPath: string;
begin
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', OrigPath)
  then begin
    Result := True;
    exit;
  end;
  { look for the path with leading and trailing semicolon }
  { Pos() returns 0 if not found }
  Result := Pos(';' + Param + ';', ';' + OrigPath + ';') = 0;
end;

[Icons]
Name: "{autoprograms}\Claim Payments"; Filename: "{app}\Claim_Payments.exe"
Name: "{autodesktop}\Claim Payments"; Filename: "{app}\Claim_Payments.exe"; Tasks: desktopicon
Name: "{autoprograms}\Claim Payments"; Filename: "{app}\Claim_Payments.exe"; IconFilename: "{app}\images\claim_payments.ico"
Name: "{autodesktop}\Claim Payments"; Filename: "{app}\Claim_Payments.exe"; IconFilename: "{app}\images\claim_payments.ico"

[Run]
Filename: "{app}\Claim_Payments.exe"; Description: "{cm:LaunchProgram,Claim Payments}"; Flags: nowait postinstall skipifsilent
