Attribute VB_Name = "Module1"

'---------------------------------------
'coded by sihab sahariar sizan (15-02-2017)
'A module of Zharudar(borrnolab)
'Basically it will repair the common broken registry files.
'Don't forget to mention my name if you are using my code.
'www.facebook.com/sizan.first
'sizan009@gmail.com (My Email)
'-------------------------------------
Option Explicit

Private lReg As Long
Private KeyHandle As Long
Private lResult As Long
Private lValueType As Long
Private lDataBufSize As Long

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_DWORD = 4

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Const KEY_ROOT_EXE As String = "\shell\open\command"
Private Const MSWCV As String = "\Microsoft\Windows\" _
    & "CurrentVersion\"

Enum KeyREG
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Function CreateKeyReg(hKey As KeyREG, sPath As _
    String) As Long
    
    lReg = RegCreateKey(hKey, sPath, KeyHandle)
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function GetStringValue(hKey As KeyREG, _
    sPath As String, sValue As String) As Long
    
    Dim sBuff As String
    Dim intZeroPos As Integer
    
    lReg = RegOpenKey(hKey, sPath, KeyHandle)
    lResult = RegQueryValueEx(KeyHandle, sValue, 0&, _
        lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        sBuff = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(KeyHandle, sValue, 0&, _
            0&, ByVal sBuff, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(sBuff, Chr$(0))
            If intZeroPos > 0 Then
                GetStringValue = Left$(sBuff, _
                    intZeroPos - 1)
            Else
                GetStringValue = sBuff
            End If
        End If
    End If
    
End Function

Public Function SetStringValue(hKey As KeyREG, _
    sPath As String, sValue As String, sData As _
    String) As Long
    
    lReg = RegCreateKey(hKey, sPath, KeyHandle)
    lReg = RegSetValueEx(KeyHandle, sValue, 0, REG_SZ, _
        ByVal sData, Len(sData))
    lReg = RegCloseKey(KeyHandle)
    
End Function

Function GetDwordValue(ByVal hKey As KeyREG, ByVal sPath _
    As String, ByVal sValueName As String) As Long
    
    Dim lBuff As Long
    
    lReg = RegOpenKey(hKey, sPath, KeyHandle)
    lDataBufSize = 4
    lResult = RegQueryValueEx(KeyHandle, sValueName, _
        0&, lValueType, lBuff, lDataBufSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDwordValue = lBuff
        End If
    End If
    
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function SetDwordValue(ByVal hKey As KeyREG, _
    ByVal sPath As String, ByVal sValueName As String, _
    ByVal lData As Long) As Long
    
    lReg = RegCreateKey(hKey, sPath, KeyHandle)
    lResult = RegSetValueEx(KeyHandle, sValueName, 0&, _
        REG_DWORD, lData, 4)
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function DeleteKey(ByVal hKey As KeyREG, _
    ByVal sKey As String) As Long
    
    lReg = RegDeleteKey(hKey, sKey)
    
End Function

Public Function DeleteValue(ByVal hKey As KeyREG, _
    ByVal sPath As String, ByVal sValue As String) As Long
    
    lReg = RegOpenKey(hKey, sPath, KeyHandle)
    lReg = RegDeleteValue(KeyHandle, sValue)
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Sub FixReg()
    '-----------------------------------------'
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PaRaY_VM"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SCardSvr32"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "readme"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ConfigVir"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "NviDiaGT"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "NarmonVirusAnti"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "rundll16"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "AVManager"
    DeleteValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Internet Explorer\Main", "Window Title"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "kebodohan"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "pemalas"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "mulut_besar"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "otak_udang"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Run", "ssvchost"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Run", "Print Epson"
    DeleteValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Run", "Walpaper"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "Software\lutil\FMR", "svchost"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "Software\lutil\FMR", "Register"
    DeleteValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Group Policy", "svchost"
    DeleteValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Group Policy", "AppMgmt"
    DeleteValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Run", "4k51k4"
    DeleteValue HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Run", "MSMSGS"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion\Run", "System Monitoring"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"
    DeleteValue HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig"
    DeleteValue HKEY_CLASSES_ROOT, _
    "exefile", "NeverShowExt"
    SetDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun", 255
    SetDwordValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\PriorityControl", "Win32PrioritySeparation", 38
    SetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction", "Enable", "Y"
    SetStringValue HKEY_CURRENT_USER, "Control Panel\Desktop", "MenuShowDelay", "0"
    SetStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control", "WaitToKillServiceTimeout", "200"
    SetStringValue HKEY_CURRENT_USER, "Control Panel\Mouse", "MouseHoverTime", "10"
    
    
    '-----------------------------------------'
    SetStringValue HKEY_CLASSES_ROOT, _
        "exefile" & KEY_ROOT_EXE, _
        "", Chr$(34) & "%1" & Chr$(34) & " %*"
    SetStringValue HKEY_CLASSES_ROOT, _
        "lnkfile" & KEY_ROOT_EXE, _
        "", Chr$(34) & "%1" & Chr$(34) & " %*"
    SetStringValue HKEY_CLASSES_ROOT, _
        "piffile" & KEY_ROOT_EXE, _
        "", Chr$(34) & "%1" & Chr$(34) & " %*"
    SetStringValue HKEY_CLASSES_ROOT, _
        "batfile" & KEY_ROOT_EXE, _
        "", Chr$(34) & "%1" & Chr$(34) & " %*"
    SetStringValue HKEY_CLASSES_ROOT, _
        "scrfile" & KEY_ROOT_EXE, _
        "", Chr$(34) & "%1" & Chr$(34) & " %*"
    SetStringValue HKEY_CLASSES_ROOT, _
        "comfile" & KEY_ROOT_EXE, _
        "", Chr$(34) & "%1" & Chr$(34) & " %*"
    SetStringValue HKEY_CLASSES_ROOT, _
        "cmdfile" & KEY_ROOT_EXE, _
        "", Chr$(34) & "%1" & Chr$(34) & " %*"
    SetStringValue HKEY_CLASSES_ROOT, _
        "regfile" & KEY_ROOT_EXE, _
        "", "regedit.exe %1"
    '-----------------------------------------'
    SetDwordValue HKEY_CURRENT_USER, _
        "Software" & MSWCV & "Explorer\Advanced\", _
        "HideFileExt", 1
    SetDwordValue HKEY_CURRENT_USER, _
        "Software" & MSWCV & "Explorer\Advanced\", _
        "Hidden", 1
    SetDwordValue HKEY_CURRENT_USER, _
        "Software" & MSWCV & "Explorer\Advanced\", _
        "ShowSuperHidden", 1
    SetDwordValue HKEY_CURRENT_USER, _
        "Software" & MSWCV & "Policies\Explorer", _
        "NoFolderOptions", 0
    SetDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE" & MSWCV & "Policies\Explorer", _
        "NoFolderOptions", 0
    '-----------------------------------------'
    SetDwordValue HKEY_CURRENT_USER, _
        "Software" & MSWCV & "Policies\System", _
        "DisableRegistryTools", 0
    SetDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE" & MSWCV & "Policies\System", _
        "DisableRegistryTools", 0
    SetDwordValue HKEY_CURRENT_USER, _
        "Software" & MSWCV & "Policies\System\", _
        "DisableTaskMgr", 0
    SetDwordValue HKEY_LOCAL_MACHINE, _
        "SOFTWARE" & MSWCV & "Policies\System\", _
        "DisableTaskMgr", 0
End Sub



