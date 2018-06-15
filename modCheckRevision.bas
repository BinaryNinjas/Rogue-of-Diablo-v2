Attribute VB_Name = "modCheckRevision"
'Battle.net Version Check Lib
' Rob@USEast

'CheckRevisionEx

    'FileExe
        '  path to the Game Exe file  ex. "c:\program files\starcraft\starcraft.exe"
    'FileStormDll
        '  path to storm.dll  ex. "c:\program files\starcraft\storm.dll"
    'FileBnetDll
        '  path to battle.snp  ex. "c:\program files\starcraft\battle.snp"
    'HashText
        '  The value string from battle.net
    'Version
        '  the returned version for the game exe
    'Checksum
        '  The resulting Checksum
    'ExeInfo
        '  The resulting Exeinfo for pre-lockdown version checks
        '  The Digest for lockdown version checks
    'PathToDLL
        '  For pre-lockdown
            '  The name of the version dll provided by battle.net
        '  For lockdown
            '  The path to the version check dll
            '  ex.  "lockdown\lockdown-IX86-00.dll"
            '  This is required.  The dll is hashed along with the game files
    'PathToLockdown01
        '  No longer Required.  Left it for backwards compatability
        '  You can pass a NULL string.
    'PathToVideoBin
        '  Path to the Video dump file
        '  Will now use both the 64k dumps and the 10k dumps
        '  Included are all dumps required
        '  STAR.bin - Works for Starcraft, Broodwar, Starcraft Japan, Starcraft Shareware
        '  W2BN.bin - Works for Warcraft II BNE
        '  DRTL.bin - Works for Diablo, Diable Shareware


' Will attempt to detect which CheckRevision to use
Public Declare Function CheckRevisionEx Lib "CheckRevision.dll" (ByVal GameFile1 As String, ByVal GameFile2 As String, ByVal GameFile3 As String, ByVal ValueString As String, ByRef version As Long, ByRef Checksum As Long, ByVal exeinfo As String, ByVal PathToDLL As String, ByVal sUnused As String, ByVal PathToVideoBin As String) As Long
' lockdown implementation
Public Declare Function CheckRevisionLD Lib "CheckRevision.dll" (ByVal GameFile1 As String, ByVal GameFile2 As String, ByVal GameFile3 As String, ByVal ValueString As String, ByRef version As Long, ByRef Checksum As Long, ByVal exeinfo As String, ByVal PathToDLL As String, ByVal sUnused As String, ByVal PathToVideoBin As String) As Long
' IX86ver#
Public Declare Function CheckRevisionA Lib "CheckRevision.dll" (ByVal GameFile1 As String, ByVal GameFile2 As String, ByVal GameFile3 As String, ByVal ValueString As String, ByRef version As Long, ByRef Checksum As Long, ByVal exeinfo As String, ByVal DLLName As String) As Long
' ver-IX86
Public Declare Function CheckRevisionB Lib "CheckRevision.dll" (ByVal GameFile1 As String, ByVal GameFile2 As String, ByVal GameFile3 As String, ByVal ValueString As String, ByRef version As Long, ByRef Checksum As Long, ByVal exeinfo As String, ByVal DLLName As String) As Long

