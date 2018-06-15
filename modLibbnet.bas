Attribute VB_Name = "modLibbnet"
Option Explicit

'' nls functions

' nls_init
' Initializes the nls session.  Returns a pointer to the NLS object.
' nlsPointer = nls_init("Username", "Password")
Public Declare Function nls_init Lib "libbnet.dll" (ByVal sUsername As String, ByVal sPassword As String) As Long

' nls_reinit
' Reinitializes an existing NLS object with a new username and password.
' newPointer = nls_reinit(nlsPointer, "Username", "Password")
Public Declare Function nls_reinit Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sUsername As String, ByVal sPassword As String) As Long

' nls_free
' Destroy an existing NLS object.
' nls_free(nlsPointer)
Public Declare Sub nls_free Lib "libbnet.dll" (ByVal lNLSPointer As Long)

' nls_account_logon
' Creates packet SID_AUTH_ACCOUNTLOGON (0x53).  Append the packet header to the sBufferOut and send to BNET.
' sBufferOut needs to have enough space to hold Len(Username) + 33
' Returns the actual length of the data in sBufferOut or 0 if failure.
' Dim sBufferOut as String
' Dim lReturnLen as Long
' sBufferOut = Space(Len(Username) + 33)
' lReturnLen = nls_account_logon(nlsPointer, sBufferOut)
Public Declare Function nls_account_logon Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long

' nls_account_logon_proof
' Creates packet SID_AUTH_ACCOUNTLOGONPROOF (0x54).  Append the packet header to the sBufferOut and send to BNET.
' sBufferOut needs to have enough space to hold 20 Bytes
' sSalt is the first 32 Bytes of the S->C SID_AUTH_ACCOUNTLOGON (0x53) packet
' sServerKey is the second 32 Bytes of the S->C SID_AUTH_ACCOUNTLOGON (0x53) packet
' Dim sBufferOut as String * 20
' nls_account_logon_proof(nlsPointer, sBufferOut, sServerKey, sSalt)
Public Declare Sub nls_account_logon_proof Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String, ByVal sServerKey As String, ByVal sSalt As String)

' nls_account_create
' Creates packet SID_AUTH_ACCOUNTCREATE (0x52).  Append the packet header to the sBufferOut and send to BNET.
' sBufferOut needs to have enough space to hold Len(Username) + 65
' Returns the actual length of the data in sBufferOut or 0 if failure.
' Dim sBufferOut as String
' Dim lReturnLen as Long
' sBufferOut = Space(Len(Username) + 33)
' lReturnLen = nls_account_create(nlsPointer, sBufferOut)
Public Declare Function nls_account_create Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long

' nls_account_change
' Creates packet SID_AUTH_ACCOUNTCHANGE (0x55).  Append the packet header to the sBufferOut and send to BNET.
' sBufferOut needs to have enough space to hold Len(Username) + 33.
' Returns the actual length of the data in sBufferOut or 0 if failure.
' Dim sBufferOut as String
' Dim lReturnLen as Long
' sBufferOut = Space(Len(Username) + 33)
' lReturnLen = nls_account_change(nlsPointer, sBufferOut)
Public Declare Function nls_account_change Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long

' nls_account_change_proof
' Creates packet SID_AUTH_ACCOUNTCHANGEPROOF (0x56).  Append the packet header to the sBufferOut and send to BNET.
' sBufferOut needs to have enough space to hold 84 Bytes.
' sSalt is the first 32 Bytes of the S->C SID_AUTH_ACCOUNTCHANGE (0x55) packet.
' sServerKey is the second 32 Bytes of the S->C SID_AUTH_ACCOUNTCHANGE (0x55) packet.
' Returns a new NLS Object.
' Dim sBufferOut * 84
' nlsPointer = nls_account_change_proof(oldNLsPointer, sBufferOut, sNewPassword, sServerKey, sSalt)
Public Declare Function nls_account_change_proof Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String, ByVal sNewPassword As String, ByVal sServerKey As String, ByVal sSalt As String) As Long

' nls_account_upgrade_proof
' This is defunct at this time.
' Creates packet SID_AUTH_ACCOUNTUPGRADEPROOF (0x58).  Append the packet header to the sBufferOut and send to BNET.
' sBufferOut needs to have enough space to hold 64 Bytes
' Returns the actual length of the data in sBufferOut or 0 if failure.
' Dim sBufferOut as String * 64
' lReturnLen = nls_account_upgrade_proof(nlsPointer, sBufferOut)
Public Declare Function nls_account_upgrade_proof Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long

'' checkrevision functions

' checkrevision_ld
' Performs a Lockdown checkrevision
    'sFile1
        '  path to the first required game file  ex. "c:\program files\starcraft\starcraft.exe"
    'sFile2
        '  path to the second required game file  ex. "c:\program files\starcraft\storm.dll"
    'sFile3
        '  path to third required game file  ex. "c:\program files\starcraft\battle.snp"
    'sValueString
        '  The value string from SID_AUTH_INFO (0x50)
    'lVersion
        '  the returned version for the game exe
    'lChecksum
        '  The resulting Checksum
    'sReturnDigest
        '  The Digest for lockdown version checks
    'sLockDownFile
        '  The path to the version check dll
        '  ex.  "lockdown\lockdown-IX86-00.dll"
        '  This is required.  The dll is hashed along with the game files
    'sVideoFile
        '  Path to the Video dump file
        '  Will use both the 64k dumps and the 10k dumps
        '  Included are all dumps required
        '  STAR.bin - Works for Starcraft, Broodwar, Starcraft Japan, Starcraft Shareware
        '  W2BN.bin - Works for Warcraft II BNE
        '  DRTL.bin - Works for Diablo, Diable Shareware
Public Declare Function checkrevision_ld Lib "libbnet.dll" (ByVal sFile1 As String, ByVal sFile2 As String, ByVal sFile3 As String, ByVal sValueString As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByVal sReturnDigest As String, ByVal sLockdownFile As String, ByVal sVideoFile As String) As Long

' checkrevision_ld_raw_video
' Performs a lockdown checkrevision.
' Same as above except for the final 2 parameters'
' sVideoBuffer is the raw data from the video buffer file.
' lVideoBufferLen is the length of the above data.
Public Declare Function checkrevision_ld_raw_video Lib "libbnet.dll" (ByVal sFile1 As String, ByVal sFile2 As String, ByVal sFile3 As String, ByVal sValueString As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByVal sReturnDigest As String, ByVal sLockdownFile As String, ByVal sVideoBuffer As String, ByVal lVideoBufferLen As Long) As Long

' checkrevision
' Performs a ver-IX86* checkrevision
    'sFile1
        '  path to the first required game file  ex. "c:\program files\starcraft\starcraft.exe"
    'sFile2
        '  path to the second required game file  ex. "c:\program files\starcraft\storm.dll"
    'sFile3
        '  path to third required game file  ex. "c:\program files\starcraft\battle.snp"
    'sValueString
        '  The value string from SID_AUTH_INFO (0x50)
    'lVersion
        '  the returned version for the game exe
    'lChecksum
        '  The resulting Checksum
    'sExeInfo
        '  The Exe Information of the game exe.
    'sMPQName
        ' The MPQ name from SID_AUTH_INFO 0x50
Public Declare Function checkrevision Lib "libbnet.dll" (ByVal sFile1 As String, ByVal sFile2 As String, ByVal sFile3 As String, ByVal sValueString As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByVal sExeInfo As String, ByVal sMPQName As String) As Long

'' cdkey functions

' decode_hash_cdkey
' Decodes and hashes the cdkey in one function.
' Works for all current BNET cdkeys.
' sBufferOut needs to have enough space to hold 20 bytes
' Returns 0 for failure 1 for success.
' Dim sBufferOut as String * 20
' Dim lPublicValue as Long, lProductID as Long
' decode_hash_cdkey("1234567890123",lClientToken, lServerToken, lPublicValue, lProductID, sBufferOut)
Public Declare Function decode_hash_cdkey Lib "libbnet.dll" (ByVal sCDKey As String, ByVal lClientToken As Long, ByVal lServerToken As Long, ByRef lPublicValue As Long, ByRef lProductID As Long, ByVal sBufferOut As String) As Long

' decode_hash_cdkey_36
' same as the above function, but for use when using SID_CDKEY2 (0x36).
Public Declare Function decode_hash_cdkey_36 Lib "libbnet.dll" (ByVal sCDKey As String, ByVal lClientToken As Long, ByVal lServerToken As Long, ByRef lPublicValue As Long, ByRef lProductID As Long, ByVal sBufferOut As String) As Long

'' broken sha1 functions

' double_hash_password
' Performs a double hash on your password.
' Used when logging in to BNET.
Public Declare Sub double_hash_password Lib "libbnet.dll" (ByVal sPassword As String, ByVal lClientToken As Long, ByVal lServerToken As Long, ByVal sBufferOut As String)

' hash_password
' Performs a single hash on your password.
' Used when creating accounts.
Public Declare Sub hash_password Lib "libbnet.dll" (ByVal sPassword As String, ByVal sBufferOut As String)

'' warden functions

' warden_create_incoming_key
' Create the incoming key needed for warden crypt
' sBufferOut needs enough space to hold 0x102 bytes
' lSeed is the first DWORD of the cdkey
Public Declare Sub warden_create_incoming_key Lib "libbnet.dll" (ByVal sBufferOut As String, ByVal lSeed As Long)

' warden_create_outgoing_key
' Create the incoming key needed for warden crypt
' sBufferOut needs enough space to hold 0x102 bytes
' lSeed is the first DWORD of the cdkey
Public Declare Sub warden_create_outgoing_key Lib "libbnet.dll" (ByVal sBufferOut As String, ByVal lSeed As Long)

' warden_crypt
' performs the warden RC4 encryption
' sKey is the key to use, this will be modified after each call
' sData is the data to encryot/decrypt
' sBufferOut needs to have enough space to hold Len(sData)
' lLength is Len(sData)
Public Declare Sub warden_crypt Lib "libbnet.dll" (ByVal sKey As String, ByVal sData As String, ByVal sBufferOut As String, ByVal lLength As Long)




