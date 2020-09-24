VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'CryptAcquireContext(
'HCRYPTPROV *phProv,    [out] The address to which the function copies a handle to the CSP.
'LPCTSTR pszContainer,  [in]  The key container name. This is a zero-terminated string that identifies the key container to the CSP
'                               vbNullChar used in this example to get the default key container.
'LPCTSTR pszProvider,   [in] The provider name. This is a zero-terminated string that specifies the CSP to be used
'DWORD dwProvType,      [in] The type of provider to acquire
'DWORD dwFlags          [in] The flag values
') As Long

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    phProv As Long, pszContainer As String, pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
    
'CryptReleaseContext(
'HCRYPTPROV hProv,  [in] A handle to the application’s CSP
'DWORD dwFlags      [in] The flag values. This parameter is reserved for future use and should always be zero
');
Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
    (ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

'CryptCreateHash(
'    HCRYPTPROV hProv,      [in] A handle to the CSP to use. An application obtains this handle using the CryptAcquireContext function.
'    ALG_ID Algid,          [in] An algorithm identifier of the hash algorithm to use
'    HCRYPTKEY hKey,        [in] For nonkeyed algorithms, this parameter should be set to zero
'    DWORD dwFlags,         [in] The flag values. This parameter is reserved for future use and should always be zero
'    HCRYPTHASH *phHash);   [out] The address to which the function copies a handle to the new hash object.
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, _
    ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long

'CryptDestroyHash(
'    HCRYPTHASH hHash [in] A handle to the hash object to be destroyed
');
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long

'CryptHashData(
'    HCRYPTHASH hHash,      [in] A handle to the hash object. An application obtains this handle using the CryptCreateHash function.
'    CONST BYTE *pbData,    [in] The address of the data to be hashed
'    DWORD dwDataLen,       [in] The number of bytes of data to be hashed.
'    DWORD dwFlags);        [in] The flag values (The Microsoft RSA Base Provider ignores this parameter)
Private Declare Function CryptHashData Lib "advapi32.dll" ( _
ByVal hHash As Long, ByVal pbData As String, _
ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
    
'CryptGetHashParam(
'   HCRYPTHASH hHash,   [in] A handle to the hash object on which to query parameters
'   DWORD dwParam,      [in] The parameter number
'   BYTE *pbData,       [out] The parameter data buffer
'   DWORD *pdwDataLen,  [in/out] The address of the parameter data length.
'   DWORD dwFlags       [in] The flag values. This parameter is reserved for future use and should always be zero.
');
Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
(ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, _
 pdwDataLen As Long, ByVal dwFlags As Long) As Long
 
Private Declare Function GetLastError Lib "kernel32" () As Long
' Name of the provider shipped with Windows by default
' Both support below declared hashing algorithms.
' MS_DEF_PROV is bundled with the operating system.
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const MS_ENHANCED_PROV = "Microsoft Enhanced Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL = 1
Private Const CRYPT_NEWKEYSET = &H8
'// Hash sub ids
'#define ALG_SID_MD2                     1
Private Const ALG_SID_MD2 = 1
'#define ALG_SID_MD4
Private Const ALG_SID_MD4 = 2
'#define ALG_SID_MD5                     3
Private Const ALG_SID_MD5 = 3
'#define ALG_SID_SHA                     4
Private Const ALG_SID_SHA = 4

' // Hash classes
'#define ALG_CLASS_HASH                  (4 << 13)
Private Const ALG_CLASS_HASH = 32768

'// Algorithm types
'#define ALG_TYPE_ANY                    (0)
Private Const ALG_TYPE_ANY = 0

'// algorithm identifier definitions
'#define CALG_MD2                (ALG_CLASS_HASH | ALG_TYPE_ANY | ALG_SID_MD2)
Private Const CALG_MD2 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
'#define CALG_MD4                (ALG_CLASS_HASH | ALG_TYPE_ANY | ALG_SID_MD4)
Private Const CALG_MD4 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
'#define CALG_MD5        (ALG_CLASS_HASH | ALG_TYPE_ANY | ALG_SID_MD5)
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
'#define CALG_SHA        (ALG_CLASS_HASH | ALG_TYPE_ANY | ALG_SID_SHA)
Private Const CALG_SHA = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)

'// CryptGetHashParam parameter number values
'#define HP_ALGID                0x0001  // Hash algorithm
Private Const HP_ALGID = 1
'#define HP_HASHVAL              0x0002  // Hash value
Private Const HP_HASHVAL = 2
'#define HP_HASHSIZE             0x0004  // Hash value size
Private Const HP_HASHSIZE = 4
'
' Block size to read from file to hash
Private Const HP_FILE_BLOCKSIZE = 160
' ERROR Constants
' For cryptHashData
Private Const ERROR_INVALID_HANDLE = 6&
Private Const ERROR_INVALID_PARAMETER = 87 '  dderror
Private Const ERROR_NTE_BAD_UID = &H80090001
Private Const ERROR_BAD_HASH = &H80090002
Private Const ERROR_BAD_LEN = &H80090004
Private Const ERROR_NTE_BAD_ALGID = &H80090008
Private Const ERROR_NTE_BAD_HASH_STATE = &H8009000C
Private Const ERROR_NTE_NO_MEMORY = &H8009000E
Private Const ERROR_NTE_FAIL = &H80090020

Public Enum HASH_ALG_ID
    HP_MD2
    HP_MD4
    HP_MD5
    HP_SHA
End Enum
Public Enum HASH_DATA_FORMAT
    HP_HEXADECIMAL
    HP_NUMERIC
    HP_ASCII
End Enum
Private Enum HASH_STATUS
    HP_STATUS_NONE
    HP_STATUS_READY
    HP_STATUS_BUSY
End Enum


' Private variables
Private m_Hash_Object As Long
Private m_Hash_Data(20) As Byte 'This value will usually be 16 or 20, depending on the hash algorithm.
Private m_Hash_DataLen As Long
Private m_Hash_Algo_Id As HASH_ALG_ID
Private m_Hash_Algorithm As Long
Private m_Hash_Status As HASH_STATUS
Private m_CSP_Provider As Long
Private m_Hash_DataReady As Boolean

Private Sub Class_Initialize()
'    m_Hash_Status = HP_STATUS_NONE
'    m_Hash_Algo_Id = HP_MD5
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    If m_Hash_Object Then CryptDestroyHash m_Hash_Object
    If m_CSP_Provider Then CryptReleaseContext m_CSP_Provider, 0
End Sub

Public Property Get Algorithm() As EC_HASH_ALG_ID
    Algorithm = m_Hash_Algo_Id
End Property
Public Property Let Algorithm(ByVal algNewValue As EC_HASH_ALG_ID)
    m_Hash_Algo_Id = algNewValue
End Property
Public Property Get IsDataReady() As Boolean
    IsDataReady = m_Hash_DataReady
End Property
' Read only property
Public Property Get HashObject() As Long
    Select Case m_Hash_Status
        Case HP_STATUS_NONE:
            Err.Raise vbObjectError + 1001, "Hash Provider", _
            "The Hash object has not been created yet."
        Case HP_STATUS_BUSY Or HP_STATUS_READY:
            HashObject = m_Hash_Object
    End Select
End Property

Public Function CreateByCSPRef(CSPProvider As Long) As Long
    If m_Hash_Object Then CryptDestroyHash m_Hash_Object
    ' I am very optimistic
    CreateByCSPRef = 1
    Select Case m_Hash_Algo_Id
        Case HP_MD2: m_Hash_Algorithm = CALG_MD2
        Case HP_MD4: m_Hash_Algorithm = CALG_MD4
        Case HP_MD5: m_Hash_Algorithm = CALG_MD5
        Case HP_SHA: m_Hash_Algorithm = CALG_SHA
    End Select
    If m_Hash_Object Then CryptDestroyHash m_Hash_Object
    If Not CBool(CryptCreateHash(CSPProvider, m_Hash_Algorithm, 0, 0, m_Hash_Object)) Then
        m_Hash_Status = HP_STATUS_NONE
        CreateByCSPRef = 0
        Err.Raise vbObjectError + 1002, "Hash Provider", _
                  "An error has occurred creating hash object"
    Else
        m_Hash_Status = HP_STATUS_READY
    End If
End Function

Public Function CreateHashObject() As Long

    If m_Hash_Object Then CryptDestroyHash m_Hash_Object
    
    Dim sProvider As String
    Dim sContainer As String
    Select Case m_Hash_Algo_Id
        Case HP_MD2: m_Hash_Algorithm = CALG_MD2
        Case HP_MD4: m_Hash_Algorithm = CALG_MD4
        Case HP_MD5: m_Hash_Algorithm = CALG_MD5
        Case HP_SHA: m_Hash_Algorithm = CALG_SHA
    End Select
    
    sContainer = vbNullChar
    sProvider = MS_ENHANCED_PROV & vbNullChar
    ' I am very optimistic
    CreateHashObject = 1
    If m_CSP_Provider = 0 Then
         'Attempt to acquire a handle to the default key container.
        If Not CBool(CryptAcquireContext(m_CSP_Provider, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, 0)) Then
            ' Attempt to get a handle to the enhanced key container
            sProvider = MS_DEF_PROV & vbNullChar
            If Not CBool(CryptAcquireContext(m_CSP_Provider, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, 0)) Then
                'If it is not possible to get a handle to none of those
                '[default] OP containers, raise an error as
                'the creation of key containers is out of the working
                'scope of this object.
                CreateHashObject = 0
                Err.Raise vbObject + 1003, "Hash Provider", "Error getting a handle to key containers"
            End If
        End If
    End If
    'Attempt to acquire a handle to a Hash object
    If Not CBool(CryptCreateHash(m_CSP_Provider, m_Hash_Algorithm, _
            0, 0, m_Hash_Object)) Then
            CreateHashObject = 0
            Err.Raise vbObject + 1002, "Hash Provider", "An error has occurred creating hash object"
    End If
    m_Hash_Status = HP_STATUS_READY
End Function

Public Function GetHashObject() As Long
''    If m_Hash_Status = HP_STATUS_READY Or m_Hash_Status = HP_STATUS_BUSY Then
        GetHashObject = m_Hash_Object
''    Else
''        GetHashObject = 0
''    End If
End Function
Public Function DigestDataFile(ByVal sFilePath As String) As Long
On Error GoTo ErrHandler
m_Hash_Status = HP_STATUS_BUSY
Dim fNum As Long, fLen As Long, fBlockBytes As Long, fLostBytes As Long
Dim fdat() As Byte, fsDat As String
Dim iCounter As Integer, jCounter As Integer
Dim lResult As Long
fNum = FreeFile()
' Optimistic
DigestDataFile = 1
' First we check if the file exists
If Trim(Dir(sFilePath)) = "" Then
    Err.Raise vbObjectError + 1007, "Hash Provider", "File not found"
End If
Open sFilePath For Binary Access Read As #fNum
    fLen = LOF(fNum)
    If fLen < HP_FILE_BLOCKSIZE Then
        ' Resize dynamic array
        ReDim fdat(fLen)
        ' Get the data
        Get #fNum, , fdat()
        ' Convert to string one by one
        For iCounter = 0 To fLen - 1
            fsDat = fsDat & Chr(fdat(iCounter))
        Next
        ' And finally digest the data
        'lResult = DigestDataBlock(fsDat)
        If Not CBool(CryptHashData(m_Hash_Object, fsDat, fLen, 0)) Then
           DigestDataFile = 0
           Err.Raise vbObjectError + 1005, "Hash Provider", _
                "An error occurred digesting the data."
                m_Hash_Status = HP_STATUS_READY '[?]
        End If
    Else
        ' Find out how many blocks the file holds
        fBlockBytes = fLen / HP_FILE_BLOCKSIZE
        ' Find out how many bytes are lost
        fLostBytes = fLen Mod HP_FILE_BLOCKSIZE
        ' Now loop the number of blocks
        ' [First resize array]
        ReDim fdat(HP_FILE_BLOCKSIZE)
        For iCounter = 1 To fBlockBytes
            ' Now get the data
            Get #fNum, , fdat()
            ' Convert to string one by one
            For jCounter = 0 To HP_FILE_BLOCKSIZE
                fsDat = fsDat & Chr(fdat(jCounter))
            Next
            ' Digest the data
            'lResult = DigestDataBlock(fsDat)
            If Not CBool(CryptHashData(m_Hash_Object, fsDat, HP_FILE_BLOCKSIZE, 0)) Then
               DigestDataFile = 0
               Err.Raise vbObjectError + 1005, "Hash Provider", _
                    "An error occurred digesting the data."
                    m_Hash_Status = HP_STATUS_READY '[?]
            End If
            ' Now clear the buffer string to do it again
            fsDat = ""
        Next
        If fLostBytes <> 0 Then
            ' Process lost bytes [bytes remaining]
            ReDim fdat(fLostBytes)
            ' Get the remaining data
            Get #fNum, , fdat()
            ' Convert to string one by one
            For jCounter = 0 To fLostBytes - 1
                fsDat = fsDat & Chr(fdat(jCounter))
            Next
            ' Digest the data [BURP!-- sorry ;~)]
            'lResult = DigestDataBlock(fsDat)
            If Not CBool(CryptHashData(m_Hash_Object, fsDat, fLostBytes, 0)) Then
               DigestDataFile = 0
               Err.Raise vbObjectError + 1005, "Hash Provider", _
                    "An error occurred digesting the data."
                    m_Hash_Status = HP_STATUS_READY '[?]
            End If
        End If
    End If
Close #fNum
Call SetHashData
m_Hash_Status = HP_STATUS_READY
Exit Function
ErrHandler:
    ' Raise error
    Err.Raise Err.Number, "Hash Provider", Err.Description
End Function
Public Function DigestDataBlock(ByVal sData As String) As Long
Dim lDataLen As Long
    If m_Hash_Status = HP_STATUS_NONE Then
        Err.Raise vbObjectError + 1001, "Hash Provider", _
            "The Hash object has not been created yet."
    End If
    m_Hash_Status = HP_STATUS_BUSY
    ' I am very optimistic
    DigestDataBlock = 1
    lDataLen = Len(sData)
    If Not CBool(CryptHashData(m_Hash_Object, sData, lDataLen, 0)) Then
        'ERROR_INVALID_HANDLE   ' One of the parameters specifies an invalid handle.
        'ERROR_INVALID_PARAMETER    ' One of the parameters contains an invalid value. This is most often an illegal pointer.
        'NTE_BAD_ALGID   ' The hHash handle specifies an algorithm that this CSP does not support.
        'NTE_BAD_FLAGS       ' The dwFlags parameter contains an invalid value.
        'NTE_BAD_HASH        ' The hash object specified by the hHash parameter is invalid.
        'NTE_BAD_HASH_STATE  ' An attempt was made to add data to a hash object that is already marked “finished.”
        'NTE_BAD_KEY     ' A keyed hash algorithm is being used, but the session key is no longer valid. This error will be generated if the session key is destroyed before the hashing operating is complete.
        'NTE_BAD_LEN     ' The CRYPT_USERDATA flag is set and the dwDataLen parameter has a nonzero value.
        'NTE_BAD_UID     ' The CSP context that was specified when the hash object was created cannot be found.
        'NTE_FAIL        ' The function failed in some unexpected way.
        'NTE_NO_MEMORY   ' The CSP ran out of memory during the operation
        DigestDataBlock = 0
        Err.Raise vbObjectError + 1005, "Hash Provider", _
            "An error occurred processing the data."
        m_Hash_Status = HP_STATUS_READY '[?]
    End If
    Call SetHashData
    m_Hash_Status = HP_STATUS_READY
End Function

Public Function GetHashData(Optional ByVal hpDataFormat As HASH_DATA_FORMAT = HP_HEXADECIMAL) As String
Dim sData As String, sHex As String
Dim iCounter As Integer
If m_Hash_Status = HP_STATUS_NONE Then
    Err.Raise vbObjectError + 1001, "Hash Provider", _
            "The Hash object has not been created yet."
End If
If m_Hash_DataLen = 0 Then
    GetHashData = "Nothing has been digested yet!"
End If
m_Hash_Status = HP_STATUS_BUSY
    Select Case hpDataFormat
    Case HP_HEXADECIMAL
        For iCounter = 0 To m_Hash_DataLen - 1
            sHex = Hex(m_Hash_Data(iCounter))
            If Len(sHex) > 1 Then
                sData = sData & sHex & vbTab
            Else
                sData = sData & "0" & sHex & vbTab
            End If
            sHex = ""
        Next
    Case HP_NUMERIC
        For iCounter = 0 To m_Hash_DataLen - 1
            sData = sData & CStr(m_Hash_Data(iCounter))
        Next
    Case HP_ASCII
        'Open "C:\Hash.txt" For Binary As #1
        For iCounter = 0 To m_Hash_DataLen - 1
        
            sData = sData & Chr(m_Hash_Data(iCounter))
        Next
        'Put #1, , sData
       'Close #1
    End Select
    GetHashData = sData
    m_Hash_Status = HP_STATUS_READY
End Function
Private Sub SetHashData()
    m_Hash_DataReady = False
    If m_Hash_Status = HP_STATUS_NONE Then Exit Sub ' For now
        ' Make a first call to find out the size needed for the buffer
'    Call CryptGetHashParam(m_Hash_Object, HP_HASHSIZE, bDataLen(0), lDataLen, 0)
    'If lDataLen = 0 Then
    '    Err.Raise vbObjectError + 1006, "Hash Provider", "An error has occurred when getting hash value."
        ' Raise Error?
   ' Else
   m_Hash_DataLen = 20&
   If Not CBool(CryptGetHashParam(m_Hash_Object, HP_HASHVAL, m_Hash_Data(0), _
                  m_Hash_DataLen, 0)) Then
                  'API error function
            m_Hash_DataLen = 0
            ' Raise Error?
    '       ERROR_INVALID_HANDLE One of the parameters specifies an invalid handle.
    '       ERROR_INVALID_PARAMETER One of the parameters contains an invalid value. This is most often an illegal pointer.
    '       NTE_BAD_FLAGS The dwFlags parameter is nonzero.
    '       NTE_BAD_HASH The hash object specified by the hHash parameter is invalid.
    '       NTE_BAD_TYPE The dwParam parameter specifies an unknown parameter number.
    '       NTE_BAD_UID The CSP context that was specified when the hash was created cannot be found.
            Err.Raise vbObjectError + 1006, "Hash Provider", "An error has occurred when getting hash value."
   End If
   m_Hash_DataReady = True
End Sub
