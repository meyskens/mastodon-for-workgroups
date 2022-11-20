Attribute VB_Name = "Module1"
Option Explicit
'==========
'URLUtility
'==========
'
'Adding the URLUtility class to a VB6 project produces a predeclared
'global object named URLUtility to your program.  You can call methods
'on this object to URLDecode and URLEncode String values.
'
'
'Note the hack implemented here for "+ encoding" of spaces in the query
'portion of a URL.  By rights everything following the ? or # in a URL
'should be passed literally.  Encoding/decoding this string is NOT part
'of the URL encode/decode process, but is part of building and parsing
'the parameter string.
'
'As a result, the API calls used here do not process the query portion
'of the URL, assuming this has already been done/will be done later as
'required.  The hack used here implements a common substitution of
'spaces by "+" characters after converting any "+" characters to "%2B"
'sequences.
'

Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = 0
Private Const INTERNET_MAX_URL_LENGTH As Long = 2048
Private Const URL_ESCAPE_PERCENT As Long = &H1000&

Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeA" ( _
    ByVal pszURL As String, _
    ByVal pszEscaped As String, _
    ByRef pcchEscaped As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeA" ( _
    ByVal pszURL As String, _
    ByVal pszUnescaped As String, _
    ByRef pcchUnescaped As Long, _
    ByVal dwFlags As Long) As Long

Public Function URLDecode( _
    ByVal URL As String, _
    Optional ByVal PlusSpace As Boolean = True) As String
    
    Dim cchUnescaped As Long
    Dim HRESULT As Long
    
    If PlusSpace Then URL = Replace$(URL, "+", " ")
    cchUnescaped = Len(URL)
    URLDecode = String$(cchUnescaped, 0)
    HRESULT = UrlUnescape(URL, URLDecode, cchUnescaped, 0)
    If HRESULT = E_POINTER Then
        URLDecode = String$(cchUnescaped, 0)
        HRESULT = UrlUnescape(URL, URLDecode, cchUnescaped, 0)
    End If
    
    If HRESULT <> S_OK Then
        Err.Raise Err.LastDllError, "URLUtility.URLDecode", _
                  "System error"
    End If
    
    URLDecode = Left$(URLDecode, cchUnescaped)
End Function

Public Function URLEncode( _
    ByVal URL As String, _
    Optional ByVal SpacePlus As Boolean = True) As String
    
    Dim cchEscaped As Long
    Dim HRESULT As Long
    
    If Len(URL) > INTERNET_MAX_URL_LENGTH Then
        Err.Raise &H8004D700, "URLUtility.URLEncode", _
                  "URL parameter too long"
    End If
    
    cchEscaped = Len(URL) * 1.5
    URLEncode = String$(cchEscaped, 0)
    HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    If HRESULT = E_POINTER Then
        URLEncode = String$(cchEscaped, 0)
        HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    End If

    If HRESULT <> S_OK Then
        Err.Raise Err.LastDllError, "URLUtility.URLEncode", _
                  "System error"
    End If
    
    URLEncode = Left$(URLEncode, cchEscaped)
    If SpacePlus Then
        URLEncode = Replace$(URLEncode, "+", "%2B")
        URLEncode = Replace$(URLEncode, " ", "+")
    End If
End Function



