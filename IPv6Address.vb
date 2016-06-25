' Module:   IPv6 Address Class
' Author:   James Merrill

Option Explicit On
Option Strict On
Public Class IPv6Address
  ' IP address and mask variables
  Private _ip As UInt128
  Private _mask As UInteger

  ' Shared variable because Classes can't be constants
  ' uxlFFFF represents a full hextet
  ' UInt128 Class does auto conversion from UInt
  Private Shared uxlFFFF As UInt128 = &HFFFFUI

  ' IP Address as string
  Public Property IP As UInt128
    Get
      Return _ip
    End Get
    Set(value As UInt128)
      ' No error checking is needed because there is no possible invalid value
      _ip = value
    End Set
  End Property

  ' Network Mask
  Public Property Mask As UInteger
    Get
      Return _mask
    End Get
    Set(value As UInteger)
      If value <= 128 Then
        _mask = value
      Else
        Throw New ArgumentOutOfRangeException
      End If
    End Set
  End Property

  ' Constructors
  ' Default
  Public Sub New()
    _ip = 0
    _mask = 0
  End Sub

  ' Copy
  Public Sub New(ByRef value As IPv6Address)
    ' Copy values from original object into new one
    _ip = value.IP
    _mask = value.Mask
  End Sub

  ' Inputs string and integer
  Public Sub New(ByVal IPAdd As String, ByVal IPMask As UInteger)
    ' This useage can result in a runtime error
    _ip = StrToIP(IPAdd)
    ' Mask property does range checking
    Mask = IPMask
  End Sub

  ' Inputs 128-bit and integer
  Public Sub New(ByVal IPAdd As UInt128, ByVal IPMask As UInteger)
    ' Copy values into new object
    _ip = IPAdd
    Mask = IPMask
  End Sub

  ' Calculate subnet IP address
  Public ReadOnly Property Subnet As UInt128
    Get ' Numeric IP ANDed with NetMask gives Subnet address
      ' Return new subnet IP address
      Return _ip And NetMask
    End Get
  End Property

  ' Mask Converts XX netmask to 128 bit number
  Public Property NetMask As UInt128
    Set(value As UInt128)
      ' Check for the 2 easy options
      If value = UInt128.MaxValue Then
        _mask = 128
      ElseIf value = UInt128.MinValue Then
        _mask = 0
      Else ' Count bits until the first 0
        Dim intCount As UInteger = 0
        Do While intCount < 128 AndAlso value.Bit(127 - CInt(intCount)) = 1
          intCount += 1UI
        Loop
        _mask = intCount
      End If
    End Set
    Get
      Dim uxlMask As UInt128 = 0
      ' Loop
      For i = 1 To CInt(_mask)
        ' Shift left 1 and add 1
        uxlMask = (uxlMask << 1) + 1
      Next
      uxlMask <<= CInt(128UI - _mask)
      Return uxlMask
    End Get
  End Property

  Public Property Hextet(index As Integer) As String
    Set(value As String)
      Dim intValue As UInteger = Convert.ToUInt32(value, 16)
      If index >= 1 AndAlso index <= 8 Then
        If intValue <= uxlFFFF Then
          ' Clear hextet bits and set them to new value - Cut value off at FFFF
          _ip = (_ip And Not (uxlFFFF << (16 * (8 - index)))) Or ((intValue And uxlFFFF) << (16 * (8 - index)))
        Else
          Throw New ArgumentOutOfRangeException
        End If
      Else
        Throw New IndexOutOfRangeException
      End If
    End Set
    Get
      ' Return value
      If index >= 1 AndAlso index <= 8 Then
        Return Hex(CType((_ip >> (16 * (8 - index))) And uxlFFFF, UInteger))
      Else
        Throw New IndexOutOfRangeException
      End If
    End Get
  End Property

  ' ToBinary converts 128 bits to binary string
  ' Converts IP address by default. Returns netmask if ConvertMask is true
  Public ReadOnly Property ToBinary(Optional convertMask As Boolean = False) As String
    Get
      ' Output string variable
      Dim strOut As String = ""
      ' Array of numbers, either IP address or Netmask
      Dim uxlNumber As UInt128 = If(convertMask, NetMask, _ip)
      ' 128 bits per number
      For i = 1 To 128
        ' Add a 1 or 0 to the string
        strOut += If(uxlNumber.Bit(128 - i) = 1, "1", "0")
      Next ' i
      Return strOut
    End Get
  End Property

  ' Compress IPv6 address
  Public ReadOnly Property Compressed As String
    Get
      ' Longest run of 0 value hextets
      Dim intLongest As Integer = 0
      ' Start of longest run of 0 value hextets
      Dim intNullHex As Integer = -1
      ' Length of current run of 0 value hextets
      Dim intCurrent As Integer = 0
      ' Start of current run of 0 value hextets
      Dim intCurrentStart As Integer
      ' Return result
      Dim strResult As String = ""
      For i = 1 To 8
        ' Look for "0" hextets
        If Hextet(i) = "0" Then
          ' Determine number of consecutive "0" hextets
          If intCurrent = 0 Then intCurrentStart = i
          intCurrent += 1
          intLongest = Math.Max(intLongest, intCurrent)
          ' Save current start position as longest
          If intLongest = intCurrent Then intNullHex = intCurrentStart
        Else
          ' Reset counter
          intCurrent = 0
        End If
      Next
      ' Cycle through hextets
      For i As Integer = 1 To 8
        If i = intNullHex + intLongest - 1 Then
          strResult += ":"
        ElseIf (i < intNullHex) Or (i >= intNullHex + intLongest) Then
          ' Add current hextet to string
          strResult += Hextet(i) + ":"
        End If
      Next
      ' Remove extra colon
      If strResult.Length > 0 Then strResult = strResult.Remove(strResult.Length - 1, 1)
      If intNullHex = 1 Then strResult = ":" + strResult
      If intNullHex + intLongest - 1 = 8 Then strResult += ":"
      ' Shave off extra 0 at beginning or end
      If strResult.Length > 2 AndAlso strResult.Substring(0, 3) = "0::" Then
        strResult = strResult.Remove(0, 1)
      ElseIf strResult.Length > 2 AndAlso strResult.Substring(strResult.Length - 3, 3) = "::0" Then
        strResult = strResult.Remove(strResult.Length - 1, 1)
      End If
      ' Return result
      Return strResult
    End Get
  End Property

  ' Expand IPv6 address
  Public ReadOnly Property Expanded As String
    Get
      ' Return String
      Dim strReturn As String = ""
      For i = 1 To 8
        ' Add colon before all but the first hextet
        If i > 1 Then strReturn += ":"
        ' Add current hextet to string
        strReturn += Hextet(i).PadLeft(4, "0"c)
      Next
      ' Return IP Address
      Return strReturn
    End Get
  End Property

  ' Check IPv6 address for validity
  Public Shared Function Valid(IPAddress As String) As Boolean
    ' Valid hex characters plus colon
    Const ValidChars As String = "0123456789ABCDEF:"
    ' Check for invalid characters
    For i = 0 To IPAddress.Length - 1
      If Not ValidChars.Contains(IPAddress.ToUpper.Substring(i, 1)) Then Return False
    Next
    ' Split IP into hextets
    Dim strHex() As String = IPAddress.Split(":"c)
    ' Too many?
    If strHex.Length > 8 Then Return False
    ' Short on Hextets and no ::
    If strHex.Length < 8 AndAlso Not strHex.Contains("") Then Return False
    ' Variable to track Blank hextet
    Dim blnFound As Boolean = False
    ' Cycle through hextets
    For i = 0 To strHex.GetUpperBound(0)
      ' If first hextet is blank and 2nd is not, return false
      If i = 0 AndAlso strHex(i) = "" AndAlso Not strHex(i + 1) = "" Then
        Return False
        ' If last hextet Is blank and previous one is not, return false
      ElseIf i = strHex.GetUpperBound(0) AndAlso strHex(i) = "" AndAlso Not strHex(i - 1) = "" Then
        Return False
        ' If current hextet is blank and a blank has been found and i is not 1 or upper bound,
        ' return false
      ElseIf strHex(i) = "" AndAlso blnFound AndAlso Not {1, strHex.GetUpperBound(0)}.Contains(i) Then
        Return False
        ' If hextet is blank, set blnFound to true
      ElseIf strHex(i) = "" Then
        blnFound = True
      ElseIf Convert.ToUInt32(strHex(i), 16) > &HFFFFUI Then
        Return False
      End If
    Next
    ' Passed all tests
    Return True
  End Function

  ' Calculate Summary IP address and keep it
  Public Sub Summarize(ByRef IPInput As IPv6Address)
    ' If either mask is zero, the summary is the entire internet
    If _mask = 0 OrElse IPInput.Mask = 0 Then
      _ip = 0
      _mask = 0
      ' If the values are the same, don't bother
    ElseIf IPInput._ip <> _ip Then
      Dim uxlNumX As UInt128
      uxlNumX = Not (_ip Xor IPInput._ip)
      _ip = NetMask And _ip And IPInput._ip
      Dim IPXOr As New IPv6Address(uxlNumX, 0)
      _mask = CUInt(IPXOr.ToBinary.IndexOf("0"))
    End If
  End Sub

  ' Remove element at index "index". Result is one element shorter.
  ' Similar to List.RemoveAt, but for arrays.
  ' From http://stackoverflow.com/questions/7169259/vb-net-removing-first-element-of-array
  Private Sub Remove(Of T)(ByRef a() As T, ByVal intPosition As Integer)
    ' Move elements after "index" down 1 position.
    Array.Copy(a, intPosition + 1, a, intPosition, a.GetUpperBound(0) - intPosition)
    ' Shorten by 1 element.
    ReDim Preserve a(a.GetUpperBound(0) - 1)
  End Sub

  ' Insert element at index "index". Result is 1 element longer.
  ' Similar to List.Insert, but for arrays.
  Private Sub Insert(Of T)(ByRef a() As T, ByVal intPosition As Integer)
    ' Lengthen by 1 element
    ReDim Preserve a(a.Length)
    ' Shift everything toward the end
    Array.Copy(a, intPosition, a, intPosition + 1, a.GetUpperBound(0) - intPosition)
  End Sub

  Public Function StrToIP(IPAdd As String) As UInt128
    Dim Hextets() As String = IPAdd.Split(":"c)
    Dim uxlIP As UInt128 = 0
    If Hextets.Length <= 8 Then
      Dim intSplit As Integer = -1
      If Hextets(0) = "" Then Remove(Hextets, 0)
      If Hextets(Hextets.GetUpperBound(0)) = "" Then Remove(Hextets, Hextets.GetUpperBound(0))
      For i = 0 To Hextets.GetUpperBound(0)
        If Hextets(i) = "" Then
          If intSplit = -1 Then
            intSplit = i
          Else
            Throw New InvalidCastException
          End If
        End If
      Next
      If Hextets.Length < 8 And intSplit = -1 Then Throw New InvalidCastException
      ' Hextets before the double colon
      For i = 0 To intSplit - 1
        Dim intHexet As UInteger = Convert.ToUInt32(Hextets(i), 16)
        uxlIP = uxlIP Or (CType(intHexet, UInt128) << ((7 - i) * 16))
      Next
      ' Hextets after double colon - All hextets if no double colon
      For i = intSplit + 1 To Hextets.GetUpperBound(0)
        Dim intIndex As Integer = i + (8 - Hextets.Length)
        Dim intHexet As UInteger = Convert.ToUInt32(Hextets(i), 16)
        uxlIP = uxlIP Or (CType(intHexet, UInt128) << ((7 - intIndex) * 16))
      Next
      Return uxlIP
    Else
      Throw New InvalidCastException
    End If
  End Function
End Class
