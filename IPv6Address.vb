' Module:   IPv6 Address Class
' Author:   James Merrill

Option Explicit On
Option Strict On
Public Class IPv6Address
  ' IP address and mask variables
  Private _ip As String
  Private _mask As UInteger

  ' IP Address as string
  Public Property IP As String
    Get
      Return _ip
    End Get
    Set(value As String)
      _ip = value
    End Set
  End Property

  ' Network Mask
  Public Property Mask As UInteger
    Get
      Return _mask
    End Get
    Set(value As UInteger)
      _mask = value
    End Set
  End Property

  ' Inputs string and integer
  Public Sub New(ByVal IPAdd As String, ByVal IPMask As UInteger)
    ' Copy values into new object
    _ip = IPAdd
    _mask = IPMask
  End Sub

  ' Inputs string and string
  Public Sub New(ByVal IPAdd As String, ByVal IPMask As String)
    ' Copy values into new object
    _ip = IPAdd
    ' Convert mask to integer first
    UInteger.TryParse(IPMask, _mask)
  End Sub

  ' Input to duplicate 
  Public Sub New(ByRef IPAdd As IPv6Address)
    ' Copy values from original object into new one
    _ip = IPAdd.IP
    _mask = IPAdd.Mask
  End Sub

  Public ReadOnly Property Numeric As UInt128
    Get
      ' Numeric formatted IP address to 128 bits
      Dim uxlReturn As UInt128 = 0
      ' Gather Hextet values and multiply by powers of 256
      For i = 0 To 7 ' Cycle through 8 hextets
        ' Shift and Or for each hextet
        uxlReturn = uxlReturn Or (CType(Convert.ToUInt64(If(Hextet(i + 1) = "", "0", Hextet(i + 1)), 16), UInt128) << (16 * (7 - i)))
      Next ' i
      Return uxlReturn
    End Get
  End Property

  ' Calculate subnet IP address
  Public ReadOnly Property Subnet As String
    Get ' Numeric IP ANDed with NetMask gives Subnet address
      ' Output new subnet IP address
      Return CalcIP(Numeric And NetMask, True)
    End Get
  End Property

  ' Mask Converts XX netmask to 128 bit number
  Public ReadOnly Property NetMask As UInt128
    Get
      Dim uxlMask As UInt128 = 0
      ' Loop
      For i = 1 To 128
        ' Shift left 1
        uxlMask <<= 1
        ' Add 1 if appropriate
        If i <= _mask Then uxlMask += 1
      Next
      Return uxlMask
    End Get
  End Property

  Public ReadOnly Property Hextet(Chosen As Integer) As String
    ' This function will not assume that IPAddress is formatted properly X.X.X.X
    Get
      ' Split IP into hextets
      Dim Hextets() As String = _ip.Split(":"c)
      ' make 8 hextets
      If Hextets.GetUpperBound(0) < 7 Then
        ' Save current upper bound
        Dim intBound = Hextets.GetUpperBound(0)
        ' Initialize counter
        Dim intCount = 0
        ' Count hextets up to blank
        Do Until Hextets(intCount) = "" OrElse intCount = intBound
          ' Increment counter
          intCount += 1
        Loop
        ' Empty hextet at beginning, end, or middle
        ' Remove empty hextet
        Remove(Hextets, intCount)
        ' Begining and ending :: will leave 2 empty hextets
        ' in the array.
        If intCount = Hextets.Length Then
          ' Decrement counter so it is not pointing out of bounds
          intCount -= 1
          ' Remove extra hextet
          Remove(Hextets, intCount)
        ElseIf intCount = 0 Then
          ' Remove extra hextet
          Remove(Hextets, 0)
        End If
        ' Expand to 8 elements
        Do While Hextets.Length < 8
          Insert(Hextets, intCount)
          ' Set value to 0 so integer conversion will work
          Hextets(intCount) = "0"
        Loop
      End If
      ' Check for blanks (This can happen here if a single 0 hetxet
      ' is replaced with ::
      For i = 0 To 7
        If Hextets(i) = "" Then Hextets(i) = "0"
      Next
      ' Return selected hextet 
      Return Hextets(Chosen - 1)
    End Get
  End Property

  ' ToBinary converts 128 bits to binary string
  ' Converts IP address by default. Returns netmask if ConvertMask is true
  Public ReadOnly Property ToBinary(Optional ConvertMask As Boolean = False) As String
    Get
      ' Output string variable
      Dim strOut As String = ""
      ' Array of numbers, either IP address or Netmask
      Dim uxlNumber As UInt128 = If(ConvertMask, NetMask, Numeric)
      ' 128 bits per number
      For i = 1 To 128
        ' Add a 1 or 0 to the string
        strOut += If(uxlNumber.Bit(128 - i) = 1, "1", "0")
      Next ' i
      Return strOut
    End Get
  End Property

  ' Check IPv6 address for validity
  Public ReadOnly Property Valid As Boolean
    Get
      ' Valid hex characters plus colon
      Const ValidChars As String = "0123456789ABCDEF:"
      ' Check for invalid characters
      For i = 0 To _ip.Length - 1
        If Not ValidChars.Contains(_ip.ToUpper.Substring(i, 1)) Then Return False
      Next
      ' Split IP into hextets
      Dim strHex() As String = _ip.Split(":"c)
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
        End If
      Next
      ' Netmask out of range
      If _mask > 128 Then Return False
      ' Passed all tests
      Return True
    End Get
  End Property

  ' Calculate Summary IP address and keep it
  Public Sub Summarize(ByRef IPInput As IPv6Address)
    ' If the values are the same, don't bother
    If Not IPInput.Numeric = Numeric() Then
      Dim uxlNumX, uxlNumA As UInt128
      uxlNumX = Not (Numeric Xor IPInput.Numeric)
      uxlNumA = NetMask And Numeric And IPInput.Numeric
      Dim IPXOr As New IPv6Address(CalcIP(uxlNumX), 0)
      _mask = CUInt(IPXOr.ToBinary.IndexOf("0") - 1)
      _ip = CalcIP(uxlNumA, True)
    End If
  End Sub

  ' CalcIP Converts 128 bits to string IP address
  ' Array must be 4 long
  Public Function CalcIP(IPNumber As UInt128, Optional Compressed As Boolean = False) As String
    ' IPNumber is a 4 element array. Runtime error will result if not.
    ' Compressed, if missing or False, will result in a fully expanded address.
    '             if True, a fully compressed address will be returned
    Dim IPAddr As String = ""
    For i = 0 To 7 ' Each of the 4 Uint array elements
      If i > 0 Then IPAddr += ":" ' Seperate each hextet
      ' UInt128 variable into 8 16 bit values
      IPAddr += Hex((IPNumber >> (16 * (7 - i))).Lo And &HFFFFUL).PadLeft(4, "0"c)
    Next
    ' Compress it if requested
    If Compressed Then IPAddr = Compress(IPAddr)
    ' Return IP string
    Return IPAddr
  End Function

  ' Compress IPv6 address
  Public Function Compress(IPaddr As String) As String
    ' Don't do all the work if it's not a valid IP address
    If Not Valid Then Return "::"
    ' Array of hex numbers - Split address at colons
    Dim strHextets() As String = Expand(IPaddr).Split(":"c)
    ' Used Expand to make fully expanded IP address (8 hextets, 4 characters wide)
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
    For i = 0 To 7
      ' Look for "0" hextets
      If Convert.ToInt32(strHextets(i), 16) = 0 Then
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
    ' See if there are any cosecutive "0" hextets
    If intLongest > 1 Then
      strHextets(intNullHex) = ""
      If intNullHex + intLongest = 8 Then
        ' Blank final hextet
        strHextets(7) = ""
      End If
      ' Squeeze together
      For i = 1 To intLongest - 1
        Remove(strHextets, intNullHex + 1)
      Next
      ' Shorten to compressed length
      ReDim Preserve strHextets(7 - (intLongest - 1))
    ElseIf intLongest = 1 Then
      strHextets(intNullHex) = ""
    End If
    ' Cycle through remaining hextets
    For i = 0 To strHextets.GetUpperBound(0)
      ' Trim leading 0s
      If strHextets(i).Length > 0 Then
        strHextets(i) = strHextets(i).Trim("0"c)
        If strHextets(i).Length = 0 Then strHextets(i) = "0"
      End If
      ' Add current hextet to string
      strResult += strHextets(i) + ":"
    Next
    ' Remove extra colon
    strResult = strResult.Remove(strResult.Length - 1, 1)
    If strHextets(0) = "" Then strResult = ":" + strResult
    If strHextets(strHextets.GetUpperBound(0)) = "" Then strResult += ":"
    ' Shave off extra 0 at beginning or end
    If strResult.Length > 2 AndAlso strResult.Substring(0, 3) = "0::" Then
      strResult = strResult.Remove(0, 1)
    ElseIf strResult.Length > 2 AndAlso strResult.Substring(strResult.Length - 3, 3) = "::0" Then
      strResult = strResult.Remove(strResult.Length - 1, 1)
    End If
    ' Return result
    Return strResult
  End Function

  ' Expand IPv6 address
  Public Function Expand(IPAddr As String) As String
    ' Temporary
    Dim tmpIPv6 = New IPv6Address(IPAddr, 0)
    ' Return String
    Dim strReturn As String = ""
    For i = 0 To 7
      ' Add colon before all but the first hextet
      If i > 0 Then strReturn += ":"
      ' Add current hextet to string
      strReturn += tmpIPv6.Hextet(i + 1).PadLeft(4, "0"c)
    Next
    ' Return IP Address
    Return strReturn
  End Function

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
End Class
