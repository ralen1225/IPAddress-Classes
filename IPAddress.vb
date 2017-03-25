Option Explicit On
Option Strict On

Public Class IPAddress
  ' IP and mask values
  Private _ip As UInteger
  Private _mask As UInteger

  ' IP property set and get to insulate private variable
  Public Property IP As UInteger
    Get
      Return _ip
    End Get
    Set(value As UInteger)
      _ip = value
    End Set
  End Property

  ' Mask property set and get to insulate private variable
  Public Property Mask As UInteger
    Get
      Return _mask
    End Get
    Set(value As UInteger)
      If value <= 32 Then
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
  Public Sub New(ByRef IPAdd As IPAddress)
    _ip = IPAdd.IP
    _mask = IPAdd.Mask
  End Sub

  ' Inputs integer and integer
  Public Sub New(ByVal IPAdd As UInteger, ByVal IPMask As UInteger)
    _ip = IPAdd
    ' Mask property does range checking
    Mask = IPMask
  End Sub

  ' Inputs string Address and Uint net prefix
  Public Sub New(ByVal IPAdd As String, ByVal IPMask As UInteger)
    _ip = StrToIP(IPAdd)
    ' Mask property does range checking
    Mask = IPMask
  End Sub

  ' Calculate subnet IP address
  Public ReadOnly Property Subnet As UInteger
    Get
      ' Return new subnet IP address
      Return _ip And NetMask
    End Get
  End Property

  ' Calculate the subnet broadcast address
  Public ReadOnly Property Broadcast As UInteger
    Get ' Numeric IP ORed with Wildcard mask (not netmask) gives broadcast
      Return _ip Or Not NetMask
    End Get
  End Property

  ' Mask Converts XX netmask to 32 bits
  Public Property NetMask As UInteger
    Set(value As UInteger)
      ' Check for the 2 easy options
      If value = UInteger.MaxValue Then
        _mask = 32
      ElseIf value = UInteger.MinValue Then
        _mask = 0
      Else ' Count bits until the first 0
        Dim intCount As UInteger = 0
        Do While (value << CInt(intCount)) >= (1UI << 31)
          intCount += 1UI
        Loop
        _mask = intCount
      End If
    End Set
    Get
      Dim intMask As UInteger = 0
      ' Loop
      For i = 1 To CInt(_mask)
        ' Shift left 1 and add 1
        intMask = (intMask << 1) + 1UI
      Next
      intMask <<= CInt(32 - _mask)
      Return intMask
    End Get
  End Property

  Public Property Octet(Index As Integer) As UInteger
    ' Index range 1-4
    Set(value As UInteger)
      If Index >= 1 AndAlso Index <= 4 Then
        If value <= 255 Then
          ' Clear octet bits and set them to new value - Cut value off at 255
          _ip = (_ip And Not (&HFFUI << (8 * (4 - Index)))) Or ((value And &HFFUI) << (8 * (4 - Index)))
        Else
          Throw New ArgumentOutOfRangeException
        End If
      Else
        Throw New IndexOutOfRangeException
      End If
    End Set
    Get
      ' Return value
      If Index >= 1 AndAlso Index <= 4 Then
        Return (_ip >> (8 * (4 - Index))) And &HFFUI
      Else
        Throw New IndexOutOfRangeException
      End If
    End Get
  End Property

  ' ToBinary converts 32 bits to binary string
  Public ReadOnly Property ToBinary(Optional ConvertMask As Boolean = False) As String
    Get
      ' Output string
      Dim strOut As String = ""
      ' Which number are we working on?
      Dim intNumber As UInteger = If(ConvertMask, NetMask, _ip)
      ' Loop for 32 bits
      For i = 0 To 31
        ' Add 1 for on bit, 0 for off bit
        strOut += If((intNumber And (1UI << (31 - i))) <> 0, "1", "0")
      Next
      ' Return result
      Return strOut
    End Get
  End Property

  ' CalcIP Converts 32 bits to string IP address
  Public Property IPString As String
    Set(value As String)
      If Valid(value) Then
        _ip = StrToIP(value)
      Else
        Throw New InvalidCastException
      End If
    End Set
    Get
      ' Empty string
      Dim strIPAddr As String = ""
      ' Generate 4 octets
      For i = 0 To 3
        ' After the first one, add a dot
        If i > 0 Then strIPAddr += "."
        ' Isolate current octet
        strIPAddr += ((_ip >> ((3 - i) * 8)) And &HFFUI).ToString
      Next
      ' Return IP string
      Return strIPAddr
    End Get
  End Property

  ' Test to see if this is a valid IP address
  Public Shared Function Valid(IPAddress As String) As Boolean
    ' Valid character list
    Dim strValidChars As String = "0123456789."
    ' Check length
    If IPAddress.Length = 0 Then Return False
    ' Check each character
    For i = 0 To IPAddress.Length - 1
      If Not strValidChars.Contains(IPAddress.Substring(i, 1)) Then Return False
    Next
    ' Split string into octets
    Dim strOctets() As String = IPAddress.Split("."c)
    ' If not 4 octets, it fails
    If strOctets.Length <> 4 Then Return False
    ' Check each octet
    For Each strOctet In strOctets
      ' Check length and value - 0-length strings will not convert to UInt
      If strOctet.Length = 0 OrElse CUInt(strOctet) > 255 Then Return False
    Next
    ' Passed all tests
    Return True
  End Function

  ' Calculate Summary IP address and keep it
  Public Sub Summarize(ByRef IPInput As IPAddress)
    ' If either mask is zero, the summary is the entire internet
    If _mask = 0 OrElse IPInput.Mask = 0 Then
      _ip = 0
      _mask = 0
      ' If values match, don't do anything
    ElseIf Not IPInput.IP = _ip Then
      ' Locate first difference
      Dim IPXOr As New IPAddress(Not (_ip Xor IPInput.IP), 0)
      ' Set new mask
      _mask = CUInt(IPXOr.ToBinary.IndexOf("0"))
      ' Calculate new IP
      _ip = NetMask And _ip And IPInput.IP
    End If
  End Sub

  Public Shared Function StrToIP(IPAdd As String) As UInteger
    Dim Octets() As String = IPAdd.Split("."c)
    Dim IntIP As UInteger = 0
    If Octets.Length = 4 Then
      For i = 0 To 3
        Dim intOctet As UInteger
        If (Not UInteger.TryParse(Octets(i), intOctet)) OrElse intOctet > 255 Then Throw New InvalidCastException
        IntIP = IntIP Or intOctet << ((3 - i) * 8)
      Next
      Return IntIP
    Else
      Throw New InvalidCastException
    End If
  End Function
End Class
